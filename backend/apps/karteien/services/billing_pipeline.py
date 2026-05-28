"""Shared builders for billing-affecting approval proposals.

The builders in this module are intentionally side-effect free: they mutate
only in-memory record clones and return a frozen proposal containing a
centralized snapshot v2 payload ready for PendingChange storage.
"""

from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Mapping

from apps.karteien.billing import (
    ZERO,
    CalculationFlags,
    build_base_amounts,
    calculate_month_values,
    get_semester_month_ranges,
    recalculate_record_months,
)
from apps.karteien.category_price import apply_category_price_to_record
from apps.karteien.models import (
    ContractStatusEntry,
    ContractStatusKind,
    KarteiRecord,
    MonthsMode,
)
from apps.karteien.services.pending_snapshot import (
    build_pending_nontracked_payload,
    build_snapshot_v2,
    clone_record_for_projection,
)


MONTH_NUMBERS: tuple[int, ...] = tuple(range(1, 13))
MONTH_FIELDS: tuple[str, ...] = tuple(f"month_{month}" for month in MONTH_NUMBERS)


@dataclass(frozen=True)
class BillingProposal:
    """Frozen billing proposal returned by pipeline builders."""

    action: str
    snapshot: dict[str, Any]
    touched_months: tuple[int, ...]
    audit_summary: str
    ui_summary: str
    warnings: tuple[str, ...] = ()
    diff: dict[str, Any] = field(default_factory=dict)


def _clone_for_proposal(record: KarteiRecord) -> KarteiRecord:
    """Clone a record and detach mutable JSON fields from the live instance."""

    clone = clone_record_for_projection(record)
    clone.base_amounts = deepcopy(getattr(record, "base_amounts", None) or {})
    clone.hours_amounts = deepcopy(getattr(record, "hours_amounts", None) or {})
    clone.discounts_disabled_months = deepcopy(
        getattr(record, "discounts_disabled_months", None) or []
    )
    return clone


def _normalize_months(months: Any) -> tuple[int, ...]:
    normalized: list[int] = []
    for raw_month in months or []:
        try:
            month = int(raw_month)
        except (TypeError, ValueError):
            continue
        if 1 <= month <= 12 and month not in normalized:
            normalized.append(month)
    return tuple(sorted(normalized))


def _month_range(start_month: int, end_month: int | None = None) -> tuple[int, ...]:
    end = 12 if end_month is None else end_month
    return _normalize_months(range(start_month, end + 1))


def _pending_meta(
    *,
    touched_months: tuple[int, ...],
    audit_summary: str,
    ui_summary: str,
    warnings: tuple[str, ...] = (),
) -> dict[str, Any]:
    return {
        "touched_months": list(touched_months),
        "warnings": list(warnings),
        "audit_summary": audit_summary,
        "ui_summary": ui_summary,
    }


def _pending_payload(
    record: KarteiRecord,
    overrides: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    payload = build_pending_nontracked_payload(record)
    if overrides:
        payload.update(dict(overrides))
    return payload


def _build_proposal(
    *,
    action: str,
    record: KarteiRecord,
    touched_months: tuple[int, ...],
    audit_summary: str,
    ui_summary: str,
    warnings: tuple[str, ...] = (),
    pending_contract_type_entry: dict[str, Any] | None = None,
    pending_contract_status_entry: dict[str, Any] | None = None,
    pending_payload_overrides: Mapping[str, Any] | None = None,
    diff: dict[str, Any] | None = None,
) -> BillingProposal:
    snapshot = build_snapshot_v2(
        record,
        pending_action=action,
        pending_nontracked_payload=_pending_payload(record, pending_payload_overrides),
        pending_contract_type_entry=pending_contract_type_entry,
        pending_contract_status_entry=pending_contract_status_entry,
        pending_meta=_pending_meta(
            touched_months=touched_months,
            audit_summary=audit_summary,
            ui_summary=ui_summary,
            warnings=warnings,
        ),
    )
    return BillingProposal(
        action=action,
        snapshot=snapshot,
        touched_months=touched_months,
        audit_summary=audit_summary,
        ui_summary=ui_summary,
        warnings=warnings,
        diff=diff or {},
    )


def _format_months(months: tuple[int, ...]) -> str:
    return ",".join(str(month) for month in months)


def _decimal_base_amounts(record: KarteiRecord) -> dict[str, Decimal]:
    base_amounts = getattr(record, "base_amounts", None) or {}
    result: dict[str, Decimal] = {}
    for field_name in MONTH_FIELDS:
        raw_value = base_amounts.get(field_name, ZERO)
        try:
            result[field_name] = Decimal(str(raw_value or ZERO))
        except (TypeError, ValueError):
            result[field_name] = ZERO
    return result


def _set_month_values(
    record: KarteiRecord,
    values: Mapping[str, Decimal],
    *,
    months: tuple[int, ...] = MONTH_NUMBERS,
) -> None:
    for month in months:
        field_name = f"month_{month}"
        setattr(record, field_name, values.get(field_name, ZERO))


def _warnings_from_flags(flags: CalculationFlags) -> tuple[str, ...]:
    warnings: list[str] = []
    if flags.percent_discount_exceeded:
        warnings.append(
            f"Percent discounts clamped from {flags.original_percent_sum} to 0.99."
        )
    if flags.clamped_to_zero_months:
        months = _format_months(_normalize_months(flags.clamped_to_zero_months))
        warnings.append(f"Month values clamped to zero: {months}.")
    return tuple(warnings)


def _timeline_rollout_warnings(
    record: KarteiRecord,
    *,
    contract_type_month: int | None = None,
    contract_status_month: int | None = None,
    contract_status_kind: str | None = None,
) -> tuple[str, ...]:
    warnings: list[str] = []
    if contract_type_month is not None and contract_type_month > 1:
        warnings.append(
            "Intra-year contract type approval is blocked until Prompt 09."
        )
    if contract_status_kind == ContractStatusKind.PAUSED:
        warnings.append("PAUSED status approval is blocked until Prompt 09.")
    elif (
        contract_status_kind == ContractStatusKind.ACTIVE
        and contract_status_month is not None
        and _has_earlier_nonactive_status(record, contract_status_month)
    ):
        warnings.append("ACTIVE reactivation approval is blocked until Prompt 09.")
    return tuple(warnings)


def _has_earlier_nonactive_status(record: KarteiRecord, month: int) -> bool:
    if (
        record.contract_terminated_from_month is not None
        and record.contract_terminated_from_month < month
    ):
        return True
    if not record.pk:
        return False
    return record.contract_status_entries.filter(
        effective_from_month__lt=month,
    ).exclude(kind=ContractStatusKind.ACTIVE).exists()


def _add_ov_marker(raw_value: str) -> str:
    raw_value = raw_value or ""
    if re.search(r"O/V", raw_value, re.IGNORECASE):
        return raw_value
    return f"{raw_value.strip()} O/V".strip()


def _remove_ov_marker(raw_value: str) -> str:
    raw_value = raw_value or ""
    result = re.sub(r"O/V", "", raw_value, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", result).strip()


def _add_kn_marker(raw_value: str) -> str:
    raw_value = raw_value or ""
    if re.search(r"(^|\s)KN(\s|$)", raw_value, re.IGNORECASE):
        return raw_value
    return f"{raw_value.strip()} KN".strip()


def _remove_kn_marker(raw_value: str) -> str:
    raw_value = raw_value or ""
    result = re.sub(r"(^|\s)KN(\s|$)", r"\1\2", raw_value, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", result).strip()


def _contract_type_label(is_monthly: bool) -> str:
    return "monthly" if is_monthly else "yearly"


def _user_id(user: Any) -> int | None:
    return getattr(user, "id", None)


def _recalc_scope_months(
    record: KarteiRecord,
    *,
    apply_from_month_1: int | None = None,
    apply_from_month_2: int | None = None,
    apply_to_month_1: int | None = None,
    apply_to_month_2: int | None = None,
) -> tuple[int, ...]:
    sem1_months, sem2_months = get_semester_month_ranges(record.year)
    scoped: list[int] = []

    def extend_scope(
        sem_months: list[int],
        apply_from: int | None,
        apply_to: int | None,
    ) -> None:
        if apply_from is None and apply_to is None:
            scoped.extend(sem_months)
            return
        start = apply_from if apply_from is not None else min(sem_months)
        end = apply_to if apply_to is not None else max(sem_months)
        scoped.extend(month for month in sem_months if start <= month <= end)

    if (
        apply_from_month_1 is not None
        or apply_to_month_1 is not None
        or (apply_from_month_2 is None and apply_to_month_2 is None)
    ):
        extend_scope(sem1_months, apply_from_month_1, apply_to_month_1)
    if (
        apply_from_month_2 is not None
        or apply_to_month_2 is not None
        or (apply_from_month_1 is None and apply_to_month_1 is None)
    ):
        extend_scope(sem2_months, apply_from_month_2, apply_to_month_2)

    return _normalize_months(scoped)


def build_apply_category_proposal(
    record: KarteiRecord,
    *,
    semester: int,
    from_month: int,
) -> BillingProposal:
    """Build an APPLY_CATEGORY proposal without writing to the live record."""

    if semester not in (1, 2):
        raise ValueError("semester must be 1 or 2")
    if not 1 <= from_month <= 12:
        raise ValueError("from_month must be 1..12")
    if record.months_mode == MonthsMode.LEGACY:
        raise ValueError("APPLY_CATEGORY requires AUTO or OVERRIDE mode.")

    proposed = _clone_for_proposal(record)
    old_mode = record.months_mode
    diff = apply_category_price_to_record(
        proposed,
        semester=semester,
        from_month=from_month,
    )
    months_updated = _normalize_months(diff["months_updated"])

    if old_mode == MonthsMode.OVERRIDE:
        proposed.months_mode = MonthsMode.AUTO

    mode_part = (
        f"; MODE[{old_mode}->{proposed.months_mode}]"
        if proposed.months_mode != old_mode
        else ""
    )
    audit_summary = f"APPLY_CATEGORY[months={_format_months(months_updated)}]{mode_part}"
    ui_summary = (
        f"Category price applied from month {from_month} "
        f"for semester {semester}."
    )
    return _build_proposal(
        action="APPLY_CATEGORY",
        record=proposed,
        touched_months=months_updated,
        audit_summary=audit_summary,
        ui_summary=ui_summary,
        diff=diff,
    )


def build_price_option_recalc_proposal(
    record: KarteiRecord,
    *,
    apply_from_month_1: int | None = None,
    apply_from_month_2: int | None = None,
    apply_to_month_1: int | None = None,
    apply_to_month_2: int | None = None,
    hours_amounts: dict[str, Decimal | str] | None = None,
) -> BillingProposal:
    """Build a PRICE_OPTION_RECALC proposal for AUTO/OVERRIDE records."""

    if record.months_mode == MonthsMode.LEGACY:
        raise ValueError("PRICE_OPTION_RECALC is not allowed for LEGACY records.")

    proposed = _clone_for_proposal(record)
    if hours_amounts is not None:
        proposed.hours_amounts = deepcopy(hours_amounts)

    old_mode = record.months_mode
    base_amounts = build_base_amounts(
        proposed,
        apply_from_month_1=apply_from_month_1,
        apply_from_month_2=apply_from_month_2,
        apply_to_month_1=apply_to_month_1,
        apply_to_month_2=apply_to_month_2,
        hours_amounts=hours_amounts,
    )
    proposed.base_amounts = {key: str(value) for key, value in base_amounts.items()}
    month_values, flags = calculate_month_values(
        proposed,
        base_amounts=base_amounts,
    )
    _set_month_values(proposed, month_values)

    if old_mode == MonthsMode.OVERRIDE:
        proposed.months_mode = MonthsMode.AUTO

    touched_months = _recalc_scope_months(
        record,
        apply_from_month_1=apply_from_month_1,
        apply_from_month_2=apply_from_month_2,
        apply_to_month_1=apply_to_month_1,
        apply_to_month_2=apply_to_month_2,
    )
    warnings = _warnings_from_flags(flags)
    mode_part = (
        f"; MODE[{old_mode}->{proposed.months_mode}]"
        if proposed.months_mode != old_mode
        else ""
    )
    audit_summary = (
        f"PRICE_OPTION_RECALC[months={_format_months(touched_months)}]"
        f"; BASE[months={_format_months(touched_months)}]"
        f"{mode_part}"
    )
    return _build_proposal(
        action="PRICE_OPTION_RECALC",
        record=proposed,
        touched_months=touched_months,
        audit_summary=audit_summary,
        ui_summary="Price-option recalculation submitted for approval.",
        warnings=warnings,
    )


def build_legacy_to_auto_proposal(
    record: KarteiRecord,
    *,
    hours_amounts: dict[str, Decimal | str] | None = None,
) -> BillingProposal:
    """Build an explicit full LEGACY_TO_AUTO proposal."""

    proposed = _clone_for_proposal(record)
    if hours_amounts is not None:
        proposed.hours_amounts = deepcopy(hours_amounts)

    old_mode = record.months_mode
    flags = recalculate_record_months(
        proposed,
        hours_amounts=hours_amounts,
    )
    proposed.months_mode = MonthsMode.AUTO
    proposed.legacy_base_amounts_enabled = False

    touched_months = MONTH_NUMBERS
    warnings = _warnings_from_flags(flags)
    audit_summary = (
        f"LEGACY_TO_AUTO[months={_format_months(touched_months)}]; "
        f"BASE[months={_format_months(touched_months)}]; "
        f"MODE[{old_mode}->{MonthsMode.AUTO}]"
    )
    return _build_proposal(
        action="LEGACY_TO_AUTO",
        record=proposed,
        touched_months=touched_months,
        audit_summary=audit_summary,
        ui_summary="Legacy record converted to AUTO for approval.",
        warnings=warnings,
    )


def build_months_override_proposal(
    record: KarteiRecord,
    *,
    month_changes: Mapping[str, Any],
    blocked_months: Mapping[int, str] | None = None,
) -> BillingProposal:
    """Build a MONTHS_OVERRIDE proposal with months_mode frozen as OVERRIDE."""

    proposed = _clone_for_proposal(record)
    old_mode = record.months_mode

    for field_name, value in month_changes.items():
        if field_name in MONTH_FIELDS:
            setattr(proposed, field_name, Decimal(str(value or ZERO)))

    for month in (blocked_months or {}):
        if 1 <= int(month) <= 12:
            setattr(proposed, f"month_{int(month)}", ZERO)

    proposed.months_mode = MonthsMode.OVERRIDE
    touched_months = MONTH_NUMBERS
    audit_summary = (
        f"MONTHS_OVERRIDE[months={_format_months(touched_months)}]; "
        f"MODE[{old_mode}->{MonthsMode.OVERRIDE}]"
    )
    return _build_proposal(
        action="MONTHS_OVERRIDE",
        record=proposed,
        touched_months=touched_months,
        audit_summary=audit_summary,
        ui_summary="Manual month override submitted for approval.",
    )


def build_contract_type_proposal(
    record: KarteiRecord,
    *,
    effective_from_month: int,
    is_monthly: bool,
    comment: str = "",
    user: Any = None,
) -> BillingProposal:
    """Build a CONTRACT_TYPE_CHANGE proposal without billing recalculation."""

    if not 1 <= effective_from_month <= 12:
        raise ValueError("effective_from_month must be 1..12")

    proposed = _clone_for_proposal(record)
    pending_payload_overrides: dict[str, Any] = {}
    if effective_from_month == 1:
        proposed.is_monthly_contract = bool(is_monthly)
        proposed.contract_type_raw = (
            _add_ov_marker(record.contract_type_raw)
            if is_monthly
            else _remove_ov_marker(record.contract_type_raw)
        )
        pending_payload_overrides = {
            "is_monthly_contract": proposed.is_monthly_contract,
            "contract_type_raw": proposed.contract_type_raw,
        }

    type_slug = _contract_type_label(is_monthly)
    type_label = "Monatsvertrag" if is_monthly else "Jahresvertrag"
    touched_months = _month_range(effective_from_month)
    warnings = _timeline_rollout_warnings(
        record,
        contract_type_month=effective_from_month,
    )
    return _build_proposal(
        action="CONTRACT_TYPE_CHANGE",
        record=proposed,
        touched_months=touched_months,
        audit_summary=f"TL[type@{effective_from_month}={type_slug}]",
        ui_summary=f"{type_label} ab Monat {effective_from_month}.",
        warnings=warnings,
        pending_contract_type_entry={
            "effective_from_month": effective_from_month,
            "is_monthly": bool(is_monthly),
            "comment": comment,
            "changed_by_id": _user_id(user),
        },
        pending_payload_overrides=pending_payload_overrides,
    )


def build_contract_status_proposal(
    record: KarteiRecord,
    *,
    effective_from_month: int,
    kind: str,
    comment: str = "",
    user: Any = None,
) -> BillingProposal:
    """Build a CONTRACT_STATUS_CHANGE proposal with frozen month effects."""

    if not 1 <= effective_from_month <= 12:
        raise ValueError("effective_from_month must be 1..12")
    allowed_kinds = {
        ContractStatusKind.ACTIVE,
        ContractStatusKind.PAUSED,
        ContractStatusKind.TERMINATED,
    }
    if kind not in allowed_kinds:
        raise ValueError("kind must be ACTIVE, PAUSED, or TERMINATED")

    proposed = _clone_for_proposal(record)
    from_month = effective_from_month
    future_entries = (
        list(
            record.contract_status_entries.filter(
                effective_from_month__gt=from_month,
            ).order_by("effective_from_month")
        )
        if record.pk
        else []
    )
    to_month = (
        future_entries[0].effective_from_month - 1 if future_entries else 12
    )
    touched_months = _month_range(from_month, to_month)

    pending_payload_overrides: dict[str, Any] = {}
    if to_month == 12:
        if kind == ContractStatusKind.TERMINATED:
            proposed.is_contract_terminated = True
            proposed.contract_status_raw = _add_kn_marker(record.contract_status_raw)
            proposed.contract_terminated_from_month = effective_from_month
            pending_payload_overrides = {
                "is_contract_terminated": True,
                "contract_status_raw": proposed.contract_status_raw,
                "contract_terminated_from_month": effective_from_month,
            }
        elif (
            kind == ContractStatusKind.ACTIVE
            and not _has_earlier_nonactive_status(record, effective_from_month)
        ):
            proposed.is_contract_terminated = False
            proposed.contract_status_raw = _remove_kn_marker(record.contract_status_raw)
            proposed.contract_terminated_from_month = None
            pending_payload_overrides = {
                "is_contract_terminated": False,
                "contract_status_raw": proposed.contract_status_raw,
                "contract_terminated_from_month": None,
            }

    month_summary = "unchanged"
    if kind in (ContractStatusKind.PAUSED, ContractStatusKind.TERMINATED):
        for month in touched_months:
            setattr(proposed, f"month_{month}", ZERO)
        month_summary = "zeroed"
    elif kind == ContractStatusKind.ACTIVE and record.months_mode == MonthsMode.AUTO:
        existing_entries = (
            list(
                record.contract_status_entries.exclude(
                    effective_from_month=effective_from_month,
                )
            )
            if record.pk
            else []
        )
        proposed_entry = ContractStatusEntry(
            record=record,
            effective_from_month=effective_from_month,
            kind=ContractStatusKind.ACTIVE,
        )
        month_values, _flags = calculate_month_values(
            proposed,
            base_amounts=_decimal_base_amounts(proposed),
            contract_status_entries=existing_entries + [proposed_entry],
        )
        _set_month_values(proposed, month_values, months=touched_months)
        month_summary = "restored"

    status_labels = {
        ContractStatusKind.ACTIVE: "Aktiv",
        ContractStatusKind.PAUSED: "Pausiert",
        ContractStatusKind.TERMINATED: "Gekuendigt",
    }
    warnings = _timeline_rollout_warnings(
        record,
        contract_status_month=effective_from_month,
        contract_status_kind=kind,
    )
    audit_summary = (
        f"TL[status@{effective_from_month}={kind}]; "
        f"MONTHS[{month_summary}={_format_months(touched_months)}]"
    )
    return _build_proposal(
        action="CONTRACT_STATUS_CHANGE",
        record=proposed,
        touched_months=touched_months,
        audit_summary=audit_summary,
        ui_summary=f"{status_labels.get(kind, kind)} ab Monat {effective_from_month}.",
        warnings=warnings,
        pending_contract_status_entry={
            "effective_from_month": effective_from_month,
            "kind": kind,
            "comment": comment,
            "changed_by_id": _user_id(user),
        },
        pending_payload_overrides=pending_payload_overrides,
    )


def build_quick_set_subject_ref_proposal(
    record: KarteiRecord,
    *,
    semester: int | None = None,
    semester_key: str | None = None,
    subject: Any,
) -> BillingProposal:
    """Build a QUICK_SET_SUBJECT_REF proposal for an unambiguous subject match."""

    if semester is None:
        if semester_key == "sem1":
            semester = 1
        elif semester_key == "sem2":
            semester = 2
    if semester not in (1, 2):
        raise ValueError("semester must be 1 or 2")

    proposed = _clone_for_proposal(record)
    ref_field = "subject1_ref" if semester == 1 else "subject2_ref"
    ref_id_field = f"{ref_field}_id"
    old_subject_id = getattr(record, ref_id_field)

    if hasattr(subject, "pk"):
        setattr(proposed, ref_field, subject)
        new_subject_id = subject.pk
        subject_name = getattr(subject, "name", str(subject))
    else:
        new_subject_id = int(subject)
        setattr(proposed, ref_id_field, new_subject_id)
        subject_name = str(subject)

    sem1_months, sem2_months = get_semester_month_ranges(record.year)
    touched_months = tuple(sem1_months if semester == 1 else sem2_months)
    sem_label = "1. HJ" if semester == 1 else "2. HJ"
    return _build_proposal(
        action="QUICK_SET_SUBJECT_REF",
        record=proposed,
        touched_months=touched_months,
        audit_summary=f"REF[{ref_id_field}={old_subject_id}->{new_subject_id}]",
        ui_summary=f"{sem_label}: {ref_field} matched to {subject_name}.",
    )
