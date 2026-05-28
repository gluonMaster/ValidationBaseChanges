"""Pending snapshot v1/v2 helpers.

Snapshot v1 is the legacy flat dict with tracked fields at the top level.
Snapshot v2 keeps that shape for compatibility and adds normalized reserved
keys for non-tracked payloads and projected timeline context.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Iterable

from django.core.exceptions import FieldDoesNotExist
from django.db import models

from apps.karteien.models import TRACKED_FIELDS, KarteiRecord


SNAPSHOT_VERSION_V2 = 2

SNAPSHOT_VERSION_KEY = "_snapshot_version"
PENDING_ACTION_KEY = "_pending_action"
PENDING_NONTRACKED_PAYLOAD_KEY = "_pending_nontracked_payload"
ROLLBACK_NONTRACKED_PAYLOAD_KEY = "_rollback_nontracked_payload"
PENDING_CONTRACT_TYPE_ENTRY_KEY = "_pending_contract_type_entry"
PENDING_CONTRACT_STATUS_ENTRY_KEY = "_pending_contract_status_entry"
PENDING_META_KEY = "_pending_meta"

RESERVED_SNAPSHOT_KEYS: tuple[str, ...] = (
    SNAPSHOT_VERSION_KEY,
    PENDING_ACTION_KEY,
    PENDING_NONTRACKED_PAYLOAD_KEY,
    ROLLBACK_NONTRACKED_PAYLOAD_KEY,
    PENDING_CONTRACT_TYPE_ENTRY_KEY,
    PENDING_CONTRACT_STATUS_ENTRY_KEY,
    PENDING_META_KEY,
)

PENDING_ACTIONS: tuple[str, ...] = (
    "FORM_SAVE",
    "MONTHS_OVERRIDE",
    "APPLY_CATEGORY",
    "PRICE_OPTION_RECALC",
    "LEGACY_TO_AUTO",
    "CONTRACT_TYPE_CHANGE",
    "CONTRACT_STATUS_CHANGE",
    "QUICK_SET_SUBJECT_REF",
    "GROUP_PREPARE_LEGACY",
)

PENDING_NONTRACKED_FIELDS: tuple[str, ...] = (
    "subject1_ref_id",
    "teacher1_ref_id",
    "price1_ref_id",
    "start_month_1",
    "end_month_1",
    "months_csv_1",
    "subject2_ref_id",
    "teacher2_ref_id",
    "price2_ref_id",
    "start_month_2",
    "end_month_2",
    "months_csv_2",
    "sepa_marker",
    "months_mode",
    "base_amounts",
    "hours_amounts",
    "legacy_base_amounts_enabled",
    "discounts_disabled",
    "discounts_disabled_months",
    "is_monthly_contract",
    "contract_type_raw",
    "is_contract_terminated",
    "contract_status_raw",
    "contract_terminated_from_month",
)


class SnapshotValidationError(ValueError):
    """Raised when building a snapshot v2 with invalid reserved payloads."""


@dataclass(frozen=True)
class SnapshotTimelineContext:
    """Projected timeline metadata that is not stored on the record instance."""

    contract_type_entry: dict[str, Any] | None = None
    contract_status_entry: dict[str, Any] | None = None


@dataclass(frozen=True)
class SnapshotProjection:
    """Projected pending/declined state for display or approval handling."""

    record: KarteiRecord
    timeline: SnapshotTimelineContext
    pending_action: str | None
    pending_meta: dict[str, Any]


def snapshot_dict(snapshot: Any) -> dict[str, Any]:
    """Return a shallow copy of a JSON snapshot if it is a dict."""

    return dict(snapshot) if isinstance(snapshot, dict) else {}


def is_snapshot_v2(snapshot: Any) -> bool:
    """Return whether the snapshot declares the v2 format."""

    return snapshot_dict(snapshot).get(SNAPSHOT_VERSION_KEY) == SNAPSHOT_VERSION_V2


def _json_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): _json_value(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_value(item) for item in value]
    return value


def build_tracked_snapshot(record: KarteiRecord) -> dict[str, Any]:
    """Build a legacy-compatible snapshot of tracked fields."""

    return {
        field_name: _json_value(getattr(record, field_name, None))
        for field_name in TRACKED_FIELDS
    }


def _field_for_name(record: KarteiRecord, field_name: str):
    try:
        return record._meta.get_field(field_name)
    except FieldDoesNotExist:
        return None


def _coerce_model_value(record: KarteiRecord, field_name: str, value: Any) -> Any:
    if value in ("", None):
        return None if _field_for_name(record, field_name) and getattr(
            _field_for_name(record, field_name), "null", False
        ) else value

    field = _field_for_name(record, field_name)
    if isinstance(field, models.DecimalField):
        return Decimal(str(value))
    if isinstance(field, models.DateField):
        if isinstance(value, date):
            return value
        return date.fromisoformat(str(value))
    if isinstance(field, models.BooleanField):
        if isinstance(value, str):
            return value.lower() in ("1", "true", "yes", "on")
        return bool(value)
    if isinstance(field, (models.PositiveSmallIntegerField, models.PositiveIntegerField, models.IntegerField)):
        return int(value) if value is not None else None
    return value


def apply_tracked_snapshot_to_record(
    record: KarteiRecord,
    snapshot: Any,
) -> KarteiRecord:
    """Apply only top-level tracked snapshot values to a record in memory."""

    data = snapshot_dict(snapshot)
    for field_name in TRACKED_FIELDS:
        if field_name not in data:
            continue
        setattr(record, field_name, _coerce_model_value(record, field_name, data[field_name]))
    return record


def build_nontracked_payload(
    record: KarteiRecord,
    fields: Iterable[str] = PENDING_NONTRACKED_FIELDS,
) -> dict[str, Any]:
    """Build a JSON-serializable non-tracked payload from a record."""

    payload: dict[str, Any] = {}
    for field_name in fields:
        payload[field_name] = _json_value(getattr(record, field_name, None))
    return payload


def build_pending_nontracked_payload(
    record: KarteiRecord,
    fields: Iterable[str] = PENDING_NONTRACKED_FIELDS,
) -> dict[str, Any]:
    """Build `_pending_nontracked_payload` for the frozen proposed state."""

    return build_nontracked_payload(record, fields=fields)


def build_rollback_nontracked_payload(
    record: KarteiRecord,
    fields: Iterable[str] = PENDING_NONTRACKED_FIELDS,
) -> dict[str, Any]:
    """Build `_rollback_nontracked_payload` for transitional prewrite paths."""

    return build_nontracked_payload(record, fields=fields)


def _normalize_dict_payload(value: Any, *, field_name: str) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise SnapshotValidationError(f"{field_name} must be a JSON object")
    return {str(key): _json_value(val) for key, val in value.items()}


def normalize_pending_nontracked_payload(payload: Any) -> dict[str, Any]:
    """Normalize `_pending_nontracked_payload` for v2 snapshot storage."""

    return _normalize_dict_payload(
        payload,
        field_name=PENDING_NONTRACKED_PAYLOAD_KEY,
    )


def normalize_rollback_nontracked_payload(payload: Any) -> dict[str, Any] | None:
    """Normalize `_rollback_nontracked_payload` for v2 snapshot storage."""

    if payload is None:
        return None
    return _normalize_dict_payload(
        payload,
        field_name=ROLLBACK_NONTRACKED_PAYLOAD_KEY,
    )


def get_pending_nontracked_payload(snapshot: Any) -> dict[str, Any]:
    """Read `_pending_nontracked_payload` from v2 snapshots."""

    data = snapshot_dict(snapshot)
    payload = data.get(PENDING_NONTRACKED_PAYLOAD_KEY)
    return dict(payload) if isinstance(payload, dict) else {}


def get_rollback_nontracked_payload(snapshot: Any) -> dict[str, Any] | None:
    """Read `_rollback_nontracked_payload` from v2 snapshots."""

    data = snapshot_dict(snapshot)
    payload = data.get(ROLLBACK_NONTRACKED_PAYLOAD_KEY)
    return dict(payload) if isinstance(payload, dict) else None


def _normalize_timeline_entry(value: Any, *, field_name: str) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        raise SnapshotValidationError(f"{field_name} must be a JSON object or null")
    entry = {str(key): _json_value(val) for key, val in value.items()}
    raw_month = entry.get("effective_from_month")
    try:
        month = int(raw_month)
    except (TypeError, ValueError):
        raise SnapshotValidationError(
            f"{field_name}.effective_from_month must be an integer 1..12"
        ) from None
    if not 1 <= month <= 12:
        raise SnapshotValidationError(
            f"{field_name}.effective_from_month must be an integer 1..12"
        )
    entry["effective_from_month"] = month
    return entry


def normalize_pending_contract_type_entry(value: Any) -> dict[str, Any] | None:
    """Normalize `_pending_contract_type_entry`."""

    return _normalize_timeline_entry(
        value,
        field_name=PENDING_CONTRACT_TYPE_ENTRY_KEY,
    )


def normalize_pending_contract_status_entry(value: Any) -> dict[str, Any] | None:
    """Normalize `_pending_contract_status_entry`."""

    return _normalize_timeline_entry(
        value,
        field_name=PENDING_CONTRACT_STATUS_ENTRY_KEY,
    )


def get_pending_contract_type_entry(snapshot: Any) -> dict[str, Any] | None:
    """Read contract-type timeline proposal from v1/v2 snapshots."""

    value = snapshot_dict(snapshot).get(PENDING_CONTRACT_TYPE_ENTRY_KEY)
    return dict(value) if isinstance(value, dict) else None


def get_pending_contract_status_entry(snapshot: Any) -> dict[str, Any] | None:
    """Read contract-status timeline proposal from v1/v2 snapshots."""

    value = snapshot_dict(snapshot).get(PENDING_CONTRACT_STATUS_ENTRY_KEY)
    return dict(value) if isinstance(value, dict) else None


def normalize_pending_meta(meta: Any) -> dict[str, Any]:
    """Normalize `_pending_meta` and keep unknown future-compatible keys."""

    if meta is None:
        meta = {}
    if not isinstance(meta, dict):
        raise SnapshotValidationError(f"{PENDING_META_KEY} must be a JSON object")

    normalized = {str(key): _json_value(value) for key, value in meta.items()}
    raw_touched = normalized.get("touched_months") or []
    touched_months: list[int] = []
    for raw_month in raw_touched:
        try:
            month = int(raw_month)
        except (TypeError, ValueError):
            continue
        if 1 <= month <= 12 and month not in touched_months:
            touched_months.append(month)

    raw_warnings = normalized.get("warnings") or []
    warnings = [str(item) for item in raw_warnings] if isinstance(raw_warnings, list) else [str(raw_warnings)]

    normalized["touched_months"] = sorted(touched_months)
    normalized["warnings"] = warnings
    normalized.setdefault("audit_summary", "")
    normalized.setdefault("ui_summary", "")
    return normalized


def get_pending_meta(snapshot: Any) -> dict[str, Any]:
    """Read `_pending_meta` from v2 snapshots using tolerant defaults."""

    try:
        return normalize_pending_meta(snapshot_dict(snapshot).get(PENDING_META_KEY))
    except SnapshotValidationError:
        return normalize_pending_meta({})


def normalize_reserved_snapshot_keys(snapshot: dict[str, Any]) -> dict[str, Any]:
    """Normalize and validate all v2 reserved keys in a snapshot dict."""

    version = snapshot.get(SNAPSHOT_VERSION_KEY)
    if version != SNAPSHOT_VERSION_V2:
        raise SnapshotValidationError(
            f"{SNAPSHOT_VERSION_KEY} must be {SNAPSHOT_VERSION_V2!r}"
        )

    action = snapshot.get(PENDING_ACTION_KEY)
    if action not in PENDING_ACTIONS:
        raise SnapshotValidationError(
            f"{PENDING_ACTION_KEY} must be one of {', '.join(PENDING_ACTIONS)}"
        )

    snapshot[PENDING_NONTRACKED_PAYLOAD_KEY] = normalize_pending_nontracked_payload(
        snapshot.get(PENDING_NONTRACKED_PAYLOAD_KEY)
    )
    snapshot[ROLLBACK_NONTRACKED_PAYLOAD_KEY] = normalize_rollback_nontracked_payload(
        snapshot.get(ROLLBACK_NONTRACKED_PAYLOAD_KEY)
    )
    snapshot[PENDING_CONTRACT_TYPE_ENTRY_KEY] = normalize_pending_contract_type_entry(
        snapshot.get(PENDING_CONTRACT_TYPE_ENTRY_KEY)
    )
    snapshot[PENDING_CONTRACT_STATUS_ENTRY_KEY] = normalize_pending_contract_status_entry(
        snapshot.get(PENDING_CONTRACT_STATUS_ENTRY_KEY)
    )
    snapshot[PENDING_META_KEY] = normalize_pending_meta(snapshot.get(PENDING_META_KEY))
    return snapshot


def build_snapshot_v2(
    record: KarteiRecord,
    *,
    pending_action: str,
    pending_nontracked_payload: dict[str, Any] | None = None,
    rollback_nontracked_payload: dict[str, Any] | None = None,
    pending_contract_type_entry: dict[str, Any] | None = None,
    pending_contract_status_entry: dict[str, Any] | None = None,
    pending_meta: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a normalized snapshot v2 using one centralized serializer."""

    snapshot = build_tracked_snapshot(record)
    snapshot.update(
        {
            SNAPSHOT_VERSION_KEY: SNAPSHOT_VERSION_V2,
            PENDING_ACTION_KEY: pending_action,
            PENDING_NONTRACKED_PAYLOAD_KEY: pending_nontracked_payload or {},
            ROLLBACK_NONTRACKED_PAYLOAD_KEY: rollback_nontracked_payload,
            PENDING_CONTRACT_TYPE_ENTRY_KEY: pending_contract_type_entry,
            PENDING_CONTRACT_STATUS_ENTRY_KEY: pending_contract_status_entry,
            PENDING_META_KEY: pending_meta or {},
        }
    )
    return normalize_reserved_snapshot_keys(snapshot)


def apply_nontracked_payload_to_record(
    record: KarteiRecord,
    payload: Any,
) -> KarteiRecord:
    """Apply non-tracked payload fields to a record in memory."""

    if not isinstance(payload, dict):
        return record

    for field_name, value in payload.items():
        if field_name in RESERVED_SNAPSHOT_KEYS or field_name in TRACKED_FIELDS:
            continue
        setattr(record, field_name, _coerce_model_value(record, field_name, value))
    return record


def apply_pending_snapshot_to_record(
    record: KarteiRecord,
    snapshot: Any,
) -> KarteiRecord:
    """Apply tracked fields plus v2 pending non-tracked payload in memory."""

    apply_tracked_snapshot_to_record(record, snapshot)
    if is_snapshot_v2(snapshot):
        apply_nontracked_payload_to_record(
            record,
            get_pending_nontracked_payload(snapshot),
        )
    return record


def clone_record_for_projection(record: KarteiRecord) -> KarteiRecord:
    """Clone concrete model fields into an unsaved in-memory instance."""

    clone = KarteiRecord()
    for field in record._meta.fields:
        setattr(clone, field.name, getattr(record, field.name))
    return clone


def build_projected_record_from_snapshot(
    record: KarteiRecord,
    snapshot: Any,
) -> SnapshotProjection:
    """Project pending/declined snapshot state on top of the live record.

    The projection follows the v2 contract:
    1. live record is the baseline;
    2. top-level tracked snapshot values are applied in memory;
    3. v2 `_pending_nontracked_payload` is applied in memory;
    4. timeline proposal entries are returned separately from the record.
    """

    projected = clone_record_for_projection(record)
    apply_pending_snapshot_to_record(projected, snapshot)
    return SnapshotProjection(
        record=projected,
        timeline=SnapshotTimelineContext(
            contract_type_entry=get_pending_contract_type_entry(snapshot),
            contract_status_entry=get_pending_contract_status_entry(snapshot),
        ),
        pending_action=snapshot_dict(snapshot).get(PENDING_ACTION_KEY),
        pending_meta=get_pending_meta(snapshot),
    )


def values_equal(value1: Any, value2: Any) -> bool:
    """Loose value comparison for projected display diffs."""

    def normalize(value: Any) -> Any:
        if isinstance(value, Decimal):
            return value.normalize() if value else None
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        if isinstance(value, str):
            stripped = value.strip()
            return stripped if stripped else None
        return value

    return normalize(value1) == normalize(value2)


def changed_fields_between_records(
    original: KarteiRecord,
    projected: KarteiRecord,
    fields: Iterable[str],
) -> set[str]:
    """Return field names that differ between live and projected records."""

    changed: set[str] = set()
    for field_name in fields:
        if not values_equal(
            getattr(original, field_name, None),
            getattr(projected, field_name, None),
        ):
            changed.add(field_name)
    return changed


def comparison_fields_for_snapshot(snapshot: Any) -> list[str]:
    """Return tracked fields plus v2 non-tracked payload keys for UI diffs."""

    fields = list(TRACKED_FIELDS)
    if is_snapshot_v2(snapshot):
        for field_name in get_pending_nontracked_payload(snapshot):
            if field_name not in fields:
                fields.append(field_name)
    return fields

