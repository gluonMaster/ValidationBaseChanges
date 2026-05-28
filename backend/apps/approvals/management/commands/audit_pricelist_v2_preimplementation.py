from __future__ import annotations

import json
import re
from typing import Any

from django.core.management.base import BaseCommand, CommandParser
from django.db import transaction
from django.db.models import Count, Min

from apps.approvals.models import DeclinedChange, PendingChange
from apps.catalog.models import CategoryKind, SubjectCategory
from apps.karteien.models import (
    ContractStatusEntry,
    ContractTypeEntry,
    KarteiRecord,
    RecordStatus,
)


SNAPSHOT_VERSION_V2 = 2
KN_TOKEN_RE = re.compile(r"(^|\s)KN(\s|$)", re.IGNORECASE)


def _new_section() -> dict[str, Any]:
    return {"count": 0, "items": [], "truncated": 0}


def _add_item(section: dict[str, Any], item: dict[str, Any], limit: int) -> None:
    section["count"] += 1
    if limit == 0 or len(section["items"]) < limit:
        section["items"].append(item)
    else:
        section["truncated"] += 1


def _snapshot_dict(snapshot: Any) -> dict[str, Any]:
    return snapshot if isinstance(snapshot, dict) else {}


def _snapshot_v1_reason(snapshot: Any) -> str | None:
    if not isinstance(snapshot, dict):
        return "snapshot is not a JSON object"
    if snapshot.get("_snapshot_version") == SNAPSHOT_VERSION_V2:
        return None
    if "_snapshot_version" not in snapshot:
        return "missing _snapshot_version (legacy/v1)"
    return f"_snapshot_version={snapshot.get('_snapshot_version')!r} (expected 2)"


def _record_ref(record: KarteiRecord) -> dict[str, Any]:
    return {
        "record_pkid": record.pkid,
        "record_id": record.id,
        "year": record.year,
        "family_id": record.family_id,
        "child_name": record.child_name,
        "status": record.status or "NORMAL",
    }


def _change_ref(change: PendingChange | DeclinedChange) -> dict[str, Any]:
    data = _record_ref(change.record)
    data.update(
        {
            "change_id": change.id,
            "created_at": change.created_at.isoformat()
            if change.created_at
            else None,
            "updated_at": change.updated_at.isoformat()
            if change.updated_at
            else None,
        }
    )
    return data


def _json_equal(left: Any, right: Any) -> bool:
    return json.dumps(left, sort_keys=True, default=str) == json.dumps(
        right,
        sort_keys=True,
        default=str,
    )


def _has_nontracked_context(record: KarteiRecord) -> bool:
    return any(
        [
            record.subject1_ref_id,
            record.teacher1_ref_id,
            record.price1_ref_id,
            record.subject2_ref_id,
            record.teacher2_ref_id,
            record.price2_ref_id,
            record.sepa_marker,
            record.months_mode,
            bool(record.base_amounts),
            bool(record.hours_amounts),
            record.legacy_base_amounts_enabled,
            record.discounts_disabled,
            bool(record.discounts_disabled_months),
            record.is_monthly_contract,
            record.contract_type_raw,
            record.is_contract_terminated,
            record.contract_status_raw,
            record.contract_terminated_from_month is not None,
        ]
    )


def _pending_split_brain_reasons(
    pending: PendingChange,
) -> tuple[str, list[str]]:
    record = pending.record
    snapshot = _snapshot_dict(pending.snapshot)
    reasons: list[str] = []
    confidence = "possible"

    old_base_amounts = snapshot.get("_old_base_amounts")
    if old_base_amounts is not None:
        confidence = "high"
        if _json_equal(record.base_amounts or {}, old_base_amounts):
            reasons.append("legacy rollback marker _old_base_amounts is present")
        else:
            reasons.append(
                "live base_amounts differs from _old_base_amounts rollback marker"
            )

    comment = (pending.admin_comment or "").strip()
    if "Quick-set subject" in comment:
        confidence = "high"
        reasons.append("admin_comment identifies QuickSetSubjectRef live-FK path")
    if "LEGACY-Vorbereitung" in comment:
        confidence = "high"
        reasons.append("admin_comment identifies group legacy preparation path")

    if record.status != RecordStatus.PENDING:
        reasons.append(
            f"open PendingChange exists but live record status is {record.status or 'NORMAL'}"
        )

    if "_pending_nontracked_payload" not in snapshot:
        if _has_nontracked_context(record):
            reasons.append(
                "snapshot has no frozen nontracked payload while live record has "
                "billing/context fields"
            )
        elif reasons:
            reasons.append("snapshot has no frozen nontracked payload")

    if _snapshot_v1_reason(pending.snapshot):
        reasons.append(_snapshot_v1_reason(pending.snapshot))

    # Timeline-only v1 proposals are already reported in the v1 snapshot section.
    # Keep this split-brain section focused on rows where live state is suspect.
    timeline_only = (
        "_pending_contract_type_entry" in snapshot
        or "_pending_contract_status_entry" in snapshot
    ) and not (
        "_old_base_amounts" in snapshot
        or "Quick-set subject" in comment
        or "LEGACY-Vorbereitung" in comment
    )
    if timeline_only and confidence != "high":
        return "possible", []

    return confidence, reasons


def _declined_split_brain_reasons(
    record: KarteiRecord,
    latest_declined: DeclinedChange | None,
) -> list[str]:
    if latest_declined is None:
        return ["record status is DECLINED but no DeclinedChange exists"]

    snapshot = _snapshot_dict(latest_declined.snapshot)
    reasons: list[str] = []

    v1_reason = _snapshot_v1_reason(latest_declined.snapshot)
    if v1_reason:
        reasons.append(v1_reason)

    old_base_amounts = snapshot.get("_old_base_amounts")
    if old_base_amounts is not None:
        if _json_equal(record.base_amounts or {}, old_base_amounts):
            reasons.append("legacy rollback marker _old_base_amounts is present")
        else:
            reasons.append(
                "live base_amounts differs from _old_base_amounts rollback marker"
            )

    if "_pending_nontracked_payload" not in snapshot and _has_nontracked_context(
        record
    ):
        reasons.append(
            "declined snapshot has no frozen nontracked payload while live record "
            "has billing/context fields"
        )

    return reasons


def _build_report(*, database: str, limit: int) -> dict[str, Any]:
    report: dict[str, Any] = {
        "audit": "PRICELIST_V2_PRE_IMPLEMENTATION",
        "read_only": True,
        "database": database,
        "summary": {},
        "sections": {
            "pending_snapshot_v1": _new_section(),
            "declined_snapshot_v1": _new_section(),
            "legacy_pending_split_brain": _new_section(),
            "declined_split_brain_risk": _new_section(),
            "dirty_legacy_contract_rows": _new_section(),
            "partial_timeline_anomalies": _new_section(),
            "missing_group_nachhilfe_years": _new_section(),
        },
    }
    sections = report["sections"]

    pending_qs = (
        PendingChange.objects.using(database)
        .select_related("record")
        .order_by("id")
    )
    for pending in pending_qs.iterator():
        reason = _snapshot_v1_reason(pending.snapshot)
        if reason:
            item = _change_ref(pending)
            item.update(
                {
                    "is_processed": pending.is_processed,
                    "reason": reason,
                }
            )
            _add_item(sections["pending_snapshot_v1"], item, limit)

        if not pending.is_processed:
            confidence, reasons = _pending_split_brain_reasons(pending)
            if reasons:
                item = _change_ref(pending)
                item.update(
                    {
                        "confidence": confidence,
                        "reasons": reasons,
                    }
                )
                _add_item(sections["legacy_pending_split_brain"], item, limit)

    declined_qs = (
        DeclinedChange.objects.using(database)
        .select_related("record")
        .order_by("id")
    )
    for declined in declined_qs.iterator():
        reason = _snapshot_v1_reason(declined.snapshot)
        if reason:
            item = _change_ref(declined)
            item.update({"reason": reason})
            _add_item(sections["declined_snapshot_v1"], item, limit)

    latest_declined_by_record: dict[int, DeclinedChange] = {}
    latest_declined_qs = (
        DeclinedChange.objects.using(database)
        .filter(record__status=RecordStatus.DECLINED)
        .select_related("record")
        .order_by("record_id", "-created_at", "-id")
    )
    for declined in latest_declined_qs.iterator():
        latest_declined_by_record.setdefault(declined.record_id, declined)

    declined_records_qs = (
        KarteiRecord.objects.using(database)
        .filter(status=RecordStatus.DECLINED)
        .order_by("year", "id", "pkid")
    )
    for record in declined_records_qs.iterator():
        latest_declined = latest_declined_by_record.get(record.pkid)
        reasons = _declined_split_brain_reasons(record, latest_declined)
        if reasons:
            item = _record_ref(record)
            if latest_declined is not None:
                item["declined_change_id"] = latest_declined.id
            item["reasons"] = reasons
            _add_item(sections["declined_split_brain_risk"], item, limit)

    contract_records_qs = (
        KarteiRecord.objects.using(database)
        .only(
            "pkid",
            "id",
            "year",
            "family_id",
            "child_name",
            "status",
            "is_contract_terminated",
            "contract_status_raw",
            "contract_terminated_from_month",
        )
        .order_by("year", "id", "pkid")
    )
    for record in contract_records_qs.iterator():
        reasons = []
        month = record.contract_terminated_from_month
        if record.is_contract_terminated and month is None:
            reasons.append(
                "is_contract_terminated=True but contract_terminated_from_month is empty"
            )
        if KN_TOKEN_RE.search(record.contract_status_raw or "") and month is None:
            reasons.append("contract_status_raw contains KN but month is empty")
        if month is not None and not 1 <= month <= 12:
            reasons.append(
                f"contract_terminated_from_month={month} is outside 1..12"
            )
        if (
            month is not None
            and not record.is_contract_terminated
            and not KN_TOKEN_RE.search(record.contract_status_raw or "")
        ):
            reasons.append(
                "contract_terminated_from_month is set without legacy termination marker"
            )
        if reasons:
            item = _record_ref(record)
            item["reasons"] = reasons
            _add_item(sections["dirty_legacy_contract_rows"], item, limit)

    record_refs = {
        row["pkid"]: row
        for row in KarteiRecord.objects.using(database)
        .values("pkid", "id", "year", "family_id", "child_name", "status")
        .iterator()
    }

    type_groups = (
        ContractTypeEntry.objects.using(database)
        .values("record_id")
        .annotate(entry_count=Count("id"), min_month=Min("effective_from_month"))
        .order_by("record_id")
    )
    for group in type_groups.iterator():
        if group["min_month"] and group["min_month"] > 1:
            ref = record_refs.get(group["record_id"])
            if ref:
                _add_item(
                    sections["partial_timeline_anomalies"],
                    {
                        "record_pkid": ref["pkid"],
                        "record_id": ref["id"],
                        "year": ref["year"],
                        "family_id": ref["family_id"],
                        "child_name": ref["child_name"],
                        "status": ref["status"] or "NORMAL",
                        "timeline": "contract_type",
                        "entry_count": group["entry_count"],
                        "first_effective_from_month": group["min_month"],
                        "reason": "timeline starts after month 1 and depends on legacy fallback",
                    },
                    limit,
                )

    status_groups = (
        ContractStatusEntry.objects.using(database)
        .values("record_id")
        .annotate(entry_count=Count("id"), min_month=Min("effective_from_month"))
        .order_by("record_id")
    )
    for group in status_groups.iterator():
        if group["min_month"] and group["min_month"] > 1:
            ref = record_refs.get(group["record_id"])
            if ref:
                _add_item(
                    sections["partial_timeline_anomalies"],
                    {
                        "record_pkid": ref["pkid"],
                        "record_id": ref["id"],
                        "year": ref["year"],
                        "family_id": ref["family_id"],
                        "child_name": ref["child_name"],
                        "status": ref["status"] or "NORMAL",
                        "timeline": "contract_status",
                        "entry_count": group["entry_count"],
                        "first_effective_from_month": group["min_month"],
                        "reason": "timeline starts after month 1 and depends on legacy fallback",
                    },
                    limit,
                )

    years = set(
        KarteiRecord.objects.using(database)
        .values_list("year", flat=True)
        .distinct()
    )
    years.update(
        SubjectCategory.objects.using(database)
        .values_list("year", flat=True)
        .distinct()
    )
    for year in sorted(years):
        categories = list(
            SubjectCategory.objects.using(database)
            .filter(year=year, name__iexact="Nachhilfe")
            .values("id", "kind", "is_active")
        )
        has_group = any(cat["kind"] == CategoryKind.GROUP for cat in categories)
        if not has_group:
            has_individual = any(
                cat["kind"] == CategoryKind.INDIVIDUAL for cat in categories
            )
            if has_individual:
                reason = (
                    "GROUP Nachhilfe is missing; INDIVIDUAL Nachhilfe exists and "
                    "current unique(year, name) blocks the dual row until schema change"
                )
            else:
                reason = "GROUP Nachhilfe is missing for this year"
            _add_item(
                sections["missing_group_nachhilfe_years"],
                {
                    "year": year,
                    "nachhilfe_categories": categories,
                    "reason": reason,
                },
                limit,
            )

    report["summary"] = {
        name: section["count"] for name, section in sections.items()
    }
    return report


class Command(BaseCommand):
    help = (
        "Read-only PRICELIST V2 pre-implementation audit for legacy snapshots, "
        "pending split-brain risk, dirty contracts, and missing GROUP Nachhilfe."
    )

    def add_arguments(self, parser: CommandParser) -> None:
        parser.add_argument(
            "--database",
            default="default",
            help="Database alias to audit. Default: default.",
        )
        parser.add_argument(
            "--limit",
            type=int,
            default=100,
            help="Max items to print per section. Use 0 for no limit.",
        )
        parser.add_argument(
            "--format",
            choices=("text", "json"),
            default="text",
            help="Report format. Default: text.",
        )

    def handle(self, *args: Any, **options: Any) -> None:
        database = options["database"]
        limit = options["limit"]
        if limit < 0:
            raise ValueError("--limit must be >= 0")

        with transaction.atomic(using=database):
            report = _build_report(database=database, limit=limit)
            transaction.set_rollback(True, using=database)

        if options["format"] == "json":
            self.stdout.write(json.dumps(report, indent=2, default=str))
            return

        self._write_text_report(report)

    def _write_text_report(self, report: dict[str, Any]) -> None:
        self.stdout.write("PRICELIST V2 pre-implementation audit (read-only)")
        self.stdout.write(f"Database: {report['database']}")
        self.stdout.write("")
        self.stdout.write("Summary:")
        for name, count in report["summary"].items():
            self.stdout.write(f"  {name}: {count}")

        for name, section in report["sections"].items():
            self.stdout.write("")
            self.stdout.write(f"[{name}] count={section['count']}")
            if section["truncated"]:
                self.stdout.write(f"  truncated: {section['truncated']}")
            for item in section["items"]:
                self.stdout.write(f"  - {json.dumps(item, sort_keys=True, default=str)}")
