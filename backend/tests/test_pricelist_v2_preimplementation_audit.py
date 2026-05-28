from __future__ import annotations

import json
from decimal import Decimal
from io import StringIO

import pytest
from django.core.management import call_command

from apps.approvals.models import DeclinedChange, PendingChange
from apps.approvals.services import build_snapshot
from apps.catalog.models import CategoryKind, SubjectCategory
from apps.karteien.models import (
    ContractStatusEntry,
    ContractStatusKind,
    MonthsMode,
    RecordStatus,
)


pytestmark = pytest.mark.django_db


def _all_month_values(value: str) -> dict[str, str]:
    return {f"month_{month}": value for month in range(1, 13)}


def _run_audit_json() -> dict:
    stdout = StringIO()
    call_command(
        "audit_pricelist_v2_preimplementation",
        "--format=json",
        stdout=stdout,
    )
    return json.loads(stdout.getvalue())


def test_pricelist_v2_preimplementation_audit_reports_v1_pending_split_brain(
    kartei_record_builder,
):
    old_base_amounts = _all_month_values("25.00")
    new_base_amounts = _all_month_values("80.00")
    record = kartei_record_builder(
        status=RecordStatus.PENDING,
        months_mode=MonthsMode.AUTO,
        base_amounts=new_base_amounts,
        month_2=Decimal("80.00"),
    )
    snapshot = build_snapshot(record)
    snapshot["_old_base_amounts"] = old_base_amounts
    PendingChange.objects.create(
        record=record,
        snapshot=snapshot,
        admin_comment="legacy category apply",
    )

    report = _run_audit_json()

    assert report["read_only"] is True
    assert report["summary"]["pending_snapshot_v1"] == 1
    assert report["summary"]["legacy_pending_split_brain"] == 1

    item = report["sections"]["legacy_pending_split_brain"]["items"][0]
    assert item["record_pkid"] == record.pkid
    assert item["confidence"] == "high"
    assert any("live base_amounts differs" in reason for reason in item["reasons"])


def test_pricelist_v2_preimplementation_audit_reports_dirty_contracts_and_nachhilfe(
    kartei_record_builder,
):
    record = kartei_record_builder(
        year=2026,
        contract_status_raw="KN",
        is_contract_terminated=True,
        contract_terminated_from_month=None,
    )
    SubjectCategory.objects.create(
        year=record.year,
        name="Nachhilfe",
        kind=CategoryKind.INDIVIDUAL,
        yearly_rate=Decimal("0.00"),
        monthly_rate=Decimal("0.00"),
    )
    ContractStatusEntry.objects.create(
        record=record,
        effective_from_month=3,
        kind=ContractStatusKind.PAUSED,
    )

    report = _run_audit_json()

    assert report["summary"]["dirty_legacy_contract_rows"] == 1
    assert report["summary"]["partial_timeline_anomalies"] == 1
    assert report["summary"]["missing_group_nachhilfe_years"] == 1

    dirty = report["sections"]["dirty_legacy_contract_rows"]["items"][0]
    assert dirty["record_pkid"] == record.pkid
    assert len(dirty["reasons"]) == 2

    missing_year = report["sections"]["missing_group_nachhilfe_years"]["items"][0]
    assert missing_year["year"] == record.year
    assert "INDIVIDUAL Nachhilfe exists" in missing_year["reason"]


def test_pricelist_v2_preimplementation_audit_does_not_modify_rows(
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.DECLINED)
    DeclinedChange.objects.create(
        record=record,
        snapshot=build_snapshot(record),
        decline_reason="legacy decline",
    )

    before = {
        "pending": PendingChange.objects.count(),
        "declined": DeclinedChange.objects.count(),
        "records": type(record).objects.count(),
        "status": type(record).objects.get(pk=record.pk).status,
    }
    _run_audit_json()
    after = {
        "pending": PendingChange.objects.count(),
        "declined": DeclinedChange.objects.count(),
        "records": type(record).objects.count(),
        "status": type(record).objects.get(pk=record.pk).status,
    }

    assert after == before
