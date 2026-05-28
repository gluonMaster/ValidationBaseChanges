from __future__ import annotations

import pytest
from django.urls import reverse

from apps.approvals.models import PendingChange
from apps.approvals.services import (
    _apply_snapshot_to_record,
    apply_decision,
    build_snapshot,
    create_or_update_pending_change_from_snapshot,
)
from apps.karteien.models import (
    ContractStatusEntry,
    ContractStatusKind,
    ContractTypeEntry,
    MonthsMode,
    RecordStatus,
)
from apps.karteien.services.pending_snapshot import (
    SNAPSHOT_VERSION_V2,
    build_projected_record_from_snapshot,
    build_snapshot_v2,
)


@pytest.fixture(autouse=True)
def disable_notification_side_effects(monkeypatch):
    from apps.notifications import services as notification_services

    monkeypatch.setattr(
        notification_services,
        "notify_pending_created",
        lambda *args, **kwargs: None,
    )
    monkeypatch.setattr(
        notification_services,
        "notify_approved",
        lambda *args, **kwargs: None,
    )
    monkeypatch.setattr(
        notification_services,
        "notify_declined_created",
        lambda *args, **kwargs: None,
    )
    monkeypatch.setattr(
        notification_services,
        "mark_pending_notifications_read_for_record",
        lambda *args, **kwargs: None,
    )


def test_apply_snapshot_ignores_reserved_metadata_keys():
    from apps.karteien.models import KarteiRecord

    record = KarteiRecord(id=1, year=2026, family_id="FAM-1", parent_name="Old")
    snapshot = {
        "parent_name": "New",
        "_old_base_amounts": {"month_1": "100.00"},
        "_pending_contract_type_entry": {
            "effective_from_month": 1,
            "is_monthly": True,
        },
    }

    updated = _apply_snapshot_to_record(record, snapshot)

    assert updated.parent_name == "New"
    assert not hasattr(updated, "_old_base_amounts")
    assert not hasattr(updated, "_pending_contract_type_entry")


@pytest.mark.django_db
def test_snapshot_v1_projection_applies_tracked_fields_only(kartei_record_builder):
    record = kartei_record_builder(parent_name="Old", sepa_marker="")
    snapshot = build_snapshot(record)
    snapshot["parent_name"] = "New"
    snapshot["_pending_nontracked_payload"] = {"sepa_marker": "SEPA"}

    projection = build_projected_record_from_snapshot(record, snapshot)

    assert projection.record.parent_name == "New"
    assert projection.record.sepa_marker == ""
    assert projection.timeline.contract_type_entry is None
    assert projection.timeline.contract_status_entry is None


@pytest.mark.django_db
def test_snapshot_v2_projection_applies_tracked_nontracked_and_timeline(
    kartei_record_builder,
):
    record = kartei_record_builder(
        parent_name="Old",
        sepa_marker="",
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "10.00"},
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="FORM_SAVE",
        pending_nontracked_payload={
            "sepa_marker": "SEPA",
            "months_mode": MonthsMode.OVERRIDE,
            "base_amounts": {"month_1": "20.00"},
        },
        pending_contract_type_entry={
            "effective_from_month": "3",
            "is_monthly": True,
        },
        pending_contract_status_entry={
            "effective_from_month": 4,
            "kind": ContractStatusKind.PAUSED,
        },
        pending_meta={"touched_months": ["4", "2"], "warnings": ["check"]},
    )
    snapshot["parent_name"] = "New"

    projection = build_projected_record_from_snapshot(record, snapshot)

    assert snapshot["_snapshot_version"] == SNAPSHOT_VERSION_V2
    assert projection.record.parent_name == "New"
    assert projection.record.sepa_marker == "SEPA"
    assert projection.record.is_sepa is True
    assert projection.record.months_mode == MonthsMode.OVERRIDE
    assert projection.record.base_amounts == {"month_1": "20.00"}
    assert projection.timeline.contract_type_entry == {
        "effective_from_month": 3,
        "is_monthly": True,
    }
    assert projection.timeline.contract_status_entry == {
        "effective_from_month": 4,
        "kind": ContractStatusKind.PAUSED,
    }
    assert projection.pending_meta["touched_months"] == [2, 4]


@pytest.mark.django_db
def test_pending_change_from_explicit_snapshot_preserves_reserved_metadata(
    monkeypatch,
    kartei_record_builder,
):
    from apps.notifications import services as notification_services

    monkeypatch.setattr(
        notification_services,
        "notify_pending_created",
        lambda *args, **kwargs: None,
    )
    record = kartei_record_builder()
    snapshot = build_snapshot(record)
    snapshot["_old_base_amounts"] = {"month_1": "100.00"}
    snapshot["_pending_contract_type_entry"] = {
        "effective_from_month": 1,
        "is_monthly": True,
        "comment": "Test type proposal",
    }
    snapshot["_pending_contract_status_entry"] = {
        "effective_from_month": 7,
        "kind": "PAUSED",
        "comment": "Test status proposal",
    }

    pending = create_or_update_pending_change_from_snapshot(
        record,
        snapshot=snapshot,
        admin_comment="Keep metadata",
    )
    pending.refresh_from_db()

    assert pending.snapshot["_old_base_amounts"] == {"month_1": "100.00"}
    assert pending.snapshot["_pending_contract_type_entry"]["effective_from_month"] == 1
    assert pending.snapshot["_pending_contract_status_entry"]["kind"] == "PAUSED"


@pytest.mark.django_db
def test_apply_decision_approved_applies_snapshot_v2_frozen_payloads(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        status=RecordStatus.PENDING,
        parent_name="Old",
        sepa_marker="",
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "10.00"},
        month_1="10.00",
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="FORM_SAVE",
        pending_nontracked_payload={
            "sepa_marker": "SEPA",
            "months_mode": MonthsMode.OVERRIDE,
            "base_amounts": {"month_1": "20.00"},
        },
        pending_contract_type_entry={
            "effective_from_month": 2,
            "is_monthly": True,
        },
        pending_contract_status_entry={
            "effective_from_month": 5,
            "kind": ContractStatusKind.PAUSED,
        },
    )
    snapshot["parent_name"] = "New"
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.parent_name == "New"
    assert record.sepa_marker == "SEPA"
    assert record.months_mode == MonthsMode.OVERRIDE
    assert record.base_amounts == {"month_1": "20.00"}
    assert ContractTypeEntry.objects.get(record=record).is_monthly is True
    assert ContractStatusEntry.objects.get(record=record).kind == ContractStatusKind.PAUSED


@pytest.mark.django_db
def test_apply_decision_declined_snapshot_v2_without_rollback_leaves_live_payload(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        status=RecordStatus.PENDING,
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "10.00"},
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="APPLY_CATEGORY",
        pending_nontracked_payload={"base_amounts": {"month_1": "20.00"}},
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(pending, "DECLINED", "no", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.DECLINED
    assert record.base_amounts == {"month_1": "10.00"}


@pytest.mark.django_db
def test_apply_decision_declined_snapshot_v2_uses_explicit_rollback_payload(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        status=RecordStatus.PENDING,
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "20.00"},
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="APPLY_CATEGORY",
        pending_nontracked_payload={"base_amounts": {"month_1": "20.00"}},
        rollback_nontracked_payload={"base_amounts": {"month_1": "10.00"}},
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(pending, "DECLINED", "no", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.DECLINED
    assert record.base_amounts == {"month_1": "10.00"}


@pytest.mark.django_db
def test_admin_pending_detail_renders_snapshot_v2_projection(
    client,
    admin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        status=RecordStatus.PENDING,
        child_name="Live V2 Child",
        months_mode=MonthsMode.AUTO,
        sepa_marker="",
        base_amounts={"month_1": "10.00"},
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="FORM_SAVE",
        pending_nontracked_payload={
            "sepa_marker": "SEPA",
            "months_mode": MonthsMode.OVERRIDE,
            "discounts_disabled": True,
            "discounts_disabled_months": [1],
        },
        pending_contract_type_entry={
            "effective_from_month": 2,
            "is_monthly": True,
        },
        pending_contract_status_entry={
            "effective_from_month": 3,
            "kind": ContractStatusKind.PAUSED,
        },
    )
    snapshot["child_name"] = "Projected V2 Child"
    PendingChange.objects.create(record=record, snapshot=snapshot)
    client.force_login(admin_user)

    response = client.get(
        reverse("karteien:record_detail", args=[record.pk]),
        {"view": "pending"},
    )

    assert response.status_code == 200
    content = response.content.decode()
    assert "Projected V2 Child" in content
    assert "SEPA" in content
    assert "OVERRIDE" in content
    assert "Rabatte deaktiviert" in content
