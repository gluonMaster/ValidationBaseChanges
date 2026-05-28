from __future__ import annotations

from decimal import Decimal

import pytest
from django.urls import reverse

from apps.approvals.models import PendingChange
from apps.approvals.services import (
    _apply_snapshot_to_record,
    apply_decision,
    build_snapshot,
    create_or_update_pending_change_from_snapshot,
)
from apps.catalog.models import Subject
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
            "effective_from_month": 1,
            "is_monthly": True,
        },
        pending_contract_status_entry={
            "effective_from_month": 5,
            "kind": ContractStatusKind.TERMINATED,
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
    assert ContractStatusEntry.objects.get(record=record).kind == ContractStatusKind.TERMINATED


@pytest.mark.django_db
def test_months_override_view_creates_v2_snapshot_without_live_prewrite(
    client,
    admin_user,
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        months_mode=MonthsMode.AUTO,
        month_1=Decimal("10.00"),
        month_2=Decimal("20.00"),
    )
    post_data = {
        "reason": "Manual correction",
        "confirm_override": "on",
        **{f"month_{month}": "0.00" for month in range(1, 13)},
        "month_1": "12.34",
        "month_2": "20.00",
    }
    client.force_login(admin_user)

    response = client.post(
        reverse("karteien:months_override", args=[record.pk]),
        post_data,
    )

    assert response.status_code == 302
    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()
    snapshot = pending.snapshot

    assert record.status == RecordStatus.PENDING
    assert record.months_mode == MonthsMode.AUTO
    assert record.month_1 == Decimal("10.00")
    assert snapshot["_snapshot_version"] == SNAPSHOT_VERSION_V2
    assert snapshot["_pending_action"] == "MONTHS_OVERRIDE"
    assert snapshot["month_1"] == "12.34"
    assert snapshot["_pending_nontracked_payload"]["months_mode"] == MonthsMode.OVERRIDE
    assert snapshot["_pending_meta"]["touched_months"] == list(range(1, 13))

    apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.months_mode == MonthsMode.OVERRIDE
    assert record.month_1 == Decimal("12.34")


@pytest.mark.django_db
def test_quick_set_subject_ref_view_freezes_ref_without_live_prewrite(
    client,
    admin_user,
    superadmin_user,
    kartei_record_builder,
):
    subject = Subject.objects.create(name="Mathematik")
    record = kartei_record_builder(subject1="Mathematik", subject1_ref=None)
    client.force_login(admin_user)

    response = client.post(
        reverse("karteien:record_quick_set_subject_ref", args=[record.pk]),
        {"semester": "sem1"},
    )

    assert response.status_code == 302
    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()
    snapshot = pending.snapshot

    assert record.status == RecordStatus.PENDING
    assert record.subject1_ref_id is None
    assert snapshot["_snapshot_version"] == SNAPSHOT_VERSION_V2
    assert snapshot["_pending_action"] == "QUICK_SET_SUBJECT_REF"
    assert snapshot["_pending_nontracked_payload"]["subject1_ref_id"] == subject.pk
    assert snapshot["_pending_meta"]["touched_months"] == [1, 2, 3, 4, 5, 6]

    apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.subject1_ref_id == subject.pk


@pytest.mark.django_db
def test_contract_type_change_view_uses_v2_payload_without_billing_rewrite(
    client,
    admin_user,
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        is_monthly_contract=False,
        contract_type_raw="",
        base_amounts={"month_1": "10.00"},
        month_1=Decimal("10.00"),
    )
    client.force_login(admin_user)

    response = client.post(
        reverse("karteien:record_contract_type_change", args=[record.pk]),
        {
            "effective_from_month": "1",
            "is_monthly": "1",
            "comment": "Switch to monthly",
        },
    )

    assert response.status_code == 302
    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()
    snapshot = pending.snapshot

    assert record.status == RecordStatus.PENDING
    assert record.is_monthly_contract is False
    assert record.base_amounts == {"month_1": "10.00"}
    assert record.month_1 == Decimal("10.00")
    assert snapshot["_snapshot_version"] == SNAPSHOT_VERSION_V2
    assert snapshot["_pending_action"] == "CONTRACT_TYPE_CHANGE"
    assert snapshot["_pending_contract_type_entry"] == {
        "effective_from_month": 1,
        "is_monthly": True,
        "comment": "Switch to monthly",
        "changed_by_id": admin_user.id,
    }
    assert snapshot["_pending_nontracked_payload"]["is_monthly_contract"] is True
    assert snapshot["_pending_nontracked_payload"]["contract_type_raw"] == "O/V"
    assert snapshot["_pending_nontracked_payload"]["base_amounts"] == {
        "month_1": "10.00",
    }

    apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.is_monthly_contract is True
    assert record.base_amounts == {"month_1": "10.00"}
    assert record.month_1 == Decimal("10.00")
    assert ContractTypeEntry.objects.get(record=record).is_monthly is True


@pytest.mark.django_db
def test_contract_status_change_view_v2_zeroes_snapshot_only_until_approve(
    client,
    admin_user,
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(
        months_mode=MonthsMode.AUTO,
        base_amounts={f"month_{month}": "100.00" for month in range(1, 13)},
        month_3=Decimal("100.00"),
        is_contract_terminated=False,
        contract_status_raw="",
        contract_terminated_from_month=None,
    )
    client.force_login(admin_user)

    response = client.post(
        reverse("karteien:record_contract_status_change", args=[record.pk]),
        {
            "effective_from_month": "3",
            "kind": ContractStatusKind.TERMINATED,
            "comment": "Terminate from March",
        },
    )

    assert response.status_code == 302
    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()
    snapshot = pending.snapshot

    assert record.status == RecordStatus.PENDING
    assert record.month_3 == Decimal("100.00")
    assert record.is_contract_terminated is False
    assert record.contract_terminated_from_month is None
    assert snapshot["_snapshot_version"] == SNAPSHOT_VERSION_V2
    assert snapshot["_pending_action"] == "CONTRACT_STATUS_CHANGE"
    assert snapshot["_pending_contract_status_entry"] == {
        "effective_from_month": 3,
        "kind": ContractStatusKind.TERMINATED,
        "comment": "Terminate from March",
        "changed_by_id": admin_user.id,
    }
    assert snapshot["month_3"] == "0.00"
    assert snapshot["_pending_nontracked_payload"]["is_contract_terminated"] is True
    assert snapshot["_pending_nontracked_payload"]["contract_terminated_from_month"] == 3
    assert snapshot["_pending_meta"]["touched_months"] == list(range(3, 13))

    apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.month_3 == Decimal("0.00")
    assert record.is_contract_terminated is True
    assert record.contract_terminated_from_month == 3
    assert ContractStatusEntry.objects.get(record=record).kind == ContractStatusKind.TERMINATED


@pytest.mark.django_db
def test_apply_decision_approved_blocks_paused_timeline_until_prompt09(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    snapshot = build_snapshot_v2(
        record,
        pending_action="CONTRACT_STATUS_CHANGE",
        pending_contract_status_entry={
            "effective_from_month": 3,
            "kind": ContractStatusKind.PAUSED,
        },
        pending_meta={"audit_summary": "Pause from March"},
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    with pytest.raises(ValueError, match="PAUSED"):
        apply_decision(pending, "APPROVED", "ok", superadmin_user)

    record.refresh_from_db()
    pending.refresh_from_db()
    assert record.status == RecordStatus.PENDING
    assert pending.is_processed is False
    assert ContractStatusEntry.objects.filter(record=record).count() == 0


@pytest.mark.django_db
def test_apply_decision_approved_blocks_active_reactivation_until_prompt09(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    ContractStatusEntry.objects.create(
        record=record,
        effective_from_month=2,
        kind=ContractStatusKind.TERMINATED,
    )
    snapshot = build_snapshot_v2(
        record,
        pending_action="CONTRACT_STATUS_CHANGE",
        pending_contract_status_entry={
            "effective_from_month": 5,
            "kind": ContractStatusKind.ACTIVE,
        },
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    with pytest.raises(ValueError, match="ACTIVE reactivation"):
        apply_decision(pending, "APPROVED", "ok", superadmin_user)

    pending.refresh_from_db()
    assert pending.is_processed is False


@pytest.mark.django_db
def test_apply_decision_approved_blocks_intrayear_contract_type_until_prompt09(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    snapshot = build_snapshot_v2(
        record,
        pending_action="CONTRACT_TYPE_CHANGE",
        pending_contract_type_entry={
            "effective_from_month": 6,
            "is_monthly": True,
        },
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    with pytest.raises(ValueError, match="intra-year contract type"):
        apply_decision(pending, "APPROVED", "ok", superadmin_user)

    pending.refresh_from_db()
    assert pending.is_processed is False
    assert ContractTypeEntry.objects.filter(record=record).count() == 0


@pytest.mark.django_db
def test_apply_decision_approved_timeline_only_history_uses_snapshot_summary(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    snapshot = build_snapshot_v2(
        record,
        pending_action="CONTRACT_STATUS_CHANGE",
        pending_contract_status_entry={
            "effective_from_month": 7,
            "kind": ContractStatusKind.TERMINATED,
        },
        pending_meta={"audit_summary": "Terminate from July"},
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(pending, "APPROVED", None, superadmin_user)

    record.refresh_from_db()
    assert "APR:" in record.history_raw
    assert "Terminate from July" in record.history_raw


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
def test_apply_decision_declined_snapshot_v2_uses_legacy_marker_rollback(
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
    )
    snapshot["_old_base_amounts"] = {"month_1": "10.00"}
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
def test_apply_decision_declined_history_uses_snapshot_audit_summary(
    superadmin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    snapshot = build_snapshot_v2(
        record,
        pending_action="FORM_SAVE",
        pending_nontracked_payload={"sepa_marker": "SEPA"},
        pending_meta={"audit_summary": "SEPA marker change"},
    )
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(pending, "DECLINED", "not accepted", superadmin_user)

    record.refresh_from_db()
    assert "DCL(" in record.history_raw
    assert "not accepted" in record.history_raw
    assert "SEPA marker change" in record.history_raw


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
