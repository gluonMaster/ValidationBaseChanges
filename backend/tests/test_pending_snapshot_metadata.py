from __future__ import annotations

import pytest

from apps.approvals.services import (
    _apply_snapshot_to_record,
    build_snapshot,
    create_or_update_pending_change_from_snapshot,
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
