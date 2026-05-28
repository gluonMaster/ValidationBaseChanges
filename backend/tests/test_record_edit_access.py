from __future__ import annotations

import pytest
from django.urls import reverse

from apps.approvals.models import DeclinedChange, PendingChange
from apps.approvals.services import build_snapshot
from apps.karteien.models import RecordStatus


@pytest.mark.django_db
def test_admin_can_open_standard_editor_for_pending_record(
    client,
    admin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.PENDING)
    PendingChange.objects.create(record=record, snapshot=build_snapshot(record))
    client.force_login(admin_user)

    response = client.get(reverse("karteien:record_update", args=[record.pk]))

    assert response.status_code == 200


@pytest.mark.django_db
def test_admin_can_open_standard_editor_for_declined_record(
    client,
    admin_user,
    kartei_record_builder,
):
    record = kartei_record_builder(status=RecordStatus.DECLINED)
    declined = DeclinedChange.objects.create(
        record=record,
        snapshot=build_snapshot(record),
        decline_reason="Needs correction",
    )
    client.force_login(admin_user)

    response = client.get(
        reverse("karteien:record_update", args=[record.pk]),
        {"declined_change_id": declined.pk},
    )

    assert response.status_code == 200


@pytest.mark.django_db
@pytest.mark.parametrize("status", [RecordStatus.PENDING, RecordStatus.DECLINED])
def test_operator_is_blocked_from_editing_pending_or_declined_records(
    client,
    operator_user,
    kartei_record_builder,
    status,
):
    record = kartei_record_builder(status=status)
    client.force_login(operator_user)

    response = client.get(reverse("karteien:record_update", args=[record.pk]))

    assert response.status_code == 302
    assert response.url == reverse("karteien:record_detail", args=[record.pk])
