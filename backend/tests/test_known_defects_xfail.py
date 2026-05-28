from __future__ import annotations

import pytest
from django.urls import reverse

from apps.karteien.models import KarteiRecord, MonthsMode
from apps.karteien.views import KarteiRecordUpdateView


@pytest.mark.django_db
@pytest.mark.xfail(
    strict=True,
    reason="PROMPT_162 not applied: Operator group navigation/catalog entry remains admin-only.",
)
def test_operator_can_open_catalog_index_for_group_navigation(operator_user, client):
    client.force_login(operator_user)

    response = client.get(reverse("catalog:index"))

    assert response.status_code == 200


@pytest.mark.xfail(
    strict=True,
    reason="PRICELIST V2 no-prewrite frozen payload semantics not applied: billing context fields are still live safe updates.",
)
def test_safe_field_update_set_does_not_live_prewrite_billing_context():
    record = KarteiRecord(
        id=1,
        year=2026,
        family_id="FAM-1",
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "100.00"},
        hours_amounts={"month_1": "2.00"},
    )

    safe_updates = KarteiRecordUpdateView()._build_safe_fields_update(record)

    assert {
        "months_mode",
        "base_amounts",
        "hours_amounts",
    }.isdisjoint(safe_updates)
