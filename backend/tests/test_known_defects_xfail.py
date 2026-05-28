from __future__ import annotations

import pytest

from apps.karteien.models import KarteiRecord, MonthsMode
from apps.karteien.views import KarteiRecordUpdateView


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
