from __future__ import annotations

from dataclasses import FrozenInstanceError
from decimal import Decimal

import pytest

from apps.karteien.models import MonthsMode
from apps.karteien.services.billing_pipeline import (
    build_legacy_to_auto_proposal,
    build_price_option_recalc_proposal,
)


pytestmark = pytest.mark.django_db


def test_price_option_recalc_builder_returns_frozen_v2_proposal_without_live_mutation(
    kartei_record_builder,
):
    record = kartei_record_builder(
        subject1="Mathematik",
        price1=Decimal("30.00"),
        months_mode=MonthsMode.AUTO,
        base_amounts={f"month_{month}": "10.00" for month in range(1, 13)},
        month_2=Decimal("10.00"),
    )

    proposal = build_price_option_recalc_proposal(
        record,
        apply_from_month_1=2,
    )

    record.refresh_from_db()
    assert record.base_amounts["month_2"] == "10.00"
    assert record.month_2 == Decimal("10.00")

    assert proposal.action == "PRICE_OPTION_RECALC"
    assert proposal.snapshot["_pending_action"] == "PRICE_OPTION_RECALC"
    assert proposal.snapshot["_pending_nontracked_payload"]["base_amounts"]["month_2"] == "30.00"
    assert proposal.snapshot["month_2"] == "30.00"
    assert proposal.snapshot["_pending_meta"]["touched_months"] == [2, 3, 4, 5, 6]

    with pytest.raises(FrozenInstanceError):
        proposal.action = "CHANGED"


def test_legacy_to_auto_builder_recalculates_all_months_in_snapshot_only(
    kartei_record_builder,
):
    record = kartei_record_builder(
        subject1="Mathematik",
        price1=Decimal("40.00"),
        months_mode=MonthsMode.LEGACY,
        month_1=Decimal("11.00"),
        month_2=Decimal("22.00"),
    )

    proposal = build_legacy_to_auto_proposal(record)

    record.refresh_from_db()
    assert record.months_mode == MonthsMode.LEGACY
    assert record.month_1 == Decimal("11.00")

    assert proposal.snapshot["_pending_action"] == "LEGACY_TO_AUTO"
    assert proposal.snapshot["_pending_nontracked_payload"]["months_mode"] == MonthsMode.AUTO
    assert proposal.snapshot["_pending_nontracked_payload"]["base_amounts"]["month_1"] == "40.00"
    assert proposal.snapshot["month_1"] == "40.00"
    assert proposal.snapshot["_pending_meta"]["touched_months"] == list(range(1, 13))
