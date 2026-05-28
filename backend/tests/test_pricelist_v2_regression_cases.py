from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest
from django.urls import reverse

from apps.approvals.models import DeclinedChange, PendingChange
from apps.approvals.services import apply_decision, build_snapshot
from apps.catalog.group_size_service import calculate_auto_group_size
from apps.catalog.models import (
    CategoryKind,
    DisciplineGroup,
    Discount,
    DiscountKind,
    DurationEntry,
    FamilyDiscount,
    Subject,
    SubjectCategory,
    SubjectCategoryLink,
)
from apps.karteien.billing import (
    calculate_month_values,
    get_month_mismatches,
    recalculate_legacy_to_auto,
)
from apps.karteien.models import (
    ContractStatusEntry,
    ContractStatusKind,
    MonthsMode,
    RecordStatus,
)


pytestmark = pytest.mark.django_db


def _all_month_values(value: str) -> dict[str, str]:
    return {f"month_{month}": value for month in range(1, 13)}


def _all_month_decimals(value: str) -> dict[str, Decimal]:
    return {f"month_{month}": Decimal(value) for month in range(1, 13)}


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


def _create_individual_category_case(kartei_record_builder, *, year: int = 2026):
    subject = Subject.objects.create(name=f"Prompt182 Ind {year}")
    category = SubjectCategory.objects.create(
        year=year,
        name=f"Prompt182 Individual {year}",
        kind=CategoryKind.INDIVIDUAL,
        yearly_rate=Decimal("40.00"),
        monthly_rate=Decimal("50.00"),
    )
    SubjectCategoryLink.objects.create(subject=subject, category=category)

    record = kartei_record_builder(
        year=year,
        subject1_ref=subject,
        subject1=subject.name,
        months_mode=MonthsMode.AUTO,
        base_amounts=_all_month_values("25.00"),
        hours_amounts=_all_month_values("2.00"),
        month_1=Decimal("25.00"),
        month_2=Decimal("25.00"),
        month_3=Decimal("25.00"),
        month_4=Decimal("25.00"),
        month_5=Decimal("25.00"),
        month_6=Decimal("25.00"),
    )
    return record


def _create_group_case(
    kartei_record_builder,
    *,
    year: int = 2026,
    record_count: int = 1,
):
    subject = Subject.objects.create(name=f"Prompt182 Group {year}")
    category = SubjectCategory.objects.create(
        year=year,
        name=f"Prompt182 Group Category {year}",
        kind=CategoryKind.GROUP,
        yearly_rate=Decimal("45.00"),
        monthly_rate=Decimal("60.00"),
        group_threshold=6,
    )
    SubjectCategoryLink.objects.create(subject=subject, category=category)
    group = DisciplineGroup.objects.create(
        subject=subject,
        year=year,
        category=category,
        auto_scaling_enabled=False,
    )
    DurationEntry.objects.create(
        group=group,
        effective_from_month=1,
        duration_minutes=45,
    )

    records = [
        kartei_record_builder(
            year=year,
            family_id=f"FAM-{year}-GROUP-{index}",
            subject1_ref=subject,
            subject1=subject.name,
            months_mode=MonthsMode.AUTO,
            base_amounts=_all_month_values("20.00"),
            month_1=Decimal("20.00"),
            month_2=Decimal("20.00"),
            month_3=Decimal("20.00"),
            month_4=Decimal("20.00"),
            month_5=Decimal("20.00"),
            month_6=Decimal("20.00"),
        )
        for index in range(record_count)
    ]
    return group, records


def test_auto_base_amounts_are_historical_base_not_silent_price_cache(
    kartei_record_builder,
):
    record = kartei_record_builder(
        subject1="Mathematik",
        price1=Decimal("999.00"),
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "100.00"},
        month_1=Decimal("100.00"),
    )

    assert get_month_mismatches(record) == set()


def test_single_category_apply_creates_pending_proposal_with_current_metadata(
    client,
    admin_user,
    kartei_record_builder,
):
    record = _create_individual_category_case(kartei_record_builder)
    client.force_login(admin_user)

    response = client.post(
        reverse("karteien:record_apply_price", args=[record.pk]),
        {
            "semester": "1",
            "from_month": "2",
            "comment": "PROMPT_182 single apply regression",
        },
    )

    assert response.status_code == 302

    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()

    assert record.status == RecordStatus.PENDING
    assert record.base_amounts["month_2"] == "80.00"
    assert pending.snapshot["month_2"] == "80.00"
    assert pending.snapshot["_old_base_amounts"]["month_2"] == "25.00"


def test_bulk_category_apply_creates_pending_proposals_for_group_records(
    client,
    admin_user,
    kartei_record_builder,
):
    group, records = _create_group_case(kartei_record_builder, record_count=2)
    client.force_login(admin_user)

    response = client.post(
        reverse(
            "catalog:group_apply_bulk",
            kwargs={"year": group.year, "pk": group.pk},
        ),
        {
            "semester": "1",
            "from_month": "2",
            "comment": "PROMPT_182 bulk apply regression",
        },
    )

    assert response.status_code == 302
    assert PendingChange.objects.filter(record__in=records, is_processed=False).count() == 2

    for record in records:
        record.refresh_from_db()
        pending = PendingChange.objects.get(record=record, is_processed=False)

        assert record.status == RecordStatus.PENDING
        assert record.base_amounts["month_2"] == "45.00"
        assert pending.snapshot["month_2"] == "45.00"
        assert pending.snapshot["_old_base_amounts"]["month_2"] == "20.00"


def test_bulk_category_apply_operator_restrictions_block_past_months(
    client,
    operator_user,
    kartei_record_builder,
    monkeypatch,
):
    from apps.karteien import validators

    class FixedDate(date):
        @classmethod
        def today(cls):
            return cls(2026, 5, 27)

    monkeypatch.setattr(validators, "date", FixedDate)

    group, records = _create_group_case(kartei_record_builder, year=2026)
    record = records[0]
    client.force_login(operator_user)

    response = client.post(
        reverse(
            "catalog:group_apply_bulk",
            kwargs={"year": group.year, "pk": group.pk},
        ),
        {
            "semester": "1",
            "from_month": "4",
            "comment": "Operator must not change past months",
        },
    )

    assert response.status_code == 302
    assert PendingChange.objects.filter(record=record, is_processed=False).count() == 0

    record.refresh_from_db()
    assert record.status == RecordStatus.NORMAL
    assert record.base_amounts["month_4"] == "20.00"


def test_contract_status_active_paused_terminated_month_values(
    kartei_record_builder,
):
    record = kartei_record_builder(
        subject1="Mathematik",
        months_mode=MonthsMode.AUTO,
        base_amounts=_all_month_values("100.00"),
    )
    ContractStatusEntry.objects.create(
        record=record,
        effective_from_month=2,
        kind=ContractStatusKind.PAUSED,
    )
    ContractStatusEntry.objects.create(
        record=record,
        effective_from_month=3,
        kind=ContractStatusKind.TERMINATED,
    )

    values, flags = calculate_month_values(
        record,
        base_amounts=_all_month_decimals("100.00"),
    )

    assert values["month_1"] == Decimal("100.00")
    assert values["month_2"] == Decimal("0.00")
    assert values["month_3"] == Decimal("0.00")
    assert flags.paused_months == [2]
    assert 3 in flags.terminated_months


def test_group_size_counts_active_and_paused_but_excludes_terminated(
    kartei_record_builder,
):
    group, records = _create_group_case(kartei_record_builder, record_count=3)
    active_record, paused_record, terminated_record = records
    ContractStatusEntry.objects.create(
        record=paused_record,
        effective_from_month=1,
        kind=ContractStatusKind.PAUSED,
    )
    ContractStatusEntry.objects.create(
        record=terminated_record,
        effective_from_month=1,
        kind=ContractStatusKind.TERMINATED,
    )

    size = calculate_auto_group_size(group, month=1)

    assert active_record.pkid in size["record_pkids"]
    assert paused_record.pkid in size["record_pkids"]
    assert terminated_record.pkid not in size["record_pkids"]
    assert size["total_size"] == 2
    assert size["billable_count"] == 1


def test_nachhilfe_month_values_do_not_apply_discounts(kartei_record_builder):
    record = kartei_record_builder(
        family_id="FAM-NH",
        subject1="Nachhilfe Mathematik",
        months_mode=MonthsMode.AUTO,
        base_amounts={"month_1": "100.00"},
    )
    discount = Discount.objects.create(
        kind=DiscountKind.PERCENT,
        value=Decimal("0.50"),
        description="Family discount that Nachhilfe must ignore",
    )
    FamilyDiscount.objects.create(
        year=record.year,
        family_id=record.family_id,
        discount=discount,
        start_month=1,
        end_month=12,
    )

    values, flags = calculate_month_values(
        record,
        base_amounts={"month_1": Decimal("100.00")},
    )

    assert values["month_1"] == Decimal("100.00")
    assert 1 in flags.nachhilfe_exempt_months


@pytest.mark.xfail(
    strict=True,
    reason="PROMPT_160 not applied",
)
def test_target_schema_allows_dual_group_and_individual_nachhilfe_categories():
    SubjectCategory.objects.create(
        year=2026,
        name="Nachhilfe",
        kind=CategoryKind.GROUP,
        yearly_rate=Decimal("0.00"),
        monthly_rate=Decimal("0.00"),
    )
    SubjectCategory.objects.create(
        year=2026,
        name="Nachhilfe",
        kind=CategoryKind.INDIVIDUAL,
        yearly_rate=Decimal("0.00"),
        monthly_rate=Decimal("0.00"),
    )

    assert SubjectCategory.objects.filter(year=2026, name="Nachhilfe").count() == 2


def test_explicit_legacy_to_auto_conversion_updates_only_touched_months(
    kartei_record_builder,
):
    record = kartei_record_builder(
        subject1="Mathematik",
        price1=Decimal("100.00"),
        months_mode=MonthsMode.LEGACY,
        month_1=Decimal("11.00"),
        month_2=Decimal("22.00"),
    )

    recalculate_legacy_to_auto(record, touched_months={2})

    assert record.months_mode == MonthsMode.AUTO
    assert record.month_1 == Decimal("11.00")
    assert record.month_2 == Decimal("100.00")
    assert record.base_amounts["month_2"] == "100.00"


def test_current_decline_rolls_back_old_base_amounts_marker(
    superadmin_user,
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
    pending = PendingChange.objects.create(record=record, snapshot=snapshot)

    apply_decision(
        pending,
        "DECLINED",
        "PROMPT_182 rollback regression",
        superadmin_user,
    )

    record.refresh_from_db()
    pending.refresh_from_db()

    assert record.status == RecordStatus.DECLINED
    assert record.base_amounts == old_base_amounts
    assert pending.is_processed is True
    assert DeclinedChange.objects.filter(record=record).exists()


@pytest.mark.xfail(
    strict=True,
    reason="target snapshot v2 behavior, not implemented before PROMPT_166...178",
)
def test_target_category_apply_decline_does_not_need_live_prewrite_or_rollback_marker(
    client,
    admin_user,
    kartei_record_builder,
):
    record = _create_individual_category_case(kartei_record_builder)
    client.force_login(admin_user)

    client.post(
        reverse("karteien:record_apply_price", args=[record.pk]),
        {
            "semester": "1",
            "from_month": "2",
            "comment": "Target no-prewrite regression",
        },
    )

    pending = PendingChange.objects.get(record=record, is_processed=False)
    record.refresh_from_db()

    assert record.status == RecordStatus.PENDING
    assert record.base_amounts["month_2"] == "25.00"
    assert "_old_base_amounts" not in pending.snapshot


@pytest.mark.xfail(
    strict=True,
    reason="known PRICELIST V2 split-brain before stabilization",
)
def test_target_billing_preview_and_category_apply_preview_use_same_pipeline(
    client,
    admin_user,
    kartei_record_builder,
):
    record = _create_individual_category_case(kartei_record_builder)
    client.force_login(admin_user)

    apply_preview = client.get(
        reverse("karteien:record_apply_price_preview", args=[record.pk]),
        {
            "semester": "1",
            "from_month": "2",
        },
    ).json()
    billing_preview = client.get(
        reverse("karteien:billing_preview_api", args=[record.pk]),
        {
            "subject1_ref": str(record.subject1_ref_id),
            "start_month_1": "1",
            "hours_month_2": "2.00",
        },
    ).json()

    assert billing_preview["months"]["2"]["value"] == apply_preview["new_months"]["month_2"]


@pytest.mark.xfail(
    strict=True,
    reason="target snapshot v2 behavior, not implemented before PROMPT_166...178",
)
def test_target_bulk_apply_does_not_prewrite_live_base_amounts(
    client,
    admin_user,
    kartei_record_builder,
):
    group, records = _create_group_case(kartei_record_builder)
    record = records[0]
    client.force_login(admin_user)

    client.post(
        reverse(
            "catalog:group_apply_bulk",
            kwargs={"year": group.year, "pk": group.pk},
        ),
        {
            "semester": "1",
            "from_month": "2",
            "comment": "Target bulk no-prewrite regression",
        },
    )

    assert PendingChange.objects.filter(record=record, is_processed=False).exists()
    record.refresh_from_db()
    assert record.status == RecordStatus.PENDING
    assert record.base_amounts["month_2"] == "20.00"
