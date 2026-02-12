"""
Pricing engine — suggested-price calculation (read-only, no side-effects).

Public API
----------
* ``determine_pricing_source(record, semester)``
    → ``PricingSource.CATEGORY`` or ``PricingSource.PRICE_OPTION``

* ``calculate_suggested_price(record, month)``
    → dict with pricing breakdown, or ``None``

* ``get_suggested_prices_for_record(record)``
    → ``{1: info|None, 2: …, 12: info|None}``

The module never writes to ``record.base_amounts``, ``record.month_*``,
or any other field.  It is purely a calculator.
"""

from __future__ import annotations

import enum
from decimal import Decimal
from typing import TYPE_CHECKING

from apps.catalog.models import (
    CategoryKind,
    DisciplineGroup,
    DurationEntry,
    SubjectCategoryLink,
    get_duration_for_month,
)
from apps.catalog.group_size_service import get_group_size_for_month
from apps.karteien.billing import get_semester_for_month, round_money_up
from apps.karteien.models import get_contract_type_for_month

if TYPE_CHECKING:
    from apps.karteien.models import KarteiRecord


# =============================================================================
# PricingSource enum
# =============================================================================

class PricingSource(enum.Enum):
    """How the suggested price is determined."""
    CATEGORY = "CATEGORY"
    PRICE_OPTION = "PRICE_OPTION"


# =============================================================================
# Determine pricing source
# =============================================================================

def determine_pricing_source(record: "KarteiRecord", semester: int) -> PricingSource:
    """
    Decide whether a record's price comes from a category or a price-option.

    Args:
        record:   KarteiRecord instance.
        semester: 1 or 2.

    Returns:
        ``PricingSource.CATEGORY`` when the subject has an active
        ``SubjectCategoryLink`` for the record's year;
        ``PricingSource.PRICE_OPTION`` otherwise.
    """
    subject_ref = record.subject1_ref if semester == 1 else record.subject2_ref
    if subject_ref is None:
        return PricingSource.PRICE_OPTION

    has_link = SubjectCategoryLink.objects.filter(
        subject=subject_ref,
        year=record.year,
        category__is_active=True,
    ).exists()

    return PricingSource.CATEGORY if has_link else PricingSource.PRICE_OPTION


# =============================================================================
# Internal helpers
# =============================================================================

def _get_subject_ref_for_month(record: "KarteiRecord", month: int):
    """Return the relevant ``subject<N>_ref`` FK for *month*."""
    semester = get_semester_for_month(month, record.year)
    return record.subject1_ref if semester == 1 else record.subject2_ref


def _get_active_link(subject_ref, year: int):
    """
    Return the *first* active ``SubjectCategoryLink`` (with category
    pre-loaded) for *subject_ref* / *year*, or ``None``.
    """
    return (
        SubjectCategoryLink.objects
        .filter(subject=subject_ref, year=year, category__is_active=True)
        .select_related("category")
        .first()
    )


# =============================================================================
# GROUP suggested price
# =============================================================================

def calculate_suggested_price_group(
    record: "KarteiRecord",
    month: int,
    *,
    link: SubjectCategoryLink | None = None,
    group: DisciplineGroup | None = None,
    duration_entries: list[DurationEntry] | None = None,
    contract_type_entries: list | None = None,
) -> dict | None:
    """
    Calculate the suggested price for a GROUP subject in *month*.

    Returns a dict::

        {
            "price": Decimal,
            "rate": Decimal,
            "contract_type": "monthly" | "yearly",
            "duration_minutes": int,
            "ue_count": Decimal,
            "base_price": Decimal,
            "size": int,
            "threshold": int,
            "scaling_applied": bool,
            "category_name": str,
            "category_kind": "GROUP",
        }

    or ``None`` when data is insufficient (no link, no group, no duration).
    """
    subject_ref = _get_subject_ref_for_month(record, month)
    if subject_ref is None:
        return None

    # --- active category link ---
    if link is None:
        link = _get_active_link(subject_ref, record.year)
    if link is None:
        return None

    category = link.category
    if category.kind != CategoryKind.GROUP:
        return None

    # --- discipline group ---
    if group is None:
        group = (
            DisciplineGroup.objects
            .filter(subject=subject_ref, year=record.year, is_active=True)
            .first()
        )
    if group is None:
        return None

    # --- contract type → rate ---
    is_monthly = get_contract_type_for_month(
        record, month, entries=contract_type_entries,
    )
    rate = category.monthly_rate if is_monthly else category.yearly_rate

    # --- duration ---
    duration_minutes = get_duration_for_month(
        group, month, entries=duration_entries,
    )
    if duration_minutes is None:
        return None

    # --- UE count (1 UE = 45 min) ---
    ue_count = Decimal(duration_minutes) / Decimal(45)

    # --- base price ---
    base_price = rate * ue_count

    # --- scaling ---
    size_info = get_group_size_for_month(group, month)
    size = size_info["size"]
    threshold = category.group_threshold
    scaling_applied = False

    if group.auto_scaling_enabled and size > 0 and size < threshold:
        price = base_price * Decimal(threshold) / Decimal(size)
        scaling_applied = True
    else:
        price = base_price

    price = round_money_up(price)

    return {
        "price": price,
        "rate": rate,
        "contract_type": "monthly" if is_monthly else "yearly",
        "duration_minutes": duration_minutes,
        "ue_count": ue_count,
        "base_price": round_money_up(base_price),
        "size": size,
        "threshold": threshold,
        "scaling_applied": scaling_applied,
        "category_name": category.name,
        "category_kind": CategoryKind.GROUP,
    }


# =============================================================================
# INDIVIDUAL suggested price
# =============================================================================

def calculate_suggested_price_individual(
    record: "KarteiRecord",
    month: int,
    *,
    link: SubjectCategoryLink | None = None,
    contract_type_entries: list | None = None,
) -> dict | None:
    """
    Calculate the suggested price for an INDIVIDUAL subject in *month*.

    Returns a dict::

        {
            "price": Decimal,
            "rate": Decimal,
            "contract_type": "monthly" | "yearly",
            "hours": Decimal,
            "category_name": str,
            "category_kind": "INDIVIDUAL",
        }

    or ``None`` when data is insufficient.
    """
    subject_ref = _get_subject_ref_for_month(record, month)
    if subject_ref is None:
        return None

    # --- active category link ---
    if link is None:
        link = _get_active_link(subject_ref, record.year)
    if link is None:
        return None

    category = link.category
    if category.kind != CategoryKind.INDIVIDUAL:
        return None

    # --- contract type → rate ---
    is_monthly = get_contract_type_for_month(
        record, month, entries=contract_type_entries,
    )
    rate = category.monthly_rate if is_monthly else category.yearly_rate

    # --- hours ---
    hours = Decimal(str(record.hours_amounts.get(f"month_{month}", 0) or 0))

    price = round_money_up(hours * rate)

    return {
        "price": price,
        "rate": rate,
        "contract_type": "monthly" if is_monthly else "yearly",
        "hours": hours,
        "category_name": category.name,
        "category_kind": CategoryKind.INDIVIDUAL,
    }


# =============================================================================
# Unified wrapper
# =============================================================================

def calculate_suggested_price(
    record: "KarteiRecord",
    month: int,
    *,
    link: SubjectCategoryLink | None = None,
    group: DisciplineGroup | None = None,
    duration_entries: list[DurationEntry] | None = None,
    contract_type_entries: list | None = None,
) -> dict | None:
    """
    Calculate the suggested price for *record* in *month*.

    Returns ``None`` when:
    - pricing source is ``PRICE_OPTION`` (no category link), or
    - required data (group, duration, etc.) is missing.
    """
    semester = get_semester_for_month(month, record.year)
    source = determine_pricing_source(record, semester)

    if source is PricingSource.PRICE_OPTION:
        return None

    # Resolve link once (if not supplied by caller)
    if link is None:
        subject_ref = _get_subject_ref_for_month(record, month)
        if subject_ref is None:
            return None
        link = _get_active_link(subject_ref, record.year)
    if link is None:
        return None

    category = link.category

    if category.kind == CategoryKind.GROUP:
        return calculate_suggested_price_group(
            record,
            month,
            link=link,
            group=group,
            duration_entries=duration_entries,
            contract_type_entries=contract_type_entries,
        )
    elif category.kind == CategoryKind.INDIVIDUAL:
        return calculate_suggested_price_individual(
            record,
            month,
            link=link,
            contract_type_entries=contract_type_entries,
        )

    return None


# =============================================================================
# Monthly map (prefetched, no N+1)
# =============================================================================

def get_suggested_prices_for_record(record: "KarteiRecord") -> dict[int, dict | None]:
    """
    Return ``{1: info|None, 2: …, 12: info|None}`` with the suggested
    price for every month of the year.

    Prefetches links, categories, groups, duration entries, and contract-type
    entries **once** to avoid per-month DB round-trips.
    """
    result: dict[int, dict | None] = {}

    # --- prefetch contract-type entries (shared across months) ---
    contract_type_entries = list(record.contract_type_entries.all())

    # --- per-semester prefetch ---
    _cache: dict[int, dict] = {}  # semester → {link, group, duration_entries}

    for semester in (1, 2):
        subject_ref = record.subject1_ref if semester == 1 else record.subject2_ref
        if subject_ref is None:
            _cache[semester] = {"link": None, "group": None, "duration_entries": None}
            continue

        link = _get_active_link(subject_ref, record.year)
        group = None
        duration_entries = None

        if link is not None and link.category.kind == CategoryKind.GROUP:
            group = (
                DisciplineGroup.objects
                .filter(subject=subject_ref, year=record.year, is_active=True)
                .first()
            )
            if group is not None:
                duration_entries = list(group.duration_entries.all())

        _cache[semester] = {
            "link": link,
            "group": group,
            "duration_entries": duration_entries,
        }

    # --- iterate months ---
    for month in range(1, 13):
        semester = get_semester_for_month(month, record.year)
        ctx = _cache[semester]

        result[month] = calculate_suggested_price(
            record,
            month,
            link=ctx["link"],
            group=ctx["group"],
            duration_entries=ctx["duration_entries"],
            contract_type_entries=contract_type_entries,
        )

    return result
