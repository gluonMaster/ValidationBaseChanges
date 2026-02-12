"""
Service: apply category-based prices to a single KarteiRecord.

Public API
----------
* ``apply_category_price_to_record(record, *, semester, from_month)``
    Mutates *record* in-memory (``base_amounts`` + ``month_*`` for affected
    months) and returns a diff dict.  Does **not** call ``record.save()``,
    create PendingChange, or modify ``record.status``.
"""

from __future__ import annotations

from decimal import Decimal
from typing import Any

from apps.catalog.pricing import get_suggested_prices_for_record
from apps.karteien.billing import (
    ZERO,
    calculate_month_values,
    get_end_month_for_semester,
    get_months_csv_for_semester,
    get_semester_month_ranges,
    get_start_month_for_semester,
)
from apps.karteien.models import (
    ContractStatusKind,
    KarteiRecord,
    get_contract_status_for_month,
)


def apply_category_price_to_record(
    record: KarteiRecord,
    *,
    semester: int,
    from_month: int,
) -> dict[str, Any]:
    """
    Apply category-based suggested prices to *record.base_amounts* starting
    from *from_month* within the selected *semester*, respecting start/end/csv
    for that semester.

    Mutates *record* in-memory (``base_amounts`` + ``month_*`` for affected
    months).

    Returns::

        {
            'months_updated': list[int],
            'old_bases':  dict[str, Decimal],
            'new_bases':  dict[str, Decimal],
            'old_months': dict[str, Decimal | None],
            'new_months': dict[str, Decimal],
        }

    Raises:
        ValueError: if no months qualify or a suggested price is unavailable.
    """

    # ------------------------------------------------------------------
    # 1. Determine months_updated
    # ------------------------------------------------------------------
    sem1_months, sem2_months = get_semester_month_ranges(record.year)
    sem_months = sem1_months if semester == 1 else sem2_months

    months_csv = get_months_csv_for_semester(record, semester)
    if months_csv is not None:
        # CSV overrides start/end – intersect with semester range
        active_sem_months = sorted(months_csv & set(sem_months))
    else:
        start = get_start_month_for_semester(record, semester)
        end = get_end_month_for_semester(record, semester)
        active_sem_months = [m for m in sem_months if start <= m <= end]

    months_updated = [m for m in active_sem_months if m >= from_month]

    if not months_updated:
        raise ValueError(
            f"Keine aktiven Monate ab Monat {from_month} "
            f"im {'1.' if semester == 1 else '2.'} Halbjahr."
        )

    # ------------------------------------------------------------------
    # 2. Snapshot old values (before mutation)
    # ------------------------------------------------------------------
    existing_bases: dict[str, str] = record.base_amounts or {}

    old_bases: dict[str, Decimal] = {}
    old_months: dict[str, Decimal | None] = {}
    for m in months_updated:
        key = f"month_{m}"
        raw = existing_bases.get(key)
        old_bases[key] = Decimal(str(raw)) if raw is not None else ZERO
        old_months[key] = getattr(record, key, None)

    # ------------------------------------------------------------------
    # 3. Apply suggested prices → base_amounts
    # ------------------------------------------------------------------
    suggested = get_suggested_prices_for_record(record)

    # Ensure base_amounts is a mutable dict (could be empty on fresh records)
    if not record.base_amounts:
        record.base_amounts = {}

    for m in months_updated:
        info = suggested.get(m)
        if info is None:
            raise ValueError(
                f"Kein Kategoriepreis verfügbar für Monat {m}."
            )
        record.base_amounts[f"month_{m}"] = str(info["price"])

    # ------------------------------------------------------------------
    # 4. Recalculate month_* values via calculate_month_values
    # ------------------------------------------------------------------
    # Build full base_amounts_decimals map (month_1..month_12)
    base_amounts_decimals: dict[str, Decimal] = {}
    for mn in range(1, 13):
        key = f"month_{mn}"
        raw = record.base_amounts.get(key)
        base_amounts_decimals[key] = Decimal(str(raw)) if raw is not None else ZERO

    calculated, _flags = calculate_month_values(
        record, base_amounts=base_amounts_decimals,
    )

    # Prefetch contract-status entries once for the loop below
    if record.pk and hasattr(record, "contract_status_entries"):
        cs_entries: list | None = list(record.contract_status_entries.all())
    else:
        cs_entries = None

    for m in months_updated:
        key = f"month_{m}"
        status = get_contract_status_for_month(record, m, entries=cs_entries)
        if status in (ContractStatusKind.PAUSED, ContractStatusKind.TERMINATED):
            value = Decimal("0.00")
        else:
            value = calculated.get(key, ZERO)
        setattr(record, key, value)

    # ------------------------------------------------------------------
    # 5. Build diff
    # ------------------------------------------------------------------
    new_bases: dict[str, Decimal] = {}
    new_months: dict[str, Decimal] = {}
    for m in months_updated:
        key = f"month_{m}"
        new_bases[key] = Decimal(str(record.base_amounts[key]))
        new_months[key] = getattr(record, key)

    return {
        "months_updated": months_updated,
        "old_bases": old_bases,
        "new_bases": new_bases,
        "old_months": old_months,
        "new_months": new_months,
    }
