"""
Warnings service — read-only computation of pricing/config warnings.

Public API
----------
* ``PricingWarning``        — dataclass carrying code, message, severity, context
* ``get_group_warnings(group)``
    → list[PricingWarning]  — warnings related to a DisciplineGroup
* ``get_record_pricing_warnings(record)``
    → list[PricingWarning]  — warnings related to a KarteiRecord

The module never writes to the database.  It is purely a diagnostic calculator.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import TYPE_CHECKING

from apps.catalog.group_size_service import get_group_size_for_month
from apps.catalog.models import (
    DurationEntry,
    SubjectCategoryLink,
    get_duration_for_month,
)
from apps.catalog.pricing import get_suggested_prices_for_record
from apps.karteien.billing import get_semester_for_month

if TYPE_CHECKING:
    from apps.catalog.models import DisciplineGroup
    from apps.karteien.models import KarteiRecord


# ---------------------------------------------------------------------------
# Data class
# ---------------------------------------------------------------------------

@dataclass
class PricingWarning:
    code: str
    message_de: str
    severity: str  # "error" | "warning" | "info"
    context: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Group warnings
# ---------------------------------------------------------------------------

def get_group_warnings(group: "DisciplineGroup") -> list[PricingWarning]:
    """Return warnings for a DisciplineGroup (read-only, no side-effects)."""
    warnings: list[PricingWarning] = []

    duration_entries = list(group.duration_entries.all())

    # --- NO_DURATION: no DurationEntry at all ---------------------------------
    if not duration_entries:
        warnings.append(PricingWarning(
            code="NO_DURATION",
            message_de="Länge nicht festgelegt — Preisberechnung nicht möglich.",
            severity="error",
            context={},
        ))

    # Per-month checks (months 1–12)
    for month in range(1, 13):
        duration = get_duration_for_month(group, month, entries=duration_entries)
        size_info = get_group_size_for_month(group, month)
        size = size_info["size"]
        billable = size_info["billable_count"]

        # --- DURATION_GAP: students present but no duration -------------------
        if duration is None and size > 0:
            warnings.append(PricingWarning(
                code="DURATION_GAP",
                message_de=(
                    f"Monat {month}: Schüler vorhanden, aber keine Länge festgelegt."
                ),
                severity="error",
                context={"month": month, "size": size},
            ))

        # --- SCALING_EMPTY: auto-scaling enabled but size == 0 ----------------
        if group.auto_scaling_enabled and size == 0:
            warnings.append(PricingWarning(
                code="SCALING_EMPTY",
                message_de=(
                    f"Monat {month}: Auto-Skalierung ist aktiv, aber die Gruppe "
                    f"ist leer — Skalierung nicht möglich."
                ),
                severity="warning",
                context={"month": month},
            ))

        # --- NO_BILLABLE: size > 0 but nobody billable -----------------------
        if size > 0 and billable == 0:
            warnings.append(PricingWarning(
                code="NO_BILLABLE",
                message_de=(
                    f"Monat {month}: Keine zahlenden Schüler "
                    f"(alle pausiert) — Berechnung ergibt 0."
                ),
                severity="warning",
                context={"month": month, "size": size},
            ))

    # --- EMPTY_GROUP: auto total_size == 0 in every month ---------------------
    all_empty = all(
        get_group_size_for_month(group, m)["auto_size"] == 0
        for m in range(1, 13)
    )
    if all_empty:
        warnings.append(PricingWarning(
            code="EMPTY_GROUP",
            message_de="Keine Schüler in der Gruppe (ACTIVE/PAUSED) in allen Monaten.",
            severity="warning",
            context={},
        ))

    return warnings


# ---------------------------------------------------------------------------
# Record warnings
# ---------------------------------------------------------------------------

def _has_category_link(subject_ref, year: int) -> bool:
    """True if subject has an active SubjectCategoryLink for the given year."""
    if subject_ref is None:
        return False
    return SubjectCategoryLink.objects.filter(
        subject=subject_ref, year=year, category__is_active=True,
    ).exists()


def _has_price_option(record: "KarteiRecord", semester: int) -> bool:
    """True if the semester slot has a PriceOption FK set."""
    ref = record.price1_ref if semester == 1 else record.price2_ref
    return ref is not None


def get_record_price_mismatches(
    record: "KarteiRecord",
) -> dict[int, dict[str, Decimal | None]]:
    """
    Return month-level price mismatches for category-based pricing.

    Keys are month numbers (1..12). Values contain:
    - ``suggested``: suggested category price
    - ``stored``: stored base amount (None if missing in ``base_amounts``)
    """
    suggested_map = get_suggested_prices_for_record(record)
    base_amounts: dict = record.base_amounts or {}
    mismatches: dict[int, dict[str, Decimal | None]] = {}

    for month in range(1, 13):
        suggested = suggested_map.get(month)
        if suggested is None:
            continue

        suggested_price = suggested["price"]
        stored_raw = base_amounts.get(f"month_{month}")
        stored_price = None
        if stored_raw is not None:
            try:
                stored_price = Decimal(str(stored_raw))
            except (ValueError, TypeError):
                stored_price = Decimal("0")

        if stored_price != suggested_price:
            mismatches[month] = {
                "stored": stored_price,
                "suggested": suggested_price,
            }

    return mismatches


def get_record_pricing_warnings(record: "KarteiRecord") -> list[PricingWarning]:
    """Return pricing warnings for a KarteiRecord (read-only, no side-effects)."""
    from apps.karteien.models import MonthsMode

    warnings: list[PricingWarning] = []

    # --- LEGACY_MODE ----------------------------------------------------------
    if record.months_mode == MonthsMode.LEGACY:
        warnings.append(PricingWarning(
            code="LEGACY_MODE",
            message_de=(
                "Abrechnungsmodus ist LEGACY — bitte zuerst auf AUTO "
                "umstellen (Monate neu berechnen)."
            ),
            severity="info",
            context={"months_mode": record.months_mode},
        ))

    # --- per-semester checks --------------------------------------------------
    for semester in (1, 2):
        subject_ref = record.subject1_ref if semester == 1 else record.subject2_ref

        # --- LEGACY_NO_REF: no subject*_ref ----------------------------------
        if subject_ref is None:
            # Only warn when there is a legacy text value
            legacy_text = (record.subject1 if semester == 1
                           else record.subject2) or ""
            if legacy_text.strip():
                warnings.append(PricingWarning(
                    code="LEGACY_NO_REF",
                    message_de=(
                        f"Halbjahr {semester}: Bitte subject{semester}_ref "
                        f"ausfüllen, um Kategorie/Gruppe zu aktivieren."
                    ),
                    severity="warning",
                    context={"semester": semester, "legacy_text": legacy_text},
                ))

        # --- NO_PRICING: no PriceOption AND no CategoryLink -------------------
        has_cat = _has_category_link(subject_ref, record.year)
        has_po = _has_price_option(record, semester)
        if subject_ref is not None and not has_cat and not has_po:
            warnings.append(PricingWarning(
                code="NO_PRICING",
                message_de=(
                    f"Halbjahr {semester}: Kein Preis definiert — "
                    f"bitte Kategorie oder PriceOption einrichten."
                ),
                severity="error",
                context={"semester": semester},
            ))

    # --- PRICE_MISMATCH: suggested != stored base_amounts ---------------------
    for month, mismatch in get_record_price_mismatches(record).items():
        suggested_price = mismatch["suggested"]
        stored_price = mismatch["stored"]
        if stored_price is None:
            warnings.append(PricingWarning(
                code="PRICE_MISMATCH",
                message_de=(
                    f"Monat {month}: Vorgeschlagener Preis "
                    f"({suggested_price}) weicht vom gespeicherten Betrag ab."
                ),
                severity="warning",
                context={
                    "month": month,
                    "suggested": str(suggested_price),
                    "stored": None,
                },
            ))
        else:
            warnings.append(PricingWarning(
                code="PRICE_MISMATCH",
                message_de=(
                    f"Monat {month}: Vorgeschlagener Preis "
                    f"({suggested_price}) weicht vom gespeicherten "
                    f"Betrag ({stored_price}) ab."
                ),
                severity="warning",
                context={
                    "month": month,
                    "suggested": str(suggested_price),
                    "stored": str(stored_price),
                },
            ))

    return warnings
