"""
Group size service — real-time calculation of effective group size per month.

Two public functions:

* ``calculate_auto_group_size(group, month)``
    Count KarteiRecord entries that belong to the group's subject slot for the
    given month, respecting start/end/csv billing window and contract status.

* ``get_group_size_for_month(group, month)``
    Return effective group size: manual override if present, otherwise auto.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from django.db.models import Q

from apps.karteien.billing import (
    _normalize_subject_name,
    get_semester_for_month,
    get_start_month_for_semester,
    get_end_month_for_semester,
    get_months_csv_for_semester,
)
from apps.karteien.models import (
    KarteiRecord,
    counts_in_group_size,
    is_billable_in_month,
)

if TYPE_CHECKING:
    from .models import DisciplineGroup


# =============================================================================
# Internal helpers
# =============================================================================

def _is_month_active_for_slot(record: KarteiRecord, semester: int, month: int) -> bool:
    """
    Return True if *month* falls within the billing window of the record's
    slot for the given *semester* (start/end month or months_csv).
    """
    months_csv_set = get_months_csv_for_semester(record, semester)
    if months_csv_set is not None:
        return month in months_csv_set

    start = get_start_month_for_semester(record, semester)
    end = get_end_month_for_semester(record, semester)
    return start <= month <= end


# =============================================================================
# Auto-size calculation
# =============================================================================

def calculate_auto_group_size(group: "DisciplineGroup", month: int) -> dict:
    """
    Count KarteiRecord entries that match the group subject for *month*.

    Returns::

        {
            'total_size': int,        # ACTIVE + PAUSED
            'billable_count': int,    # only ACTIVE
            'record_pkids': list[int] # pkids for debugging
        }

    Algorithm (best-effort, no pending projection):

    1. Determine semester from *month* and *group.year*.
    2. Build a queryset of KarteiRecords for that year.
    3. Filter candidates by subject slot (``subject<N>_ref`` FK match, with
       legacy fallback via normalised name comparison).
    4. For every candidate, check:
       a) month is inside the record's billing window (start/end/csv),
       b) ``counts_in_group_size`` → total_size,
       c) ``is_billable_in_month`` → billable_count.
    """
    year = group.year
    subject = group.subject  # catalog.Subject instance
    semester = get_semester_for_month(month, year)

    # Base queryset — all records for this year
    records_qs = KarteiRecord.objects.filter(year=year)

    # ------------------------------------------------------------------
    # Build candidate queryset per semester slot
    # ------------------------------------------------------------------
    if semester == 1:
        ref_field = "subject1_ref_id"
        legacy_field = "subject1"
        ref_null_field = "subject1_ref__isnull"
    else:
        ref_field = "subject2_ref_id"
        legacy_field = "subject2"
        ref_null_field = "subject2_ref__isnull"

    # Direct FK match
    direct_q = Q(**{ref_field: subject.id})

    # Legacy fallback: ref is NULL AND legacy text field is not empty
    legacy_q = Q(**{ref_null_field: True}) & ~Q(**{legacy_field: ""})

    candidates = records_qs.filter(direct_q | legacy_q)

    # Normalised subject name for legacy comparison
    norm_subject = _normalize_subject_name(subject.name)

    # ------------------------------------------------------------------
    # Iterate candidates, applying filters in Python
    # ------------------------------------------------------------------
    total_size = 0
    billable_count = 0
    record_pkids: list[int] = []

    for rec in candidates.iterator():
        # For legacy-only records, verify normalised name match
        ref_value = getattr(rec, ref_field)
        if ref_value is None:
            legacy_value = getattr(rec, legacy_field) or ""
            if _normalize_subject_name(legacy_value) != norm_subject:
                continue

        # Check billing window for the slot
        if not _is_month_active_for_slot(rec, semester, month):
            continue

        # Contract-status checks
        if counts_in_group_size(rec, month):
            total_size += 1
            record_pkids.append(rec.pkid)
            if is_billable_in_month(rec, month):
                billable_count += 1

    return {
        "total_size": total_size,
        "billable_count": billable_count,
        "record_pkids": record_pkids,
    }


# =============================================================================
# Effective size (manual override or auto)
# =============================================================================

def get_group_size_for_month(group: "DisciplineGroup", month: int) -> dict:
    """
    Return effective group size for *month*.

    Returns::

        {
            'size': int,
            'auto_size': int,
            'billable_count': int,
            'is_manual': bool,
            'manual_size': int | None,
        }

    If a manual override exists (via ``get_manual_size_for_month``),
    ``size`` equals the override; otherwise ``size`` equals auto_size.
    """
    from .models import get_manual_size_for_month

    auto = calculate_auto_group_size(group, month)
    manual = get_manual_size_for_month(group, month, year=group.year)

    if manual is not None:
        size = manual
        is_manual = True
    else:
        size = auto["total_size"]
        is_manual = False

    return {
        "size": size,
        "auto_size": auto["total_size"],
        "billable_count": auto["billable_count"],
        "is_manual": is_manual,
        "manual_size": manual,
    }
