"""
Catalog bootstrap services.

ensure_default_categories(year) creates the default SubjectCategory rows
for a given year when none exist yet (first-open bootstrap).

copy_categories_between_years(source_year, target_year) copies categories,
links, and discipline groups from one year to another.
"""

from __future__ import annotations

import dataclasses
from decimal import Decimal
from typing import List

from django.db import transaction

from .models import (
    CategoryKind,
    DisciplineGroup,
    SubjectCategory,
    SubjectCategoryLink,
)

DEFAULT_GROUP_CATEGORIES = [
    "Russisch",
    "Deutsch",
    "Englisch",
    "Schach",
    "Tanzen",
    "Kunst",
    "Musik",
    "Logorhythmik",
]

DEFAULT_INDIVIDUAL_CATEGORIES = [
    "Nachhilfe",
    "Logopädie",
    "Einzelunterricht",
]


def ensure_default_categories(year: int) -> bool:
    """
    Create default SubjectCategory rows for the given year if none exist.

    Returns True if created, False if already existed.
    """
    if SubjectCategory.objects.filter(year=year).exists():
        return False

    categories = []
    for name in DEFAULT_GROUP_CATEGORIES:
        categories.append(
            SubjectCategory(
                year=year,
                name=name,
                kind=CategoryKind.GROUP,
                yearly_rate=Decimal("0.00"),
                monthly_rate=Decimal("0.00"),
            )
        )
    for name in DEFAULT_INDIVIDUAL_CATEGORIES:
        categories.append(
            SubjectCategory(
                year=year,
                name=name,
                kind=CategoryKind.INDIVIDUAL,
                yearly_rate=Decimal("0.00"),
                monthly_rate=Decimal("0.00"),
            )
        )

    SubjectCategory.objects.bulk_create(categories)
    return True


# ---------------------------------------------------------------------------
# Copy categories between years
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class CopyCategoriesResult:
    """Statistics returned by ``copy_categories_between_years``."""

    categories_created: int = 0
    categories_skipped: int = 0
    links_created: int = 0
    links_skipped: int = 0
    groups_created: int = 0
    warnings: List[str] = dataclasses.field(default_factory=list)


@transaction.atomic
def copy_categories_between_years(
    source_year: int,
    target_year: int,
    *,
    overwrite: bool = False,
) -> CopyCategoriesResult:
    """
    Copy active ``SubjectCategory`` rows (with links & groups) from
    *source_year* to *target_year*.

    * ``DurationEntry`` / ``GroupSizeEntry`` / ``Contract*Entry`` are **not**
      copied — they are year-specific.
    * Only categories with ``is_active=True`` in the source year are copied.

    Parameters
    ----------
    source_year:
        The year to copy from.
    target_year:
        The year to copy into.
    overwrite:
        When *False* (default), refuse to copy if *target_year* already has
        categories and return a warning instead.  When *True*, deactivate
        existing categories in *target_year* and delete their links before
        copying.

    Returns
    -------
    CopyCategoriesResult
        Dataclass with counters and warnings.
    """
    result = CopyCategoriesResult()

    # --- validation ----------------------------------------------------------
    if source_year == target_year:
        raise ValueError("source_year and target_year must be different.")

    target_has_categories = SubjectCategory.objects.filter(
        year=target_year,
    ).exists()

    if target_has_categories and not overwrite:
        result.warnings.append(
            f"Target year {target_year} already contains categories. "
            "Use overwrite=True to replace them."
        )
        return result

    # --- overwrite: deactivate & clean up ------------------------------------
    if target_has_categories and overwrite:
        SubjectCategoryLink.objects.filter(year=target_year).delete()
        SubjectCategory.objects.filter(year=target_year).update(is_active=False)

    # --- fetch source data ---------------------------------------------------
    source_categories = (
        SubjectCategory.objects.filter(year=source_year, is_active=True)
        .order_by("pk")
    )

    if not source_categories.exists():
        result.warnings.append(
            f"No active categories found in source year {source_year}."
        )
        return result

    # Pre-load source links keyed by source category id
    source_links_qs = SubjectCategoryLink.objects.filter(
        year=source_year,
    ).select_related("subject", "category")

    source_links_by_cat: dict[int, list[SubjectCategoryLink]] = {}
    for link in source_links_qs:
        source_links_by_cat.setdefault(link.category_id, []).append(link)

    # Track subjects already linked in target_year (may exist from earlier
    # categories created in this same call, or from non-overwritten data).
    linked_subjects_in_target: set[int] = set(
        SubjectCategoryLink.objects.filter(year=target_year).values_list(
            "subject_id", flat=True
        )
    )

    # --- copy each category --------------------------------------------------
    for src_cat in source_categories:
        # Create new category in target year
        new_cat = SubjectCategory.objects.create(
            year=target_year,
            name=src_cat.name,
            kind=src_cat.kind,
            yearly_rate=src_cat.yearly_rate,
            monthly_rate=src_cat.monthly_rate,
            group_threshold=src_cat.group_threshold,
            is_active=True,
        )
        result.categories_created += 1

        # Copy links
        for src_link in source_links_by_cat.get(src_cat.pk, []):
            subject = src_link.subject

            if subject.pk in linked_subjects_in_target:
                result.links_skipped += 1
                result.warnings.append(
                    f"Subject '{subject.name}' (pk={subject.pk}) already "
                    f"linked in {target_year} — skipped."
                )
                continue

            SubjectCategoryLink.objects.create(
                subject=subject,
                category=new_cat,
                # year is set automatically in save()
            )
            linked_subjects_in_target.add(subject.pk)
            result.links_created += 1

            # For GROUP categories, ensure a DisciplineGroup exists
            if new_cat.kind == CategoryKind.GROUP:
                dg, created = DisciplineGroup.objects.update_or_create(
                    subject=subject,
                    year=target_year,
                    defaults={
                        "category": new_cat,
                        "is_active": True,
                    },
                )
                if created:
                    result.groups_created += 1

    return result
