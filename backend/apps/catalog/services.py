"""
Catalog bootstrap services.

ensure_default_categories(year) creates the default SubjectCategory rows
for a given year when none exist yet (first-open bootstrap).
"""

from decimal import Decimal

from .models import CategoryKind, SubjectCategory

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
