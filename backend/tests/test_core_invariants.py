from __future__ import annotations

import pytest
from django.db import IntegrityError, transaction

from apps.karteien.models import KarteiRecord, RecordStatus


def test_record_status_normal_is_empty_string():
    record = KarteiRecord(id=1, year=2026, family_id="FAM-1")

    assert RecordStatus.NORMAL == ""
    assert RecordStatus.NORMAL.value == ""
    assert not record.status


def test_kartei_record_uses_surrogate_pk_and_domain_key_constraint():
    constraints_by_name = {
        constraint.name: constraint
        for constraint in KarteiRecord._meta.constraints
    }

    assert KarteiRecord._meta.pk.name == "pkid"
    assert not KarteiRecord._meta.get_field("id").primary_key
    assert list(constraints_by_name["unique_kartei_year_id"].fields) == ["year", "id"]


@pytest.mark.django_db
def test_domain_key_is_year_plus_access_id(kartei_record_builder):
    first = kartei_record_builder(year=2025, id=42)
    second = kartei_record_builder(year=2026, id=42)

    assert first.pk != second.pk
    assert (first.year, first.id) == (2025, 42)
    assert (second.year, second.id) == (2026, 42)

    with pytest.raises(IntegrityError):
        with transaction.atomic():
            kartei_record_builder(year=2025, id=42)
