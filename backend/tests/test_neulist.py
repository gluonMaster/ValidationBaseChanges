from __future__ import annotations

import pytest

from apps.approvals.models import SuperadminState
from apps.approvals.services import (
    _get_last_seen_id_for_year,
    get_new_records,
    update_last_seen_id,
)


def test_last_seen_lookup_uses_year_mapping_without_legacy_fallback():
    state = SuperadminState(
        last_seen_id=999,
        last_seen_by_year={"2025": 12},
    )

    assert _get_last_seen_id_for_year(state, 2025) == 12
    assert _get_last_seen_id_for_year(state, 2026) == 0


@pytest.mark.django_db
def test_get_new_records_filters_by_per_year_last_seen(
    superadmin_user,
    kartei_record_builder,
):
    SuperadminState.objects.create(
        user=superadmin_user,
        last_seen_id=999,
        last_seen_by_year={"2025": 100},
    )
    kartei_record_builder(year=2025, id=100)
    new_2025 = kartei_record_builder(year=2025, id=101)
    new_2026 = kartei_record_builder(year=2026, id=1)

    assert get_new_records(superadmin_user, 2025) == [new_2025]
    assert get_new_records(superadmin_user, 2026) == [new_2026]


@pytest.mark.django_db
def test_update_last_seen_id_updates_year_bucket_and_legacy_field(
    superadmin_user,
    kartei_record_builder,
):
    kartei_record_builder(year=2025, id=50)
    kartei_record_builder(year=2025, id=60)
    kartei_record_builder(year=2026, id=10)

    update_last_seen_id(superadmin_user, year=2025)

    state = SuperadminState.objects.get(user=superadmin_user)
    assert state.last_seen_by_year == {"2025": 60}
    assert state.last_seen_id == 60
