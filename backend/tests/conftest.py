from __future__ import annotations

from collections.abc import Callable
from itertools import count
from typing import Any

import pytest
from django.contrib.auth import get_user_model

from apps.accounts.models import UserRole
from apps.karteien.models import KarteiRecord, RecordStatus


@pytest.fixture(autouse=True)
def allow_testserver(settings):
    """Django test client uses testserver as the default host."""
    settings.ALLOWED_HOSTS = [*settings.ALLOWED_HOSTS, "testserver"]


@pytest.fixture
def user_builder(db) -> Callable[..., Any]:
    sequence = count(1)

    def build_user(
        *,
        role: str = UserRole.USER,
        username: str | None = None,
        password: str = "test-password",
        **overrides,
    ):
        user_number = next(sequence)
        user_model = get_user_model()
        username = username or f"{role.lower()}_{user_number}"
        return user_model.objects.create_user(
            username=username,
            password=password,
            role=role,
            **overrides,
        )

    return build_user


@pytest.fixture
def admin_user(user_builder):
    return user_builder(role=UserRole.ADMIN, username="admin_user")


@pytest.fixture
def operator_user(user_builder):
    return user_builder(role=UserRole.OPERATOR, username="operator_user")


@pytest.fixture
def superadmin_user(user_builder):
    return user_builder(role=UserRole.SUPERADMIN, username="superadmin_user")


@pytest.fixture
def kartei_record_builder(db) -> Callable[..., KarteiRecord]:
    sequence = count(1000)

    def build_record(**overrides) -> KarteiRecord:
        access_id = overrides.pop("id", next(sequence))
        year = overrides.pop("year", 2026)
        defaults = {
            "id": access_id,
            "year": year,
            "family_id": f"FAM-{year}-{access_id}",
            "parent_name": "Test Parent",
            "child_name": "Test Child",
            "status": RecordStatus.NORMAL,
        }
        defaults.update(overrides)
        return KarteiRecord.objects.create(**defaults)

    return build_record
