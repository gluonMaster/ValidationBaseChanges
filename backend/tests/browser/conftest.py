from __future__ import annotations

from collections.abc import Callable
from itertools import count
from typing import Any

import pytest
from django.contrib.auth import get_user_model

from apps.accounts.models import UserRole
from apps.approvals.models import DeclinedChange, PendingChange
from apps.approvals.services import build_snapshot
from apps.karteien.models import KarteiRecord, RecordStatus


@pytest.fixture(autouse=True)
def allow_live_server_host(settings):
    settings.ALLOWED_HOSTS = [
        *settings.ALLOWED_HOSTS,
        "localhost",
        "127.0.0.1",
        "[::1]",
        "testserver",
    ]


@pytest.fixture
def browser_user_password() -> str:
    return "browser-smoke-password"


@pytest.fixture
def browser_user_builder(
    transactional_db,
    browser_user_password: str,
) -> Callable[..., Any]:
    sequence = count(1)

    def build_user(
        *,
        role: str,
        username: str | None = None,
        password: str | None = None,
        **overrides,
    ):
        user_number = next(sequence)
        user_model = get_user_model()
        username = username or f"browser_{role.lower()}_{user_number}"
        return user_model.objects.create_user(
            username=username,
            password=password or browser_user_password,
            role=role,
            **overrides,
        )

    return build_user


@pytest.fixture
def browser_users(browser_user_builder):
    return {
        UserRole.ADMIN: browser_user_builder(
            role=UserRole.ADMIN,
            username="browser_admin",
        ),
        UserRole.OPERATOR: browser_user_builder(
            role=UserRole.OPERATOR,
            username="browser_operator",
        ),
        UserRole.SUPERADMIN: browser_user_builder(
            role=UserRole.SUPERADMIN,
            username="browser_superadmin",
        ),
        UserRole.USER: browser_user_builder(
            role=UserRole.USER,
            username="browser_user",
        ),
    }


@pytest.fixture
def browser_record_builder(transactional_db) -> Callable[..., KarteiRecord]:
    sequence = count(9000)

    def build_record(**overrides) -> KarteiRecord:
        access_id = overrides.pop("id", next(sequence))
        year = overrides.pop("year", 2026)
        defaults = {
            "id": access_id,
            "year": year,
            "family_id": f"SMOKE-{year}-{access_id}",
            "parent_name": "Browser Smoke Parent",
            "child_name": "Browser Smoke Child",
            "status": RecordStatus.NORMAL,
        }
        defaults.update(overrides)
        return KarteiRecord.objects.create(**defaults)

    return build_record


@pytest.fixture
def browser_normal_record(browser_record_builder):
    return browser_record_builder()


@pytest.fixture
def browser_pending_change(browser_record_builder):
    record = browser_record_builder(
        id=9101,
        family_id="SMOKE-PENDING",
        child_name="Browser Pending Original",
        status=RecordStatus.PENDING,
    )
    snapshot = build_snapshot(record)
    snapshot["child_name"] = "Browser Pending Proposed"
    return PendingChange.objects.create(
        record=record,
        snapshot=snapshot,
        admin_comment="Browser smoke pending change",
    )


@pytest.fixture
def browser_declined_record(browser_record_builder):
    record = browser_record_builder(
        id=9102,
        family_id="SMOKE-DECLINED",
        child_name="Browser Declined Child",
        status=RecordStatus.DECLINED,
    )
    DeclinedChange.objects.create(
        record=record,
        snapshot=build_snapshot(record),
        decline_reason="Browser smoke declined change",
    )
    return record


@pytest.fixture
def browser_restricted_records(browser_pending_change, browser_declined_record):
    return [browser_pending_change.record, browser_declined_record]


@pytest.fixture
def browser_page(request):
    playwright_api = pytest.importorskip(
        "playwright.sync_api",
        reason=(
            "playwright is not installed; run "
            "`pip install -r backend/requirements.txt`."
        ),
    )
    browser_name = _selected_browser_name(request)
    playwright = None

    try:
        playwright = playwright_api.sync_playwright().start()
        browser_type = getattr(playwright, browser_name)
        browser = browser_type.launch()
    except AttributeError:
        if playwright is not None:
            playwright.stop()
        pytest.skip(f"Unsupported Playwright browser: {browser_name}")
    except playwright_api.Error as exc:
        message = str(exc)
        if (
            "Executable doesn't exist" in message
            or "playwright install" in message
            or "Playwright was just installed" in message
        ):
            if playwright is not None:
                playwright.stop()
            pytest.skip(
                "Playwright browser binary is not installed; run "
                "`python -m playwright install chromium`."
            )
        if playwright is not None:
            playwright.stop()
        raise

    context = browser.new_context()
    page = context.new_page()
    try:
        yield page
    finally:
        context.close()
        browser.close()
        if playwright is not None:
            playwright.stop()


def _selected_browser_name(request) -> str:
    try:
        value = request.config.getoption("--browser")
    except Exception as exc:
        if exc.__class__.__name__ == "ValueError":
            return "chromium"
        raise
    if isinstance(value, (list, tuple)):
        return value[0] if value else "chromium"
    return value or "chromium"
