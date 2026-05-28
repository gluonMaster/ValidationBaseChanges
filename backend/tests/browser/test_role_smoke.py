from __future__ import annotations

import re

import pytest
from django.urls import reverse

from apps.accounts.models import UserRole


pytestmark = [
    pytest.mark.browser,
    pytest.mark.django_db(transaction=True),
]


ROLE_ROOT_PATHS = [
    (UserRole.ADMIN, "/karteien/"),
    (UserRole.OPERATOR, "/karteien/"),
    (UserRole.SUPERADMIN, "/approvals/superadmin/pending/"),
    (UserRole.USER, "/user/"),
]


def _url(live_server, path: str) -> str:
    return f"{live_server.url}{path}"


def _login(page, live_server, user, password: str) -> None:
    page.goto(_url(live_server, reverse("login")), wait_until="domcontentloaded")
    page.locator("#id_username").fill(user.username)
    page.locator("#id_password").fill(password)
    page.locator("form button[type='submit']").click()
    page.wait_for_load_state("domcontentloaded")


def _expect_path(page, live_server, path: str) -> None:
    expected = re.compile(rf"^{re.escape(live_server.url)}{re.escape(path)}(?:[?#].*)?$")
    page.wait_for_url(expected)


@pytest.mark.parametrize(("role", "expected_path"), ROLE_ROOT_PATHS)
def test_login_and_root_redirect_per_role(
    live_server,
    browser_users,
    browser_user_password,
    role,
    expected_path,
    browser_page,
):
    user = browser_users[role]

    _login(browser_page, live_server, user, browser_user_password)
    browser_page.goto(_url(live_server, "/"), wait_until="domcontentloaded")

    _expect_path(browser_page, live_server, expected_path)


def test_admin_sees_kartei_create_and_edit_entry_points(
    live_server,
    browser_users,
    browser_user_password,
    browser_normal_record,
    browser_page,
):
    _login(
        browser_page,
        live_server,
        browser_users[UserRole.ADMIN],
        browser_user_password,
    )

    browser_page.goto(
        _url(live_server, f"/karteien/?year={browser_normal_record.year}"),
        wait_until="domcontentloaded",
    )
    browser_page.locator("#btn-new-record").wait_for(state="visible")
    browser_page.locator(
        f"a[href='/karteien/{browser_normal_record.pk}/edit/']"
    ).first.wait_for(state="visible")

    browser_page.goto(
        _url(live_server, f"/karteien/create/?year={browser_normal_record.year}"),
        wait_until="domcontentloaded",
    )
    browser_page.locator("input[name='family_id']").wait_for(state="visible")

    browser_page.goto(
        _url(live_server, f"/karteien/{browser_normal_record.pk}/edit/"),
        wait_until="domcontentloaded",
    )
    browser_page.locator("input[name='family_id']").wait_for(state="visible")


def test_operator_standard_editor_redirects_for_pending_and_declined_records(
    live_server,
    browser_users,
    browser_user_password,
    browser_restricted_records,
    browser_page,
):
    _login(
        browser_page,
        live_server,
        browser_users[UserRole.OPERATOR],
        browser_user_password,
    )

    for record in browser_restricted_records:
        browser_page.goto(
            _url(live_server, f"/karteien/{record.pk}/edit/"),
            wait_until="domcontentloaded",
        )

        _expect_path(browser_page, live_server, f"/karteien/{record.pk}/")


def test_superadmin_opens_pending_overview_and_war_ist_detail(
    live_server,
    browser_users,
    browser_user_password,
    browser_pending_change,
    browser_page,
):
    _login(
        browser_page,
        live_server,
        browser_users[UserRole.SUPERADMIN],
        browser_user_password,
    )

    browser_page.goto(
        _url(
            live_server,
            f"/approvals/superadmin/pending/?year={browser_pending_change.record.year}",
        ),
        wait_until="domcontentloaded",
    )
    browser_page.locator(
        f"a[href='/approvals/superadmin/pending/{browser_pending_change.pk}/']"
    ).wait_for(state="visible")

    browser_page.goto(
        _url(live_server, f"/approvals/superadmin/pending/{browser_pending_change.pk}/"),
        wait_until="domcontentloaded",
    )
    browser_page.locator("input[name='decision'][value='APPROVED']").wait_for(
        state="attached"
    )
    browser_page.locator("input[name='decision'][value='DECLINED']").wait_for(
        state="attached"
    )


def test_user_cabinet_and_record_route_are_read_only(
    live_server,
    browser_users,
    browser_user_password,
    browser_normal_record,
    browser_page,
):
    _login(
        browser_page,
        live_server,
        browser_users[UserRole.USER],
        browser_user_password,
    )

    browser_page.goto(_url(live_server, "/"), wait_until="domcontentloaded")
    _expect_path(browser_page, live_server, "/user/")
    browser_page.locator("form[action='/user/search/']").wait_for(state="visible")

    browser_page.goto(
        _url(live_server, f"/user/record/{browser_normal_record.pk}/"),
        wait_until="domcontentloaded",
    )
    browser_page.get_by_text(
        browser_normal_record.child_name,
        exact=True,
    ).first.wait_for(state="visible")
    assert browser_page.locator(
        f"a[href='/karteien/{browser_normal_record.pk}/edit/']"
    ).count() == 0
    assert browser_page.locator("a[href^='/karteien/create/']").count() == 0
