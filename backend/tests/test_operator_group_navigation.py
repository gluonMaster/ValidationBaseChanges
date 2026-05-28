from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest
from django.urls import reverse

from apps.catalog.models import (
    CategoryKind,
    DisciplineGroup,
    DurationEntry,
    Subject,
    SubjectCategory,
    SubjectCategoryLink,
)


pytestmark = pytest.mark.django_db


def _create_group(*, year: int | None = None) -> DisciplineGroup:
    year = year or date.today().year
    subject = Subject.objects.create(name=f"Prompt162 Group {year}")
    category = SubjectCategory.objects.create(
        year=year,
        name=f"Prompt162 Group Category {year}",
        kind=CategoryKind.GROUP,
        yearly_rate=Decimal("45.00"),
        monthly_rate=Decimal("60.00"),
        group_threshold=6,
    )
    SubjectCategoryLink.objects.create(subject=subject, category=category)
    group = DisciplineGroup.objects.create(
        subject=subject,
        year=year,
        category=category,
        auto_scaling_enabled=False,
    )
    DurationEntry.objects.create(
        group=group,
        effective_from_month=1,
        duration_minutes=45,
    )
    return group


def _html(response) -> str:
    return response.content.decode(response.charset or "utf-8")


def test_operator_can_navigate_group_bulk_flow_without_admin_only_controls(
    client,
    operator_user,
):
    group = _create_group()
    year = group.year
    group_list_url = reverse("catalog:group_list", kwargs={"year": year})
    group_detail_url = reverse(
        "catalog:group_detail",
        kwargs={"year": year, "pk": group.pk},
    )
    preview_url = reverse(
        "catalog:group_apply_preview",
        kwargs={"year": year, "pk": group.pk},
    )

    client.force_login(operator_user)

    response = client.get(group_list_url)
    assert response.status_code == 200
    html = _html(response)
    assert f'href="{group_list_url}"' in html
    assert f'href="{group_detail_url}"' in html
    assert 'href="/catalog/"' not in html
    assert reverse(
        "catalog:group_toggle_scaling",
        kwargs={"year": year, "pk": group.pk},
    ) not in html

    response = client.get(group_detail_url)
    assert response.status_code == 200
    html = _html(response)
    assert f'href="{group_list_url}"' in html
    assert f'action="{preview_url}"' in html
    assert "Massenaktualisierung" in html
    assert 'href="/catalog/"' not in html
    for admin_only_url_name in (
        "group_toggle_scaling",
        "group_duration_add",
        "group_size_add",
        "group_prepare_legacy",
    ):
        assert reverse(
            f"catalog:{admin_only_url_name}",
            kwargs={"year": year, "pk": group.pk},
        ) not in html

    response = client.get(
        preview_url,
        {
            "semester": "1",
            "from_month": "1",
            "comment": "Prompt 162 operator preview",
        },
    )
    assert response.status_code == 200
    html = _html(response)
    assert f'href="{group_list_url}"' in html
    assert f'href="{group_detail_url}"' in html
    assert 'href="/catalog/"' not in html


def test_admin_group_pages_keep_catalog_navigation_and_admin_controls(
    client,
    admin_user,
):
    group = _create_group()
    year = group.year

    client.force_login(admin_user)

    response = client.get(reverse("catalog:group_list", kwargs={"year": year}))
    assert response.status_code == 200
    html = _html(response)
    assert 'id="katalogDropdown"' in html
    assert 'href="/catalog/"' in html
    assert reverse(
        "catalog:group_toggle_scaling",
        kwargs={"year": year, "pk": group.pk},
    ) in html

    response = client.get(
        reverse("catalog:group_detail", kwargs={"year": year, "pk": group.pk}),
    )
    assert response.status_code == 200
    html = _html(response)
    for admin_only_url_name in (
        "group_toggle_scaling",
        "group_duration_add",
        "group_size_add",
        "group_prepare_legacy",
    ):
        assert reverse(
            f"catalog:{admin_only_url_name}",
            kwargs={"year": year, "pk": group.pk},
        ) in html


def test_operator_still_cannot_open_admin_catalog_index(client, operator_user):
    client.force_login(operator_user)

    response = client.get(reverse("catalog:index"))

    assert response.status_code == 302
    assert response.url == reverse("karteien:record_list")
