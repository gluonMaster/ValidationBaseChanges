"""
URL configuration for the reporting app.

API Endpoints:
- GET /api/reporting/recent-changes/ — list recent changes across all records
- GET /api/reporting/recent-changes-by-record/ — list changes grouped by record
"""

from django.urls import path

from apps.reporting.views import RecentChangesByRecordView, RecentChangesView

app_name = "reporting"

urlpatterns = [
    # Recent changes endpoints
    path(
        "recent-changes/",
        RecentChangesView.as_view(),
        name="recent-changes",
    ),
    path(
        "recent-changes-by-record/",
        RecentChangesByRecordView.as_view(),
        name="recent-changes-by-record",
    ),
]
