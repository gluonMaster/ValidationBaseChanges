"""
URL configuration for the history app.

API Endpoints:
- GET /api/history/records/<year>/<record_id>/ — list history events for a record
- GET /api/history/events/<id>/ — get a single history event
- POST /api/history/records/<year>/<record_id>/sync/ — sync history from raw

Note: record_id is the Access ID (not globally unique); year + record_id form the domain key.
"""

from django.urls import path

from apps.history.views import (
    HistoryEventDetailView,
    RecordHistoryView,
    SyncHistoryView,
)

app_name = "history"

urlpatterns = [
    # Record history endpoints (using domain key: year + Access ID)
    path(
        "records/<int:year>/<int:record_id>/",
        RecordHistoryView.as_view(),
        name="record-history",
    ),
    path(
        "records/<int:year>/<int:record_id>/sync/",
        SyncHistoryView.as_view(),
        name="sync-history",
    ),
    
    # Single event endpoint (using Django PK)
    path(
        "events/<int:event_id>/",
        HistoryEventDetailView.as_view(),
        name="event-detail",
    ),
]
