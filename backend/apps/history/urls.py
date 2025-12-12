"""
URL configuration for the history app.

API Endpoints:
- GET /api/history/records/<id>/ — list history events for a record
- GET /api/history/events/<id>/ — get a single history event
- POST /api/history/records/<id>/sync/ — sync history from raw
"""

from django.urls import path

from apps.history.views import (
    HistoryEventDetailView,
    RecordHistoryView,
    SyncHistoryView,
)

app_name = "history"

urlpatterns = [
    # Record history endpoints
    path(
        "records/<int:record_id>/",
        RecordHistoryView.as_view(),
        name="record-history",
    ),
    path(
        "records/<int:record_id>/sync/",
        SyncHistoryView.as_view(),
        name="sync-history",
    ),
    
    # Single event endpoint
    path(
        "events/<int:event_id>/",
        HistoryEventDetailView.as_view(),
        name="event-detail",
    ),
]
