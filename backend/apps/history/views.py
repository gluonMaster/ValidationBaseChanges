"""
Views for the history app.

This module contains:
- RecordHistoryView — view history events for a specific KarteiRecord
- HistoryEventDetailView — view a single history event
- SyncHistoryView — trigger history sync from raw for a record

API Endpoints:
- GET /api/history/records/<id>/ — list history events for a record
- GET /api/history/events/<id>/ — get a single history event
- POST /api/history/records/<id>/sync/ — sync history from raw

See ARCHITECTURE.md Section 2.4 for history app overview.
"""

from __future__ import annotations

from django.shortcuts import get_object_or_404
from rest_framework import status
from rest_framework.request import Request
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.history.models import HistoryEvent
from apps.history.serializers import (
    HistoryEventDetailSerializer,
    HistoryEventSerializer,
    RecordHistorySerializer,
)
from apps.history.services import sync_history_from_raw
from apps.karteien.models import KarteiRecord


class RecordHistoryView(APIView):
    """
    View history events for a specific KarteiRecord.
    
    GET /api/history/records/<record_id>/
    
    Returns the record basic info and its history events in chronological order.
    
    Query Parameters:
    - limit: Maximum number of events to return (default: 100)
    - offset: Number of events to skip (default: 0)
    - event_type: Filter by event type (CHANGE, CREATE, APPROVE, DECLINE, IMPORT, RUCK)
    """
    
    def get(self, request: Request, record_id: int) -> Response:
        """Get history events for a record."""
        # Get the record
        record = get_object_or_404(KarteiRecord, id=record_id)
        
        # Build queryset with optional filters
        queryset = HistoryEvent.objects.filter(record=record).order_by("event_time", "id")
        
        # Filter by event_type if provided
        event_type = request.query_params.get("event_type")
        if event_type:
            queryset = queryset.filter(event_type=event_type.upper())
        
        # Get total count before pagination
        total_events = queryset.count()
        
        # Apply pagination
        try:
            limit = int(request.query_params.get("limit", 100))
            offset = int(request.query_params.get("offset", 0))
        except (TypeError, ValueError):
            limit = 100
            offset = 0
        
        limit = min(max(1, limit), 1000)  # Clamp between 1 and 1000
        offset = max(0, offset)
        
        events = queryset[offset:offset + limit]
        
        # Serialize response
        response_data = {
            "id": record.id,
            "year": record.year,
            "family_id": record.family_id,
            "parent_name": record.parent_name,
            "child_name": record.child_name,
            "history_events": HistoryEventSerializer(events, many=True).data,
            "total_events": total_events,
        }
        
        serializer = RecordHistorySerializer(data=response_data)
        serializer.is_valid()
        
        return Response(response_data)


class HistoryEventDetailView(APIView):
    """
    View a single history event.
    
    GET /api/history/events/<event_id>/
    
    Returns detailed information about a history event including record info.
    """
    
    def get(self, request: Request, event_id: int) -> Response:
        """Get a single history event."""
        event = get_object_or_404(
            HistoryEvent.objects.select_related("record", "user"),
            id=event_id,
        )
        
        serializer = HistoryEventDetailSerializer(event)
        return Response(serializer.data)


class SyncHistoryView(APIView):
    """
    Trigger history synchronization from raw history field.
    
    POST /api/history/records/<record_id>/sync/
    
    Parses the record's history_raw field and creates HistoryEvent objects
    for any events not already in the database.
    
    Returns the list of newly created events.
    """
    
    def post(self, request: Request, record_id: int) -> Response:
        """Sync history from raw for a record."""
        # Get the record
        record = get_object_or_404(KarteiRecord, id=record_id)
        
        # Run sync
        created_events = sync_history_from_raw(record)
        
        # Serialize response
        serializer = HistoryEventSerializer(created_events, many=True)
        
        return Response(
            {
                "message": f"Synced {len(created_events)} new history events.",
                "created_count": len(created_events),
                "events": serializer.data,
            },
            status=status.HTTP_201_CREATED if created_events else status.HTTP_200_OK,
        )

