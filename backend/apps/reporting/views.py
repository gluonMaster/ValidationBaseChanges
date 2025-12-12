"""
Views for the reporting app.

This module contains:
- RecentChangesView — list of recent history events across all records

API Endpoints:
- GET /api/reporting/recent-changes/ — list recent changes with filters

See ARCHITECTURE.md Section 2.6 for reporting app overview.
"""

from __future__ import annotations

from datetime import datetime

from django.db.models import Q
from django.utils import timezone
from rest_framework.request import Request
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.history.models import HistoryEvent
from apps.reporting.serializers import (
    RecentChangeItemSerializer,
    RecentChangesResponseSerializer,
)


class RecentChangesView(APIView):
    """
    List recent history events across all records.
    
    GET /api/reporting/recent-changes/
    
    Returns a list of recent change events with record context,
    useful for dashboard/overview displays (аналог grossGeschichte).
    
    Query Parameters:
    - limit: Maximum number of events to return (default: 50, max: 500)
    - offset: Number of events to skip (default: 0)
    - year: Filter by record year (e.g., 2024, 2025)
    - date_from: Filter events from this date (YYYY-MM-DD)
    - date_to: Filter events until this date (YYYY-MM-DD)
    - event_type: Filter by event type (CHANGE, CREATE, APPROVE, DECLINE, IMPORT, RUCK)
    - family_id: Filter by family ID (partial match)
    
    Response:
    {
        "items": [...],
        "total_count": 123,
        "limit": 50,
        "offset": 0,
        "filters": {...}
    }
    """
    
    def get(self, request: Request) -> Response:
        """Get recent changes with optional filters."""
        # Parse pagination parameters
        try:
            limit = int(request.query_params.get("limit", 50))
            offset = int(request.query_params.get("offset", 0))
        except (TypeError, ValueError):
            limit = 50
            offset = 0
        
        limit = min(max(1, limit), 500)  # Clamp between 1 and 500
        offset = max(0, offset)
        
        # Build queryset
        queryset = (
            HistoryEvent.objects
            .select_related("record", "user")
            .order_by("-event_time", "-id")
        )
        
        # Track applied filters for response
        filters_applied: dict = {}
        
        # Filter by year
        year_param = request.query_params.get("year")
        if year_param:
            try:
                year = int(year_param)
                queryset = queryset.filter(record__year=year)
                filters_applied["year"] = year
            except (TypeError, ValueError):
                pass
        
        # Filter by date range
        date_from = request.query_params.get("date_from")
        if date_from:
            try:
                from_date = datetime.strptime(date_from, "%Y-%m-%d")
                from_date = timezone.make_aware(from_date)
                queryset = queryset.filter(event_time__gte=from_date)
                filters_applied["date_from"] = date_from
            except (TypeError, ValueError):
                pass
        
        date_to = request.query_params.get("date_to")
        if date_to:
            try:
                to_date = datetime.strptime(date_to, "%Y-%m-%d")
                # Include the entire day
                to_date = to_date.replace(hour=23, minute=59, second=59)
                to_date = timezone.make_aware(to_date)
                queryset = queryset.filter(event_time__lte=to_date)
                filters_applied["date_to"] = date_to
            except (TypeError, ValueError):
                pass
        
        # Filter by event type
        event_type = request.query_params.get("event_type")
        if event_type:
            queryset = queryset.filter(event_type=event_type.upper())
            filters_applied["event_type"] = event_type.upper()
        
        # Filter by family_id (partial match)
        family_id = request.query_params.get("family_id")
        if family_id:
            queryset = queryset.filter(record__family_id__icontains=family_id)
            filters_applied["family_id"] = family_id
        
        # Get total count before pagination
        total_count = queryset.count()
        
        # Apply pagination
        events = queryset[offset:offset + limit]
        
        # Build response items
        items = []
        for event in events:
            record = event.record
            user_name = None
            if event.user:
                user_name = event.user.get_full_name() or event.user.username
            
            item_data = {
                "event_id": event.id,
                "event_time": event.event_time,
                "event_type": event.event_type,
                "event_type_display": event.get_event_type_display(),
                "changes": event.changes,
                "comment": event.comment,
                "record_id": record.id,
                "year": record.year,
                "family_id": record.family_id,
                "parent_name": record.parent_name,
                "child_name": record.child_name,
                "user_name": user_name,
            }
            items.append(item_data)
        
        # Build response
        response_data = {
            "items": items,
            "total_count": total_count,
            "limit": limit,
            "offset": offset,
            "filters": filters_applied,
        }
        
        return Response(response_data)


class RecentChangesByRecordView(APIView):
    """
    List recent changes grouped by record.
    
    GET /api/reporting/recent-changes-by-record/
    
    Returns recent changes grouped by record, showing the latest event
    for each record that has changes in the specified period.
    
    Query Parameters:
    - limit: Maximum number of records to return (default: 50, max: 200)
    - year: Filter by record year
    - date_from: Filter events from this date (YYYY-MM-DD)
    - date_to: Filter events until this date (YYYY-MM-DD)
    """
    
    def get(self, request: Request) -> Response:
        """Get recent changes grouped by record."""
        # Parse pagination parameters
        try:
            limit = int(request.query_params.get("limit", 50))
        except (TypeError, ValueError):
            limit = 50
        
        limit = min(max(1, limit), 200)
        
        # Build base queryset
        queryset = (
            HistoryEvent.objects
            .select_related("record", "user")
        )
        
        # Track applied filters
        filters_applied: dict = {}
        
        # Filter by year
        year_param = request.query_params.get("year")
        if year_param:
            try:
                year = int(year_param)
                queryset = queryset.filter(record__year=year)
                filters_applied["year"] = year
            except (TypeError, ValueError):
                pass
        
        # Filter by date range
        date_from = request.query_params.get("date_from")
        if date_from:
            try:
                from_date = datetime.strptime(date_from, "%Y-%m-%d")
                from_date = timezone.make_aware(from_date)
                queryset = queryset.filter(event_time__gte=from_date)
                filters_applied["date_from"] = date_from
            except (TypeError, ValueError):
                pass
        
        date_to = request.query_params.get("date_to")
        if date_to:
            try:
                to_date = datetime.strptime(date_to, "%Y-%m-%d")
                to_date = to_date.replace(hour=23, minute=59, second=59)
                to_date = timezone.make_aware(to_date)
                queryset = queryset.filter(event_time__lte=to_date)
                filters_applied["date_to"] = date_to
            except (TypeError, ValueError):
                pass
        
        # Get unique records with their latest event
        # Using distinct on record_id and ordering by event_time desc
        from django.db.models import Max
        
        # Get record IDs with their latest event time
        record_latest = (
            queryset
            .values("record_id")
            .annotate(latest_event_time=Max("event_time"))
            .order_by("-latest_event_time")
            [:limit]
        )
        
        # Build response items
        items = []
        for entry in record_latest:
            record_id = entry["record_id"]
            latest_time = entry["latest_event_time"]
            
            # Get the latest event for this record
            latest_event = (
                queryset
                .filter(record_id=record_id, event_time=latest_time)
                .first()
            )
            
            if latest_event:
                record = latest_event.record
                user_name = None
                if latest_event.user:
                    user_name = latest_event.user.get_full_name() or latest_event.user.username
                
                # Count total events for this record in the period
                event_count = queryset.filter(record_id=record_id).count()
                
                item_data = {
                    "record_id": record.id,
                    "year": record.year,
                    "family_id": record.family_id,
                    "parent_name": record.parent_name,
                    "child_name": record.child_name,
                    "latest_event": {
                        "event_id": latest_event.id,
                        "event_time": latest_event.event_time,
                        "event_type": latest_event.event_type,
                        "event_type_display": latest_event.get_event_type_display(),
                        "comment": latest_event.comment,
                        "user_name": user_name,
                    },
                    "event_count": event_count,
                }
                items.append(item_data)
        
        return Response({
            "items": items,
            "total_records": len(items),
            "limit": limit,
            "filters": filters_applied,
        })

