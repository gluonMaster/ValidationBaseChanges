"""
Views for the notifications app.

This module contains:
- NotificationListView: List notifications for the current user
- NotificationReadView: Mark a notification as read
- NotificationUnreadCountView: Get count of unread notifications

API endpoints:
- GET /api/notifications/ - List notifications (with optional filtering)
- POST /api/notifications/<id>/read/ - Mark notification as read
- POST /api/notifications/read-all/ - Mark all notifications as read
- GET /api/notifications/unread-count/ - Get unread notification count

See ARCHITECTURE.md section 2.5 for architecture details.
"""

from rest_framework import status
from rest_framework.generics import ListAPIView
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from rest_framework.views import APIView

from .models import Notification
from .serializers import NotificationSerializer
from .services import (
    get_notifications,
    get_unread_count,
    mark_all_notifications_read,
    mark_notification_read,
)


class NotificationListView(ListAPIView):
    """
    List notifications for the current user.
    
    GET /api/notifications/
    
    Query parameters:
    - unread: If 'true', only return unread notifications
    - type: Filter by notification type (PENDING_CREATED, DECLINED_CREATED)
    - limit: Maximum number of notifications to return
    """
    serializer_class = NotificationSerializer
    permission_classes = [IsAuthenticated]
    
    def get_queryset(self):
        """Get notifications for the current user with optional filtering."""
        unread_only = self.request.query_params.get('unread', '').lower() == 'true'
        notification_type = self.request.query_params.get('type', None)
        limit_str = self.request.query_params.get('limit', None)
        
        limit = None
        if limit_str:
            try:
                limit = int(limit_str)
            except ValueError:
                pass
        
        return get_notifications(
            user=self.request.user,
            unread_only=unread_only,
            notification_type=notification_type,
            limit=limit,
        )


class NotificationUnreadCountView(APIView):
    """
    Get the count of unread notifications for the current user.
    
    GET /api/notifications/unread-count/
    
    Returns:
        {"count": <number>}
    """
    permission_classes = [IsAuthenticated]
    
    def get(self, request):
        """Return the count of unread notifications."""
        count = get_unread_count(request.user)
        return Response({"count": count})


class NotificationReadView(APIView):
    """
    Mark a notification as read.
    
    POST /api/notifications/<id>/read/
    
    Returns the updated notification or 404 if not found.
    """
    permission_classes = [IsAuthenticated]
    
    def post(self, request, notification_id):
        """Mark the specified notification as read."""
        notification = mark_notification_read(notification_id, request.user)
        
        if notification is None:
            return Response(
                {"error": "Notification not found"},
                status=status.HTTP_404_NOT_FOUND,
            )
        
        serializer = NotificationSerializer(notification)
        return Response(serializer.data)


class NotificationReadAllView(APIView):
    """
    Mark all notifications as read for the current user.
    
    POST /api/notifications/read-all/
    
    Returns:
        {"marked_count": <number>}
    """
    permission_classes = [IsAuthenticated]
    
    def post(self, request):
        """Mark all notifications as read."""
        count = mark_all_notifications_read(request.user)
        return Response({"marked_count": count})
