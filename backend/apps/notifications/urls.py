"""
URL configuration for the notifications app.

Endpoints:
- GET /api/notifications/ - List notifications for current user
- GET /api/notifications/unread-count/ - Get count of unread notifications
- POST /api/notifications/<id>/read/ - Mark notification as read
- POST /api/notifications/read-all/ - Mark all notifications as read
"""

from django.urls import path

from .views import (
    NotificationListView,
    NotificationReadAllView,
    NotificationReadView,
    NotificationUnreadCountView,
)

app_name = 'notifications'

urlpatterns = [
    # List notifications
    path('', NotificationListView.as_view(), name='notification-list'),
    
    # Unread count
    path('unread-count/', NotificationUnreadCountView.as_view(), name='unread-count'),
    
    # Mark all as read
    path('read-all/', NotificationReadAllView.as_view(), name='read-all'),
    
    # Mark single notification as read
    path('<int:notification_id>/read/', NotificationReadView.as_view(), name='notification-read'),
]
