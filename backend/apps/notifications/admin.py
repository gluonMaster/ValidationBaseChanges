"""
Django admin configuration for the notifications app.
"""

from django.contrib import admin
from django.utils.html import format_html

from .models import Notification


@admin.register(Notification)
class NotificationAdmin(admin.ModelAdmin):
    """Admin interface for Notification model."""
    
    list_display = [
        'id',
        'recipient',
        'type',
        'record_link',
        'created_at',
        'is_read_display',
    ]
    
    list_filter = [
        'type',
        'created_at',
        ('read_at', admin.EmptyFieldListFilter),
    ]
    
    search_fields = [
        'recipient__username',
        'recipient__email',
        'payload',
    ]
    
    readonly_fields = [
        'id',
        'created_at',
        'read_at',
    ]
    
    ordering = ['-created_at']
    
    date_hierarchy = 'created_at'
    
    def is_read_display(self, obj: Notification) -> str:
        """Display read status as colored icon."""
        if obj.is_read:
            return format_html(
                '<span style="color: green;">✓ Read</span>'
            )
        return format_html(
            '<span style="color: orange;">○ Unread</span>'
        )
    is_read_display.short_description = 'Status'
    
    def record_link(self, obj: Notification) -> str:
        """Display link to related KarteiRecord."""
        if obj.record is None:
            return '-'
        return format_html(
            '<a href="/admin/karteien/karteirecord/{}/change/">'
            'Record #{} ({})</a>',
            obj.record.pk,
            obj.record.id,
            obj.record.year,
        )
    record_link.short_description = 'Record'
