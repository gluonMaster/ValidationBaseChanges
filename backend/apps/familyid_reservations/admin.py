from django.contrib import admin
from django.utils.html import format_html

from .models import FamilyIdReservation


@admin.register(FamilyIdReservation)
class FamilyIdReservationAdmin(admin.ModelAdmin):
    """Admin interface for FamilyIdReservation model (read-only view)."""
    
    list_display = [
        'family_id',
        'reserved_at',
        'reserved_by',
        'status_badge',
        'used_at',
        'note',
    ]
    list_filter = ['is_used', 'reserved_at']
    search_fields = ['family_id', 'note', 'reserved_by__username']
    ordering = ['-reserved_at']
    readonly_fields = [
        'family_id',
        'reserved_at',
        'reserved_by',
        'is_used',
        'used_at',
    ]
    date_hierarchy = 'reserved_at'
    list_per_page = 50

    @admin.display(description='Status')
    def status_badge(self, obj):
        """Display status as a colored badge."""
        if obj.is_used:
            return format_html(
                '<span style="background-color: #6c757d; color: white; '
                'padding: 3px 8px; border-radius: 3px; font-size: 11px;">'
                'Used</span>'
            )
        return format_html(
            '<span style="background-color: #28a745; color: white; '
            'padding: 3px 8px; border-radius: 3px; font-size: 11px;">'
            'Active</span>'
        )

    def has_add_permission(self, request):
        """Disable adding via admin (use API/view instead)."""
        return False

    def has_change_permission(self, request, obj=None):
        """Allow only note editing."""
        return True

    def has_delete_permission(self, request, obj=None):
        """Allow deletion of unused reservations only."""
        if obj and obj.is_used:
            return False
        return True

    def get_readonly_fields(self, request, obj=None):
        """All fields are readonly except note."""
        return [
            'family_id',
            'reserved_at',
            'reserved_by',
            'is_used',
            'used_at',
        ]
