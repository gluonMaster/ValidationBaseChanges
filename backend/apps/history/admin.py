"""
Admin configuration for the history app.
"""

from django.contrib import admin

from apps.history.models import HistoryEvent


@admin.register(HistoryEvent)
class HistoryEventAdmin(admin.ModelAdmin):
    """Admin interface for HistoryEvent model."""
    
    list_display = [
        "id",
        "record_id",
        "event_time",
        "event_type",
        "get_changed_fields_display",
        "comment_preview",
        "user",
        "created_at",
    ]
    
    list_filter = [
        "event_type",
        "event_time",
        "record__year",
        "created_at",
    ]
    
    search_fields = [
        "record__family_id",
        "record__parent_name",
        "record__child_name",
        "comment",
        "raw_history_fragment",
    ]
    
    readonly_fields = [
        "id",
        "created_at",
        "updated_at",
        "raw_history_fragment",
    ]
    
    raw_id_fields = ["record", "user"]
    
    ordering = ["-event_time", "-id"]
    
    date_hierarchy = "event_time"
    
    fieldsets = [
        (None, {
            "fields": ["id", "record", "user"],
        }),
        ("Event Details", {
            "fields": ["event_time", "event_type", "changes", "comment"],
        }),
        ("Legacy Data", {
            "fields": ["raw_history_fragment"],
            "classes": ["collapse"],
        }),
        ("Timestamps", {
            "fields": ["created_at", "updated_at"],
            "classes": ["collapse"],
        }),
    ]
    
    def get_changed_fields_display(self, obj: HistoryEvent) -> str:
        """Display list of changed fields."""
        fields = obj.changed_fields
        if len(fields) > 3:
            return f"{', '.join(fields[:3])}, +{len(fields) - 3} more"
        return ", ".join(fields) if fields else "-"
    
    get_changed_fields_display.short_description = "Changed Fields"
    
    def comment_preview(self, obj: HistoryEvent) -> str:
        """Display truncated comment."""
        if obj.comment:
            return obj.comment[:50] + "..." if len(obj.comment) > 50 else obj.comment
        return "-"
    
    comment_preview.short_description = "Comment"

