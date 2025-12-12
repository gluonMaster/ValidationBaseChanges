"""
Serializers for the history app.

This module contains DRF serializers for:
- HistoryEvent — serialization of history events
- HistoryEventList — list serialization with record context
"""

from __future__ import annotations

from rest_framework import serializers

from apps.history.models import EventType, HistoryEvent


class HistoryEventSerializer(serializers.ModelSerializer):
    """
    Serializer for HistoryEvent model.
    
    Includes computed fields for display convenience.
    """
    
    event_type_display = serializers.CharField(
        source="get_event_type_display",
        read_only=True,
    )
    
    changed_fields = serializers.ListField(
        child=serializers.CharField(),
        read_only=True,
    )
    
    user_name = serializers.SerializerMethodField()
    
    class Meta:
        model = HistoryEvent
        fields = [
            "id",
            "record_id",
            "event_time",
            "event_type",
            "event_type_display",
            "changes",
            "comment",
            "changed_fields",
            "user_name",
            "raw_history_fragment",
            "created_at",
        ]
        read_only_fields = fields
    
    def get_user_name(self, obj: HistoryEvent) -> str | None:
        """Get the username of the user who made the change."""
        if obj.user:
            return obj.user.get_full_name() or obj.user.username
        return None


class HistoryEventDetailSerializer(HistoryEventSerializer):
    """
    Detailed serializer for HistoryEvent with full record info.
    """
    
    record_info = serializers.SerializerMethodField()
    
    class Meta(HistoryEventSerializer.Meta):
        fields = HistoryEventSerializer.Meta.fields + ["record_info"]
    
    def get_record_info(self, obj: HistoryEvent) -> dict:
        """Get basic info about the associated record."""
        record = obj.record
        return {
            "id": record.id,
            "year": record.year,
            "family_id": record.family_id,
            "parent_name": record.parent_name,
            "child_name": record.child_name,
        }


class RecordHistorySerializer(serializers.Serializer):
    """
    Serializer for the record history list response.
    
    Combines record basic info with its history events.
    """
    
    id = serializers.IntegerField()
    year = serializers.IntegerField()
    family_id = serializers.CharField()
    parent_name = serializers.CharField()
    child_name = serializers.CharField()
    history_events = HistoryEventSerializer(many=True)
    total_events = serializers.IntegerField()


class HistoryEventCreateSerializer(serializers.ModelSerializer):
    """
    Serializer for creating new HistoryEvent.
    
    Used when recording new changes programmatically.
    """
    
    class Meta:
        model = HistoryEvent
        fields = [
            "record",
            "event_time",
            "event_type",
            "changes",
            "comment",
            "raw_history_fragment",
        ]
    
    def validate_event_type(self, value: str) -> str:
        """Validate that event_type is a valid choice."""
        valid_types = [choice[0] for choice in EventType.choices]
        if value not in valid_types:
            raise serializers.ValidationError(
                f"Invalid event_type. Must be one of: {valid_types}"
            )
        return value
    
    def validate_changes(self, value: dict) -> dict:
        """Validate changes structure."""
        if not isinstance(value, dict):
            raise serializers.ValidationError("Changes must be a dictionary.")
        
        for field_name, change_data in value.items():
            if not isinstance(change_data, dict):
                raise serializers.ValidationError(
                    f"Change data for '{field_name}' must be a dictionary."
                )
            if "old" not in change_data or "new" not in change_data:
                raise serializers.ValidationError(
                    f"Change data for '{field_name}' must have 'old' and 'new' keys."
                )
        
        return value
