"""
Serializers for the notifications app.

This module contains DRF serializers for:
- Notification: Serializes notification data for API responses
"""

from rest_framework import serializers

from .models import Notification


class NotificationSerializer(serializers.ModelSerializer):
    """
    Serializer for Notification model.
    
    Includes:
    - All notification fields
    - is_read computed property
    - Related record info (if available)
    """
    
    is_read = serializers.BooleanField(read_only=True)
    type_display = serializers.CharField(
        source='get_type_display',
        read_only=True,
    )
    
    # Nested record info (basic fields)
    record_info = serializers.SerializerMethodField()
    
    class Meta:
        model = Notification
        fields = [
            'id',
            'type',
            'type_display',
            'payload',
            'created_at',
            'read_at',
            'is_read',
            'record_info',
        ]
        read_only_fields = [
            'id',
            'type',
            'type_display',
            'payload',
            'created_at',
            'read_at',
            'is_read',
            'record_info',
        ]
    
    def get_record_info(self, obj: Notification) -> dict | None:
        """
        Get basic info about the related KarteiRecord.
        
        Returns:
            Dictionary with record info, or None if no record.
        """
        if obj.record is None:
            return None
        
        record = obj.record
        return {
            'id': record.id,
            'year': record.year,
            'family_id': record.family_id or '',
            'parent_name': record.parent_name or '',
            'child_name': record.child_name or '',
        }
