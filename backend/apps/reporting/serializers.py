"""
Serializers for the reporting app.

This module contains DRF serializers for:
- RecentChangeItem — serialization of recent change entries
- RecentChangesResponse — list response with metadata
"""

from __future__ import annotations

from rest_framework import serializers


class RecentChangeItemSerializer(serializers.Serializer):
    """
    Serializer for a single recent change item.
    
    Combines history event info with record context.
    """
    
    # Event info
    event_id = serializers.IntegerField()
    event_time = serializers.DateTimeField()
    event_type = serializers.CharField()
    event_type_display = serializers.CharField()
    changes = serializers.DictField()
    comment = serializers.CharField(allow_blank=True)
    
    # Record info
    record_id = serializers.IntegerField()
    year = serializers.IntegerField()
    family_id = serializers.CharField()
    parent_name = serializers.CharField()
    child_name = serializers.CharField()
    
    # User info
    user_name = serializers.CharField(allow_null=True)


class RecentChangesResponseSerializer(serializers.Serializer):
    """
    Serializer for the recent changes list response.
    """
    
    items = RecentChangeItemSerializer(many=True)
    total_count = serializers.IntegerField()
    limit = serializers.IntegerField()
    offset = serializers.IntegerField()
    filters = serializers.DictField()
