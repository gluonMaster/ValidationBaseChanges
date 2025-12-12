"""
Models for the history app.

This module contains:
- HistoryEvent — нормализованные события изменений (record, event_time, user, event_type, changes)
- EventType — типы событий истории
- Поддержка raw_history_fragment для совместимости с текстовым форматом из VBA (AZ/Value52)

Mapping:
- VBA module: Export_HistoryBuilder, Export_HistoryParser
- Excel column: AZ (52)
- Access field: Value52

See DOMAIN_MODEL.md Section 5 for history format details.
See ARCHITECTURE.md Section 2.4 for history app overview.
"""

from __future__ import annotations

from django.conf import settings
from django.db import models


# =============================================================================
# Choices / Enums
# =============================================================================

class EventType(models.TextChoices):
    """
    Type of history event.
    
    Based on VBA history format and operations:
    - CHANGE: regular field change (tracked fields)
    - CREATE: new record created
    - APPROVE: pending change approved by Superadmin
    - DECLINE: pending change declined by Superadmin (DCL tag in history)
    - IMPORT: record imported from legacy data
    - RUCK: retroactive change (RUCK: prefix in history)
    """
    CHANGE = "CHANGE", "Field Change"
    CREATE = "CREATE", "Record Created"
    APPROVE = "APPROVE", "Approved"
    DECLINE = "DECLINE", "Declined"
    IMPORT = "IMPORT", "Imported"
    RUCK = "RUCK", "Retroactive Change"


# =============================================================================
# HistoryEvent Model
# =============================================================================

class HistoryEvent(models.Model):
    """
    Normalized history event for a KarteiRecord.
    
    Stores individual change events parsed from the raw history string (AZ/Value52).
    Each event represents one "session" from the legacy format:
        [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||
    
    The raw_history_fragment field stores the original text segment for debugging
    and backward compatibility.
    
    Changes are stored as JSON with structure:
        {
            "field_name": {"old": <value>, "new": <value>},
            ...
        }
    
    Field names use Django model field names (e.g., "family_id", "month_1").
    """
    
    # -------------------------------------------------------------------------
    # Primary Key
    # -------------------------------------------------------------------------
    
    id = models.BigAutoField(primary_key=True)
    
    # -------------------------------------------------------------------------
    # Relations
    # -------------------------------------------------------------------------
    
    record = models.ForeignKey(
        "karteien.KarteiRecord",
        on_delete=models.CASCADE,
        related_name="history_events",
        help_text="The KarteiRecord this event belongs to.",
    )
    
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="history_events",
        help_text="User who made the change. Null for imported/legacy events.",
    )
    
    # -------------------------------------------------------------------------
    # Event Data
    # -------------------------------------------------------------------------
    
    event_time = models.DateTimeField(
        db_index=True,
        help_text="When the event occurred. Parsed from history date or set at creation.",
    )
    
    event_type = models.CharField(
        max_length=20,
        choices=EventType.choices,
        default=EventType.CHANGE,
        db_index=True,
        help_text="Type of event: CHANGE, CREATE, APPROVE, DECLINE, IMPORT, RUCK.",
    )
    
    changes = models.JSONField(
        default=dict,
        blank=True,
        help_text=(
            "JSON structure of changes: "
            '{"field_name": {"old": <value>, "new": <value>}, ...}'
        ),
    )
    
    comment = models.TextField(
        blank=True,
        default="",
        help_text="User comment for this change session (parsed from /@...@/).",
    )
    
    # -------------------------------------------------------------------------
    # Legacy Compatibility
    # -------------------------------------------------------------------------
    
    raw_history_fragment = models.TextField(
        blank=True,
        default="",
        help_text=(
            "Original history string segment from AZ/Value52. "
            "Stored for debugging and backward compatibility."
        ),
    )
    
    # -------------------------------------------------------------------------
    # Timestamps
    # -------------------------------------------------------------------------
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    # -------------------------------------------------------------------------
    # Meta
    # -------------------------------------------------------------------------
    
    class Meta:
        db_table = "history_event"
        ordering = ["-event_time", "-id"]
        indexes = [
            models.Index(fields=["record", "event_time"]),
            models.Index(fields=["event_type", "event_time"]),
        ]
        verbose_name = "History Event"
        verbose_name_plural = "History Events"
    
    # -------------------------------------------------------------------------
    # String Representation
    # -------------------------------------------------------------------------
    
    def __str__(self) -> str:
        record_id = self.record_id if self.record_id else "?"
        return f"HistoryEvent(record={record_id}, type={self.event_type}, time={self.event_time})"
    
    # -------------------------------------------------------------------------
    # Properties
    # -------------------------------------------------------------------------
    
    @property
    def changed_fields(self) -> list[str]:
        """Return list of field names that were changed in this event."""
        if isinstance(self.changes, dict):
            return list(self.changes.keys())
        return []
    
    @property
    def is_retroactive(self) -> bool:
        """Check if this is a retroactive (RUCK) change."""
        return self.event_type == EventType.RUCK
    
    @property
    def is_decline(self) -> bool:
        """Check if this is a decline event."""
        return self.event_type == EventType.DECLINE
