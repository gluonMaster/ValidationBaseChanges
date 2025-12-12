"""
Models for the notifications app.

This module contains:
- Notification — уведомления для пользователей (recipient, type, payload, read_at)

Notifications are used for internal communication:
- Superadmin: receives notifications about new pending changes
- Admin: receives notifications about declined changes

See ARCHITECTURE.md section 2.5 for architecture details.
"""

from __future__ import annotations

from django.db import models
from django.conf import settings


class NotificationType(models.TextChoices):
    """
    Types of notifications in the system.
    
    PENDING_CREATED: A new pending change was created (for Superadmin).
    DECLINED_CREATED: A pending change was declined (for Admin).
    APPROVED: A pending change was approved (for Admin).
    """
    PENDING_CREATED = "PENDING_CREATED", "New pending change"
    DECLINED_CREATED = "DECLINED_CREATED", "Change declined"
    APPROVED = "APPROVED", "Change approved"


class Notification(models.Model):
    """
    Internal notification for users.
    
    Notifications are created automatically when:
    - A risky change creates a PendingChange (notifies Superadmins)
    - A Superadmin declines a change (notifies Admins)
    
    Key behavior:
    - Notifications are idempotent: duplicate notifications for the same
      record and type are not created if one already exists unread.
    - read_at is set when user marks notification as read.
    - Notifications are not year-specific; they reference KarteiRecord
      which already contains the year field.
    """
    
    # -------------------------------------------------------------------------
    # Primary Key
    # -------------------------------------------------------------------------
    
    id = models.BigAutoField(
        primary_key=True,
        help_text="Auto-generated primary key.",
    )
    
    # -------------------------------------------------------------------------
    # Recipient
    # -------------------------------------------------------------------------
    
    recipient = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="notifications",
        help_text="User who receives this notification.",
    )
    
    # -------------------------------------------------------------------------
    # Notification Type
    # -------------------------------------------------------------------------
    
    type = models.CharField(
        max_length=32,
        choices=NotificationType.choices,
        db_index=True,
        help_text="Type of notification (PENDING_CREATED, DECLINED_CREATED).",
    )
    
    # -------------------------------------------------------------------------
    # Link to KarteiRecord (optional)
    # -------------------------------------------------------------------------
    
    record = models.ForeignKey(
        "karteien.KarteiRecord",
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name="notifications",
        help_text="Reference to the KarteiRecord this notification is about. "
                  "Null if notification is not about a specific record.",
    )
    
    # -------------------------------------------------------------------------
    # Payload (additional data)
    # -------------------------------------------------------------------------
    
    payload = models.JSONField(
        default=dict,
        blank=True,
        help_text="Additional data for the notification. Can include: "
                  "family_id, parent_name, child_name, decline_reason, etc.",
    )
    
    # -------------------------------------------------------------------------
    # Timestamps
    # -------------------------------------------------------------------------
    
    created_at = models.DateTimeField(
        auto_now_add=True,
        db_index=True,
        help_text="When the notification was created.",
    )
    
    read_at = models.DateTimeField(
        null=True,
        blank=True,
        db_index=True,
        help_text="When the notification was marked as read. Null if unread.",
    )
    
    # -------------------------------------------------------------------------
    # Meta
    # -------------------------------------------------------------------------
    
    class Meta:
        db_table = "notifications_notification"
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["recipient", "read_at"]),
            models.Index(fields=["recipient", "type", "read_at"]),
        ]
        verbose_name = "Notification"
        verbose_name_plural = "Notifications"
    
    # -------------------------------------------------------------------------
    # Properties
    # -------------------------------------------------------------------------
    
    @property
    def is_read(self) -> bool:
        """Return True if notification has been read."""
        return self.read_at is not None
    
    # -------------------------------------------------------------------------
    # String representation
    # -------------------------------------------------------------------------
    
    def __str__(self) -> str:
        status = "read" if self.is_read else "unread"
        return f"Notification({self.type}, {self.recipient}, {status})"
    
    # -------------------------------------------------------------------------
    # Methods
    # -------------------------------------------------------------------------
    
    def mark_as_read(self) -> None:
        """
        Mark this notification as read.
        
        Sets read_at to current datetime and saves.
        """
        from django.utils import timezone
        
        if self.read_at is None:
            self.read_at = timezone.now()
            self.save(update_fields=["read_at"])
