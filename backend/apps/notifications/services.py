"""
Services for the notifications app.

This module contains:
- notify_pending_created: Create notifications for Superadmins about new pending changes
- notify_declined_created: Create notifications for Admins about declined changes
- get_unread_count: Get count of unread notifications for a user
- get_notifications: Get list of notifications for a user

These services implement internal notification logic. External notifications
(e-mail, Telegram, etc.) will be added in future steps.

See ARCHITECTURE.md section 2.5 and DOMAIN_MODEL.md section 3 for context.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from django.db.models import QuerySet
from django.utils import timezone

from .models import Notification, NotificationType


if TYPE_CHECKING:
    from apps.accounts.models import User
    from apps.approvals.models import DeclinedChange, PendingChange
    from apps.karteien.models import KarteiRecord


# =============================================================================
# Constants
# =============================================================================

# Role names used for filtering users
# These should match the role values in accounts.User model
ROLE_SUPERADMIN = "SUPERADMIN"
ROLE_ADMIN = "ADMIN"


# =============================================================================
# Helper Functions
# =============================================================================

def _get_users_by_role(role: str) -> QuerySet:
    """
    Get all users with the specified role.
    
    Args:
        role: Role name (ADMIN, SUPERADMIN).
    
    Returns:
        QuerySet of users with that role.
    """
    # Import here to avoid circular imports
    from django.contrib.auth import get_user_model
    
    User = get_user_model()
    
    # Filter by role field if it exists
    if hasattr(User, 'role'):
        return User.objects.filter(role=role, is_active=True)
    
    # Fallback: use groups if role field doesn't exist
    return User.objects.filter(groups__name=role, is_active=True)


def _build_record_payload(record: KarteiRecord) -> dict:
    """
    Build a notification payload from a KarteiRecord.
    
    Args:
        record: The KarteiRecord to extract info from.
    
    Returns:
        Dictionary with record info for notification payload.
    """
    return {
        "record_id": record.id,
        "year": record.year,
        "family_id": record.family_id or "",
        "parent_name": record.parent_name or "",
        "child_name": record.child_name or "",
    }


# =============================================================================
# Notification Creation Functions
# =============================================================================

def notify_pending_created(
    record: KarteiRecord,
    pending: PendingChange,
) -> list[Notification]:
    """
    Create notifications for all Superadmins about a new pending change.
    
    This is called when a risky change is made and a PendingChange is created.
    All Superadmin users receive a PENDING_CREATED notification.
    
    Idempotency: Does not create duplicate notifications if an unread
    notification already exists for this record and recipient.
    
    Args:
        record: The KarteiRecord that has the pending change.
        pending: The PendingChange that was created.
    
    Returns:
        List of created Notification instances.
    
    Example:
        >>> pending = create_or_update_pending_change(record)
        >>> notifications = notify_pending_created(record, pending)
        >>> print(f"Notified {len(notifications)} superadmins")
    """
    created_notifications: list[Notification] = []
    
    # Get all Superadmin users
    superadmins = _get_users_by_role(ROLE_SUPERADMIN)
    
    # Build payload
    payload = _build_record_payload(record)
    payload["pending_id"] = pending.id
    
    for user in superadmins:
        # Check for existing unread notification (idempotency)
        existing = Notification.objects.filter(
            recipient=user,
            record=record,
            type=NotificationType.PENDING_CREATED,
            read_at__isnull=True,
        ).exists()
        
        if not existing:
            notification = Notification.objects.create(
                recipient=user,
                type=NotificationType.PENDING_CREATED,
                record=record,
                payload=payload,
            )
            created_notifications.append(notification)
    
    return created_notifications


def notify_declined_created(
    record: KarteiRecord,
    declined: DeclinedChange,
) -> list[Notification]:
    """
    Create notifications for all Admins about a declined change.
    
    This is called when a Superadmin declines a pending change.
    All Admin users receive a DECLINED_CREATED notification.
    
    In future versions, this could be refined to notify only the Admin
    who originally made the change.
    
    Idempotency: Does not create duplicate notifications if an unread
    notification already exists for this record and recipient.
    
    Args:
        record: The KarteiRecord that was declined.
        declined: The DeclinedChange that was created.
    
    Returns:
        List of created Notification instances.
    
    Example:
        >>> declined = create_declined_change(record, reason, superadmin)
        >>> notifications = notify_declined_created(record, declined)
        >>> print(f"Notified {len(notifications)} admins")
    """
    created_notifications: list[Notification] = []
    
    # Get all Admin users
    admins = _get_users_by_role(ROLE_ADMIN)
    
    # Build payload
    payload = _build_record_payload(record)
    payload["declined_id"] = declined.id
    payload["decline_reason"] = declined.decline_reason or ""
    if declined.declined_by:
        payload["declined_by"] = str(declined.declined_by)
    
    for user in admins:
        # Check for existing unread notification (idempotency)
        existing = Notification.objects.filter(
            recipient=user,
            record=record,
            type=NotificationType.DECLINED_CREATED,
            read_at__isnull=True,
        ).exists()
        
        if not existing:
            notification = Notification.objects.create(
                recipient=user,
                type=NotificationType.DECLINED_CREATED,
                record=record,
                payload=payload,
            )
            created_notifications.append(notification)
    
    return created_notifications


def notify_approved(
    record: KarteiRecord,
    pending: PendingChange,
    approved_by: User | None = None,
) -> list[Notification]:
    """
    Create notifications for all Admins about an approved change.

    This is called when a Superadmin approves a pending change.
    All Admin users receive an APPROVED notification.

    Args:
        record: The KarteiRecord that was approved.
        pending: The PendingChange that was approved.
        approved_by: The Superadmin who approved.

    Returns:
        List of created Notification instances.
    """
    created_notifications: list[Notification] = []

    # Get all Admin users
    admins = _get_users_by_role(ROLE_ADMIN)

    # Build payload
    payload = _build_record_payload(record)
    payload["pending_id"] = pending.id
    if approved_by:
        payload["approved_by"] = str(approved_by)

    for user in admins:
        # Check for existing unread notification (idempotency)
        existing = Notification.objects.filter(
            recipient=user,
            record=record,
            type=NotificationType.APPROVED,
            read_at__isnull=True,
        ).exists()

        if not existing:
            notification = Notification.objects.create(
                recipient=user,
                type=NotificationType.APPROVED,
                record=record,
                payload=payload,
            )
            created_notifications.append(notification)

    return created_notifications


# =============================================================================
# Query Functions
# =============================================================================

def get_unread_count(user: User) -> int:
    """
    Get the count of unread notifications for a user.
    
    Args:
        user: The user to count notifications for.
    
    Returns:
        Number of unread notifications.
    """
    return Notification.objects.filter(
        recipient=user,
        read_at__isnull=True,
    ).count()


def get_notifications(
    user: User,
    unread_only: bool = False,
    notification_type: str | None = None,
    limit: int | None = None,
) -> QuerySet[Notification]:
    """
    Get notifications for a user with optional filtering.
    
    Args:
        user: The user to get notifications for.
        unread_only: If True, only return unread notifications.
        notification_type: Filter by notification type (PENDING_CREATED, DECLINED_CREATED).
        limit: Maximum number of notifications to return.
    
    Returns:
        QuerySet of Notification instances, ordered by created_at descending.
    """
    queryset = Notification.objects.filter(recipient=user)
    
    if unread_only:
        queryset = queryset.filter(read_at__isnull=True)
    
    if notification_type:
        queryset = queryset.filter(type=notification_type)
    
    queryset = queryset.select_related("record").order_by("-created_at")
    
    if limit:
        queryset = queryset[:limit]
    
    return queryset


def mark_notification_read(notification_id: int, user: User) -> Notification | None:
    """
    Mark a notification as read.
    
    Args:
        notification_id: ID of the notification to mark as read.
        user: The user who owns the notification (for security).
    
    Returns:
        The updated Notification, or None if not found or not owned by user.
    """
    try:
        notification = Notification.objects.get(id=notification_id, recipient=user)
        notification.mark_as_read()
        return notification
    except Notification.DoesNotExist:
        return None


def mark_all_notifications_read(user: User) -> int:
    """
    Mark all unread notifications as read for a user.
    
    Args:
        user: The user whose notifications to mark as read.
    
    Returns:
        Number of notifications marked as read.
    """
    count = Notification.objects.filter(
        recipient=user,
        read_at__isnull=True,
    ).update(read_at=timezone.now())
    
    return count
