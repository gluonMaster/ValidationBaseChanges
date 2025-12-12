"""
Models for the approvals app.

This module contains:
- PendingChange — записи, ожидающие одобрения Superadmin (аналог pre_tblKartei)
- DeclinedChange — отклонённые записи (аналог decl_tblKartei)

Mapping:
- Access table: pre_tblKartei -> PendingChange
- Access table: decl_tblKartei -> DeclinedChange

The models store a snapshot of changed data as JSON, linking back to the
original KarteiRecord. This provides flexibility for storing arbitrary
field values while maintaining referential integrity.

See DOMAIN_MODEL.md sections 2.1-2.2 for Access table structure.
"""

from __future__ import annotations

from django.db import models


class PendingChange(models.Model):
    """
    Pending change awaiting Superadmin approval.

    Maps to Access table: pre_tblKartei

    When an Admin makes a risky change (changing tracked fields on an existing
    record), instead of updating KarteiRecord directly, a PendingChange is
    created/updated. This holds the proposed new values until Superadmin
    approves or declines.

    Key behavior:
    - One pending change per KarteiRecord at a time (FK unique constraint).
    - Subsequent risky edits to the same record update the existing PendingChange.
    - Upon approval: changes apply to KarteiRecord, PendingChange is marked processed.
    - Upon decline: PendingChange is deleted, DeclinedChange is created.

    The snapshot field stores all tracked field values as JSON:
    {
        "family_id": "...",
        "parent_name": "...",
        "child_name": "...",
        "birthdate": "YYYY-MM-DD" or null,
        "address": "...",
        "phone": "...",
        "mobile": "...",
        "email": "...",
        "subject1": "...",
        "price1": "123.45" or null,
        "subject2": "...",
        "price2": "123.45" or null,
        "month_1": "..." ... "month_12": "...",
        "extra1": "...",
        "extra2": "...",
        "extra3": "..."
    }
    """

    # -------------------------------------------------------------------------
    # Primary Key
    # -------------------------------------------------------------------------

    id = models.BigAutoField(
        primary_key=True,
        help_text="Auto-generated primary key (not the same as KarteiRecord.id).",
    )

    # -------------------------------------------------------------------------
    # Link to KarteiRecord
    # -------------------------------------------------------------------------

    record = models.OneToOneField(
        "karteien.KarteiRecord",
        on_delete=models.CASCADE,
        related_name="pending_change",
        help_text="Reference to the original KarteiRecord. OneToOne ensures only "
                  "one pending change per record at a time.",
    )

    # -------------------------------------------------------------------------
    # Snapshot of proposed changes
    # -------------------------------------------------------------------------

    snapshot = models.JSONField(
        default=dict,
        help_text="JSON snapshot of proposed field values. Contains all tracked "
                  "fields from KarteiRecord. Format: {'field_name': value, ...}. "
                  "See DOMAIN_MODEL.md Section 4 for tracked fields list.",
    )

    # -------------------------------------------------------------------------
    # Processing status
    # -------------------------------------------------------------------------

    is_processed = models.BooleanField(
        default=False,
        db_index=True,
        help_text="True if this pending change has been processed (approved/declined). "
                  "Processed records are kept for audit trail.",
    )

    # -------------------------------------------------------------------------
    # Timestamps
    # -------------------------------------------------------------------------

    created_at = models.DateTimeField(
        auto_now_add=True,
        help_text="Timestamp when pending change was first created.",
    )

    updated_at = models.DateTimeField(
        auto_now=True,
        help_text="Timestamp of last update (e.g., Admin made further edits).",
    )

    class Meta:
        db_table = "approvals_pending_change"
        ordering = ["-created_at"]
        verbose_name = "Pending Change"
        verbose_name_plural = "Pending Changes"
        indexes = [
            models.Index(
                fields=["is_processed", "created_at"],
                name="idx_pending_processed_created",
            ),
        ]

    def __str__(self) -> str:
        status = "processed" if self.is_processed else "pending"
        return f"PendingChange(record={self.record_id}, {status})"


class DeclinedChange(models.Model):
    """
    Declined change rejected by Superadmin.

    Maps to Access table: decl_tblKartei

    When Superadmin declines a pending change, the data moves here along with
    the decline reason. Admin can then view declined records, fix issues,
    and resubmit (creating a new PendingChange).

    Key behavior:
    - Created when Superadmin declines a PendingChange.
    - Contains the same snapshot as the rejected PendingChange.
    - Admin sees DECLINED status on the record and can fix/resubmit.
    - Multiple declined changes can exist for the same record (history of rejections).
    """

    # -------------------------------------------------------------------------
    # Primary Key
    # -------------------------------------------------------------------------

    id = models.BigAutoField(
        primary_key=True,
        help_text="Auto-generated primary key.",
    )

    # -------------------------------------------------------------------------
    # Link to KarteiRecord
    # -------------------------------------------------------------------------

    record = models.ForeignKey(
        "karteien.KarteiRecord",
        on_delete=models.CASCADE,
        related_name="declined_changes",
        help_text="Reference to the KarteiRecord that had changes declined. "
                  "ForeignKey (not OneToOne) allows multiple decline events per record.",
    )

    # -------------------------------------------------------------------------
    # Snapshot of declined changes
    # -------------------------------------------------------------------------

    snapshot = models.JSONField(
        default=dict,
        help_text="JSON snapshot of declined field values. Same format as "
                  "PendingChange.snapshot. Preserved for audit trail.",
    )

    # -------------------------------------------------------------------------
    # Decline information
    # -------------------------------------------------------------------------

    decline_reason = models.TextField(
        blank=True,
        default="",
        help_text="Reason provided by Superadmin for declining the change. "
                  "May include instructions for Admin on how to fix the data.",
    )

    declined_by = models.ForeignKey(
        "accounts.User",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="declined_changes",
        help_text="Superadmin user who declined the change.",
    )

    # -------------------------------------------------------------------------
    # Timestamps
    # -------------------------------------------------------------------------

    created_at = models.DateTimeField(
        auto_now_add=True,
        help_text="Timestamp when change was declined.",
    )

    updated_at = models.DateTimeField(
        auto_now=True,
        help_text="Timestamp of last update.",
    )

    class Meta:
        db_table = "approvals_declined_change"
        ordering = ["-created_at"]
        verbose_name = "Declined Change"
        verbose_name_plural = "Declined Changes"
        indexes = [
            models.Index(
                fields=["record", "created_at"],
                name="idx_declined_record_created",
            ),
        ]

    def __str__(self) -> str:
        reason_preview = self.decline_reason[:30] + "..." if len(self.decline_reason) > 30 else self.decline_reason
        return f"DeclinedChange(record={self.record_id}, reason='{reason_preview}')"


class SuperadminState(models.Model):
    """
    Tracks Superadmin state for features like NeuList.

    Stores per-user settings such as:
    - last_seen_id: Last seen record ID for NeuList feature
    - last_seen_date: Alternative to ID-based tracking

    This replaces the hidden NeuConfig sheet in the VBA implementation.

    See valid_NeuList.bas for VBA reference:
    - LastSeenID stored in hidden NeuConfig sheet
    - Neu sheet displays all records with ID > LastSeenID
    """

    # -------------------------------------------------------------------------
    # Primary Key
    # -------------------------------------------------------------------------

    id = models.BigAutoField(
        primary_key=True,
        help_text="Auto-generated primary key.",
    )

    # -------------------------------------------------------------------------
    # Link to User
    # -------------------------------------------------------------------------

    user = models.OneToOneField(
        "accounts.User",
        on_delete=models.CASCADE,
        related_name="superadmin_state",
        help_text="Superadmin user this state belongs to.",
    )

    # -------------------------------------------------------------------------
    # NeuList State
    # -------------------------------------------------------------------------

    last_seen_id = models.PositiveIntegerField(
        default=0,
        help_text="Last seen record ID. Records with ID > this value "
                  "are considered 'new' in NeuList view.",
    )

    last_seen_date = models.DateTimeField(
        null=True,
        blank=True,
        help_text="Alternative: last seen date. Records created after this "
                  "are considered 'new'. Optional, ID-based is primary.",
    )

    # -------------------------------------------------------------------------
    # Timestamps
    # -------------------------------------------------------------------------

    updated_at = models.DateTimeField(
        auto_now=True,
        help_text="Timestamp of last state update.",
    )

    class Meta:
        db_table = "approvals_superadmin_state"
        verbose_name = "Superadmin State"
        verbose_name_plural = "Superadmin States"

    def __str__(self) -> str:
        return f"SuperadminState(user={self.user_id}, last_seen_id={self.last_seen_id})"
