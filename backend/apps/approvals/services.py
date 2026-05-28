"""
Services for the approvals app.

This module contains:
- Risk classification logic (is_risky_change, classify_change)
- Pending/declined change management (create_or_update_pending_change, create_declined_change)

These services implement the business logic previously handled by:
- Export_RiskClassification.IsRiskyChange
- Export_RiskClassification.HasTrackedFieldChanges
- Export_OverlayPending / Export_DeclinedTools

Key concepts:
- TRACKED_FIELDS: Fields that, when changed, trigger risk classification.
- Risk classification compares old vs new values of tracked fields.
- New records (no original) are always considered SAFE.
- Changes to only non-tracked fields are considered SAFE.

See DOMAIN_MODEL.md Section 4 for the tracked fields list.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import TYPE_CHECKING, Any, Literal

from apps.karteien.models import TRACKED_FIELDS, KarteiRecord, RecordStatus
from apps.karteien.services.pending_snapshot import (
    build_tracked_snapshot,
    apply_pending_snapshot_to_record,
    apply_nontracked_payload_to_record,
    get_pending_contract_status_entry,
    get_pending_contract_type_entry,
    get_rollback_nontracked_payload,
    is_snapshot_v2,
)

from .models import DeclinedChange, PendingChange, SuperadminState


if TYPE_CHECKING:
    from apps.accounts.models import User


# =============================================================================
# Type Aliases
# =============================================================================

ChangeClassification = Literal["SAFE", "RISKY"]


# =============================================================================
# Constants
# =============================================================================

# Re-export for convenience
RISK_TRACKED_FIELDS: tuple[str, ...] = TRACKED_FIELDS


# =============================================================================
# Value Normalization Helpers
# =============================================================================

def _normalize_value(value: Any) -> Any:
    """
    Normalize a field value for comparison.

    Handles:
    - None vs empty string (both normalize to None for strings, None for others)
    - Decimal precision normalization
    - Date string conversion
    - Whitespace stripping for strings

    Args:
        value: The field value to normalize.

    Returns:
        Normalized value suitable for comparison.
    """
    if value is None:
        return None

    # String normalization
    if isinstance(value, str):
        stripped = value.strip()
        # Treat empty strings as None for comparison purposes
        return stripped if stripped else None

    # Decimal normalization (remove trailing zeros for comparison)
    if isinstance(value, Decimal):
        # Normalize to remove trailing zeros: Decimal("10.00") -> Decimal("10")
        return value.normalize() if value else None

    # Date normalization (convert to ISO string for consistent comparison)
    if isinstance(value, date):
        return value.isoformat()

    return value


def _values_equal(value1: Any, value2: Any) -> bool:
    """
    Compare two values with normalization.

    This handles edge cases like:
    - None vs ""
    - Decimal("10.00") vs Decimal("10")
    - Date objects vs date strings

    Args:
        value1: First value.
        value2: Second value.

    Returns:
        True if values are considered equal after normalization.
    """
    norm1 = _normalize_value(value1)
    norm2 = _normalize_value(value2)
    return norm1 == norm2


# =============================================================================
# Risk Classification Functions
# =============================================================================

def is_risky_change(
    local: KarteiRecord,
    original: KarteiRecord | None,
) -> bool:
    """
    Determine if a change is risky and requires Superadmin approval.

    This function implements the TRACKED_FIELDS risk policy:
    - New records (original is None): Always SAFE (False)
    - Existing records: RISKY (True) if any tracked field changed

    Tracked fields are defined in DOMAIN_MODEL.md Section 4:
    - family_id, parent_name, child_name, birthdate, address
    - phone, mobile, email
    - subject1, price1, subject2, price2
    - month_1 through month_12
    - extra1, extra2, extra3

    Args:
        local: The modified KarteiRecord (current state).
        original: The original KarteiRecord from database, or None for new records.

    Returns:
        True if the change is risky (should go to pending), False if safe.

    Example:
        >>> record = KarteiRecord.objects.get(id=123, year=2025)
        >>> modified = ... # record with changes
        >>> if is_risky_change(modified, record):
        ...     create_or_update_pending_change(modified)
        ... else:
        ...     record.save()  # safe change, apply directly
    """
    # New records are always safe
    if original is None:
        return False

    # Compare each tracked field
    for field_name in TRACKED_FIELDS:
        local_value = getattr(local, field_name, None)
        original_value = getattr(original, field_name, None)

        if not _values_equal(local_value, original_value):
            # Found a tracked field that changed -> risky
            return True

    # No tracked fields changed -> safe
    return False


def classify_change(
    local: KarteiRecord,
    original: KarteiRecord | None,
) -> ChangeClassification:
    """
    Classify a change as SAFE or RISKY.

    Convenience wrapper around is_risky_change that returns a string literal
    instead of a boolean.

    Args:
        local: The modified KarteiRecord.
        original: The original KarteiRecord, or None for new records.

    Returns:
        "SAFE" if the change can be applied directly,
        "RISKY" if it requires Superadmin approval.

    Example:
        >>> classification = classify_change(modified_record, original_record)
        >>> if classification == "RISKY":
        ...     # Create pending change
        ... else:
        ...     # Apply directly
    """
    return "RISKY" if is_risky_change(local, original) else "SAFE"


def get_changed_tracked_fields(
    local: KarteiRecord,
    original: KarteiRecord | None,
) -> dict[str, tuple[Any, Any]]:
    """
    Get a dictionary of all changed tracked fields.

    Useful for building history entries and displaying "War/Ist" (was/is)
    comparisons in the Superadmin UI.

    Args:
        local: The modified KarteiRecord.
        original: The original KarteiRecord, or None for new records.

    Returns:
        Dictionary mapping field_name -> (original_value, new_value)
        for each tracked field that changed.
        Empty dict if original is None (new record).

    Example:
        >>> changes = get_changed_tracked_fields(modified, original)
        >>> for field, (old_val, new_val) in changes.items():
        ...     print(f"{field}: {old_val} -> {new_val}")
    """
    if original is None:
        return {}

    changes: dict[str, tuple[Any, Any]] = {}

    for field_name in TRACKED_FIELDS:
        local_value = getattr(local, field_name, None)
        original_value = getattr(original, field_name, None)

        if not _values_equal(local_value, original_value):
            changes[field_name] = (original_value, local_value)

    return changes


# =============================================================================
# Snapshot Building
# =============================================================================

def build_snapshot(record: KarteiRecord) -> dict[str, Any]:
    """
    Build a JSON-serializable snapshot of tracked fields from a KarteiRecord.

    This snapshot is stored in PendingChange.snapshot and DeclinedChange.snapshot.

    Args:
        record: The KarteiRecord to snapshot.

    Returns:
        Dictionary with all tracked field values, JSON-serializable.
        Dates are converted to ISO strings, Decimals to strings.
    """
    return build_tracked_snapshot(record)


def _build_diff_string(old_snapshot: dict[str, Any], new_snapshot: dict[str, Any]) -> str:
    """
    Build a diff string in the format TAG(old->new);TAG(old->new);...
    
    Compares tracked fields between old and new snapshots and generates
    history tags for changed fields.
    
    Args:
        old_snapshot: Snapshot before changes (from build_snapshot).
        new_snapshot: Snapshot after changes (from PendingChange.snapshot).
    
    Returns:
        String of field changes in history format, e.g., "M01(0->48.30);PR1(0->46.00);"
        Returns empty string if no changes detected.
    """
    from apps.karteien.models import HISTORY_FIELD_TAGS
    
    diff_parts = []
    
    for field_name in TRACKED_FIELDS:
        old_val = old_snapshot.get(field_name)
        new_val = new_snapshot.get(field_name)
        
        # Compare using normalized values
        if not _values_equal(old_val, new_val):
            # Get history tag for this field
            tag = HISTORY_FIELD_TAGS.get(field_name)
            if tag:
                # Format values for display
                # For empty/None values, use "0" for numeric fields, empty string for others
                if old_val in (None, ""):
                    old_display = "0" if field_name.startswith(("month_", "price")) else ""
                else:
                    old_display = str(old_val)
                
                if new_val in (None, ""):
                    new_display = "0" if field_name.startswith(("month_", "price")) else ""
                else:
                    new_display = str(new_val)
                
                # Escape special characters in values (parentheses and semicolons)
                # to prevent parsing issues
                old_display = old_display.replace("(", "[").replace(")", "]").replace(";", ",")
                new_display = new_display.replace("(", "[").replace(")", "]").replace(";", ",")
                
                # Add to diff string
                diff_parts.append(f"{tag}({old_display}->{new_display})")
    
    return ";".join(diff_parts) + ";" if diff_parts else ""


# =============================================================================
# Pending/Declined Change Management
# =============================================================================

def create_or_update_pending_change(
    record: KarteiRecord,
    admin_comment: str | None = None,
) -> PendingChange:
    """
    Create or update a pending change for a KarteiRecord.

    If a pending change already exists for this record, it will be updated
    with the new snapshot. Otherwise, a new pending change is created.

    This should be called after is_risky_change returns True.

    Also creates notifications for all Superadmin users about the new/updated
    pending change.

    Args:
        record: The KarteiRecord with proposed changes.
        admin_comment: Optional comment provided by Admin explaining the change.
                       Will be stored in PendingChange.admin_comment.

    Returns:
        The created or updated PendingChange instance.

    Raises:
        ValueError: If record has no primary key (not saved).

    Example:
        >>> if is_risky_change(modified, original):
        ...     pending = create_or_update_pending_change(
        ...         modified,
        ...         admin_comment="Preis korrigiert lt. Vertrag",
        ...     )
        ...     # Record status will be set to PENDING
    """
    if record.pk is None:
        raise ValueError("Cannot create pending change for unsaved record")

    snapshot = build_snapshot(record)

    defaults: dict[str, Any] = {
        "snapshot": snapshot,
        "is_processed": False,
    }
    
    # Include admin_comment if provided
    if admin_comment is not None:
        defaults["admin_comment"] = admin_comment

    pending, created = PendingChange.objects.update_or_create(
        record=record,
        defaults=defaults,
    )

    # Create notifications for Superadmins
    from apps.notifications.services import notify_pending_created
    notify_pending_created(record, pending)

    # TODO: In future steps:
    # - Update record.status to PENDING
    # - Log the action for audit

    return pending


def create_or_update_pending_change_from_snapshot(
    record: KarteiRecord,
    snapshot: dict[str, Any],
    admin_comment: str | None = None,
) -> PendingChange:
    """
    Create or update a pending change with an explicit snapshot.

    Unlike create_or_update_pending_change which builds snapshot from record,
    this function uses the provided snapshot directly. This is used when
    resubmitting a declined change that was edited - we want to use the
    corrected snapshot from DeclinedChange, not the current record state.

    Args:
        record: The KarteiRecord to create pending change for.
        snapshot: The snapshot dict with proposed field values.
        admin_comment: Optional comment provided by Admin explaining the change.
                       Will be stored in PendingChange.admin_comment.

    Returns:
        The created or updated PendingChange instance.

    Raises:
        ValueError: If record has no primary key (not saved).

    Example:
        >>> # Resubmit corrected declined change
        >>> pending = create_or_update_pending_change_from_snapshot(
        ...     record=declined.record,
        ...     snapshot=declined.snapshot,  # edited snapshot
        ...     admin_comment="Korrektur nach Rücksprache",
        ... )
    """
    if record.pk is None:
        raise ValueError("Cannot create pending change for unsaved record")

    defaults: dict[str, Any] = {
        "snapshot": snapshot,
        "is_processed": False,
    }
    
    # Include admin_comment if provided
    if admin_comment is not None:
        defaults["admin_comment"] = admin_comment

    pending, created = PendingChange.objects.update_or_create(
        record=record,
        defaults=defaults,
    )

    # Create notifications for Superadmins
    from apps.notifications.services import notify_pending_created
    notify_pending_created(record, pending)

    return pending


def create_declined_change(
    record: KarteiRecord,
    reason: str,
    declined_by: User | None = None,
) -> DeclinedChange:
    """
    Create a declined change record.

    This is called when a Superadmin declines a pending change. The pending
    change data is moved to a declined change record with the decline reason.

    NOTE: This is a skeleton implementation. Full integration with
    PendingChange deletion, status updates, and notifications will be added later.

    Args:
        record: The KarteiRecord whose changes were declined.
        reason: Reason for declining, provided by Superadmin.
        declined_by: The Superadmin user who declined (optional).

    Returns:
        The created DeclinedChange instance.

    Raises:
        ValueError: If record has no primary key (not saved).

    Example:
        >>> # Superadmin declines a pending change
        >>> declined = create_declined_change(
        ...     record=kartei_record,
        ...     reason="Invalid address format",
        ...     declined_by=superadmin_user,
        ... )
    """
    if record.pk is None:
        raise ValueError("Cannot create declined change for unsaved record")

    # Try to get snapshot from existing pending change, or build new one
    try:
        pending = PendingChange.objects.get(record=record, is_processed=False)
        snapshot = pending.snapshot
    except PendingChange.DoesNotExist:
        # No pending change, build snapshot from record
        snapshot = build_snapshot(record)

    declined = DeclinedChange.objects.create(
        record=record,
        snapshot=snapshot,
        decline_reason=reason,
        declined_by=declined_by,
    )

    # Create notifications for Admins
    from apps.notifications.services import notify_declined_created
    notify_declined_created(record, declined)

    # TODO: In future steps:
    # - Delete or mark pending change as processed
    # - Update record.status to DECLINED
    # - Write DCL entry to history_raw

    return declined


def apply_pending_change(pending: PendingChange) -> KarteiRecord:
    """
    Apply an approved pending change to the original record.

    This is called when a Superadmin approves a pending change.

    NOTE: This is a skeleton implementation. Full integration with
    history writing and notifications will be added later.

    Args:
        pending: The approved PendingChange to apply.

    Returns:
        The updated KarteiRecord.

    Example:
        >>> # Superadmin approves
        >>> record = apply_pending_change(pending_change)
        >>> # record now has the approved values
    """
    record = pending.record
    snapshot = pending.snapshot

    # Apply snapshot values to record
    for field_name, value in snapshot.items():
        if field_name in TRACKED_FIELDS:
            # Convert string representations back to proper types
            field = record._meta.get_field(field_name)

            if value is not None:
                if isinstance(field, (type(record._meta.get_field("price1")),)):
                    # DecimalField
                    value = Decimal(value) if value else None
                elif isinstance(field, (type(record._meta.get_field("birthdate")),)):
                    # DateField - parse ISO string
                    if isinstance(value, str):
                        from datetime import datetime
                        value = datetime.strptime(value, "%Y-%m-%d").date()

            setattr(record, field_name, value)

    # Mark pending as processed
    pending.is_processed = True
    pending.save(update_fields=["is_processed", "updated_at"])

    # TODO: In future steps:
    # - Update record.status to NORMAL
    # - Write APR entry to history_raw
    # - Create HistoryEvent
    # - Save record

    return record


# =============================================================================
# Decision Application (Superadmin Workflow)
# =============================================================================

def apply_decision(
    pending: PendingChange,
    decision: str,
    comment: str | None,
    user: User | None = None,
) -> KarteiRecord:
    """
    Apply a Superadmin decision to a pending change.

    This is the main entry point for processing Superadmin decisions,
    mirroring the VBA valid_ApproveFlow.SyncDecisions logic.

    Process:
    - APPROVED: Apply changes to KarteiRecord, update status to NORMAL
    - DECLINED: Create DeclinedChange, update status to DECLINED

    Both cases:
    - Mark pending as processed
    - Create notifications for Admin/Operator
    - Write to history (APR/DCL tags)

    Args:
        pending: The PendingChange to process.
        decision: "APPROVED" or "DECLINED".
        comment: Optional comment (required for decline, informational for approve).
        user: The Superadmin user making the decision.

    Returns:
        The updated KarteiRecord.

    Raises:
        ValueError: If decision is invalid or pending already processed.

    Example:
        >>> pending = PendingChange.objects.get(id=123)
        >>> record = apply_decision(pending, "APPROVED", None, superadmin)
    """
    from django.db import transaction
    from apps.notifications.services import (
        notify_approved,
        notify_declined_created,
        mark_pending_notifications_read_for_record,
    )

    decision = decision.upper()
    if decision not in ("APPROVED", "DECLINED"):
        raise ValueError(f"Invalid decision: {decision}. Must be APPROVED or DECLINED.")

    if pending.is_processed:
        raise ValueError(f"PendingChange {pending.id} is already processed.")

    record = pending.record

    # Build combined comment: Admin comment + Superadmin comment
    admin_comment = getattr(pending, "admin_comment", "") or ""
    admin_comment = admin_comment.strip()
    superadmin_comment = (comment or "").strip()

    # Combine comments for history entry
    if admin_comment and superadmin_comment:
        combined_comment = f"Admin: {admin_comment}; Superadmin: {superadmin_comment}"
    elif admin_comment:
        combined_comment = f"Admin: {admin_comment}"
    elif superadmin_comment:
        combined_comment = superadmin_comment
    else:
        combined_comment = None

    with transaction.atomic():
        if decision == "APPROVED":
            # Build snapshot of current state BEFORE applying changes
            old_snapshot = build_snapshot(record)
            new_snapshot = pending.snapshot
            
            # Compute diff for history entry
            diff_string = _build_diff_string(old_snapshot, new_snapshot)
            
            # Apply changes from snapshot to record
            record = _apply_snapshot_to_record(record, new_snapshot)
            
            # Update status to NORMAL
            record.status = RecordStatus.NORMAL
            record.save()

            # ── Create ContractTypeEntry / ContractStatusEntry from pending metadata ──
            _create_contract_entries_from_snapshot(record, new_snapshot)

            # Mark pending as processed
            pending.is_processed = True
            pending.save(update_fields=["is_processed", "updated_at"])

            # Write APR entry to history (append to history_raw) with diff
            _write_history_entry(record, "APR", user, combined_comment, diff_string)

            # Create notifications for Admin
            try:
                notify_approved(record, pending, user)
            except Exception:
                pass  # Don't fail if notification fails

            # Mark all PENDING_CREATED notifications for this record as read
            # so they disappear from Superadmins' notification lists
            # NOTE: Use record.pkid (Django PK), NOT record.id (domain key)
            try:
                mark_pending_notifications_read_for_record(record.pkid)
            except Exception:
                pass  # Don't fail if notification cleanup fails

        elif decision == "DECLINED":
            if not comment:
                comment = "Keine Begründung angegeben"
            
            # Recalculate combined_comment for decline (use actual decline reason)
            if admin_comment and comment:
                combined_comment = f"Admin: {admin_comment}; Superadmin: {comment}"
            elif admin_comment:
                combined_comment = f"Admin: {admin_comment}; Superadmin: {comment}"
            else:
                combined_comment = comment

            # Create DeclinedChange with snapshot
            declined = DeclinedChange.objects.create(
                record=record,
                snapshot=pending.snapshot,
                decline_reason=comment,
                declined_by=user,
            )

            # Update record status to DECLINED
            record.status = RecordStatus.DECLINED
            pending_snapshot = (
                pending.snapshot
                if isinstance(pending.snapshot, dict)
                else {}
            )
            rollback_payload = (
                get_rollback_nontracked_payload(pending_snapshot)
                if is_snapshot_v2(pending_snapshot)
                else None
            )
            if rollback_payload:
                rollback_updates = {
                    key: value
                    for key, value in rollback_payload.items()
                    if key not in TRACKED_FIELDS
                }
                rollback_updates["status"] = RecordStatus.DECLINED
                apply_nontracked_payload_to_record(record, rollback_payload)
                KarteiRecord.objects.filter(pk=record.pk).update(**rollback_updates)
            elif not is_snapshot_v2(pending_snapshot) and (
                pending_snapshot.get("_old_base_amounts") is not None
            ):
                record.base_amounts = pending_snapshot["_old_base_amounts"]
                record.save(update_fields=["status", "base_amounts"])
            else:
                record.save(update_fields=["status"])

            # Mark pending as processed
            pending.is_processed = True
            pending.save(update_fields=["is_processed", "updated_at"])

            # Write DCL entry to history
            _write_history_entry(record, "DCL", user, combined_comment)

            # Create notifications for Admin
            try:
                notify_declined_created(record, declined)
            except Exception:
                pass  # Don't fail if notification fails

            # Mark all PENDING_CREATED notifications for this record as read
            # so they disappear from Superadmins' notification lists
            # NOTE: Use record.pkid (Django PK), NOT record.id (domain key)
            try:
                mark_pending_notifications_read_for_record(record.pkid)
            except Exception:
                pass  # Don't fail if notification cleanup fails

    return record


def _create_contract_entries_from_snapshot(
    record: KarteiRecord,
    snapshot: dict[str, Any],
) -> None:
    """
    Create ContractTypeEntry / ContractStatusEntry from pending metadata.

    Called only on APPROVE. The snapshot may contain special keys
    ``_pending_contract_type_entry`` and ``_pending_contract_status_entry``
    that describe timeline entries requested by the Admin/Operator.
    These keys are **not** tracked fields and are silently ignored by
    ``_apply_snapshot_to_record``.

    Each value is a dict with at least ``effective_from_month`` and the
    type-specific payload (``is_monthly`` or ``kind``).
    """
    from apps.karteien.models import ContractTypeEntry, ContractStatusEntry

    meta_type = get_pending_contract_type_entry(snapshot)
    if meta_type and isinstance(meta_type, dict):
        ContractTypeEntry.objects.update_or_create(
            record=record,
            effective_from_month=int(meta_type["effective_from_month"]),
            defaults={
                "is_monthly": meta_type.get("is_monthly", False),
                "changed_by_id": meta_type.get("changed_by_id"),
                "comment": meta_type.get("comment", ""),
            },
        )

    meta_status = get_pending_contract_status_entry(snapshot)
    if meta_status and isinstance(meta_status, dict):
        ContractStatusEntry.objects.update_or_create(
            record=record,
            effective_from_month=int(meta_status["effective_from_month"]),
            defaults={
                "kind": meta_status["kind"],
                "changed_by_id": meta_status.get("changed_by_id"),
                "comment": meta_status.get("comment", ""),
            },
        )


def _apply_snapshot_to_record(
    record: KarteiRecord,
    snapshot: dict[str, Any],
) -> KarteiRecord:
    """
    Apply a snapshot dictionary to a KarteiRecord.

    Snapshot v1 applies top-level tracked fields only. Snapshot v2 also
    applies the normalized `_pending_nontracked_payload` in memory.

    Args:
        record: The record to update.
        snapshot: Dictionary of field_name -> value.

    Returns:
        The updated record (not saved).
    """
    return apply_pending_snapshot_to_record(record, snapshot)


def _write_history_entry(
    record: KarteiRecord,
    entry_type: str,
    user: User | None,
    comment: str | None,
    diff_string: str = "",
) -> None:
    """
    Append an entry to the record's history_raw field.

    Entry types:
    - APR: Approved by Superadmin
    - DCL: Declined by Superadmin
    - ADM: Admin comment on SAFE change

    Format matches VBA Export_HistoryBuilder:
    - APR: APR:<user>/[TAG(old->new);...]/@<comment>@/<date>||
    - DCL: DCL(<N>-><comment>)/@<date>@/||
    - ADM: ADM:<user>/[TAG(old->new);...]/@<comment>@/<date>||

    Args:
        record: The record to update history for.
        entry_type: "APR", "DCL", or "ADM".
        user: The user who made the decision.
        comment: Optional comment.
        diff_string: Optional diff string with field changes (TAG(old->new);...).
    """
    from django.utils import timezone

    now = timezone.now()
    date_str = now.strftime("%d.%m.%Y")
    
    user_name = user.username if user else "System"

    if entry_type == "APR":
        # Format: APR:user/[diff]/@comment@/date||
        parts = [f"APR:{user_name}/"]
        if diff_string:
            parts.append(diff_string)
        if comment:
            parts.append(f"/@{comment}@/")
        parts.append(f"{date_str}||")
        entry = "".join(parts)
    elif entry_type == "DCL":
        # Count existing declines for numbering
        dcl_count = record.declined_changes.count()
        entry = f"DCL({dcl_count}->{comment or 'Abgelehnt'})/@{user_name}@/{date_str}||"
    elif entry_type == "ADM":
        # Admin comment on SAFE change (direct save without approval)
        if comment or diff_string:
            parts = [f"ADM:{user_name}/"]
            if diff_string:
                parts.append(diff_string)
            if comment:
                parts.append(f"/@{comment}@/")
            parts.append(f"{date_str}||")
            entry = "".join(parts)
        else:
            # No comment or diff, no entry needed
            return
    else:
        return

    # Append to history_raw
    if record.history_raw:
        record.history_raw = record.history_raw + entry
    else:
        record.history_raw = entry

    record.save(update_fields=["history_raw"])


def write_history_entry(
    record: KarteiRecord,
    entry_type: str,
    user: User | None,
    comment: str | None,
    diff_string: str = "",
) -> None:
    """
    Public wrapper to append an entry to the record's history_raw field.

    Entry types:
    - APR: Approved by Superadmin
    - DCL: Declined by Superadmin
    - ADM: Admin comment on SAFE change

    Format uses /@...@/ markers for new-format recognition:
    - APR: APR:<user>/[TAG(old->new);...]/@<comment>@/<date>||
    - DCL: DCL(<N>-><comment>)/@<user>@/<date>||
    - ADM: ADM:<user>/[TAG(old->new);...]/@<comment>@/<date>||

    Args:
        record: The record to update history for.
        entry_type: "APR", "DCL", or "ADM".
        user: The user who made the decision.
        comment: Optional comment (for ADM, entry is skipped if empty).
        diff_string: Optional diff string with field changes (TAG(old->new);...).
    """
    _write_history_entry(record, entry_type, user, comment, diff_string)


# =============================================================================
# Bulk Operations (Superadmin)
# =============================================================================

def approve_all_pending(
    user: User | None = None,
    year: int | None = None,
) -> tuple[int, list[str]]:
    """
    Approve all pending changes.

    Mirrors VBA valid_ApproveFlow.ApproveAllPending.

    Args:
        user: The Superadmin user.
        year: Optional year filter.

    Returns:
        Tuple of (approved_count, error_messages).
    """
    from django.db import transaction

    qs = PendingChange.objects.filter(is_processed=False).select_related("record")
    if year:
        qs = qs.filter(record__year=year)

    approved_count = 0
    errors: list[str] = []

    for pending in qs:
        try:
            with transaction.atomic():
                apply_decision(pending, "APPROVED", None, user)
            approved_count += 1
        except Exception as e:
            errors.append(f"ID {pending.record_id}: {str(e)}")

    return approved_count, errors


def decline_all_pending(
    comment: str,
    user: User | None = None,
    year: int | None = None,
) -> tuple[int, list[str]]:
    """
    Decline all pending changes with a common comment.

    Mirrors VBA valid_ApproveFlow.DeclineAllPending.

    Args:
        comment: Common decline reason for all records.
        user: The Superadmin user.
        year: Optional year filter.

    Returns:
        Tuple of (declined_count, error_messages).
    """
    from django.db import transaction

    qs = PendingChange.objects.filter(is_processed=False).select_related("record")
    if year:
        qs = qs.filter(record__year=year)

    declined_count = 0
    errors: list[str] = []

    for pending in qs:
        try:
            with transaction.atomic():
                apply_decision(pending, "DECLINED", comment, user)
            declined_count += 1
        except Exception as e:
            errors.append(f"ID {pending.record_id}: {str(e)}")

    return declined_count, errors


def clear_all_decisions(year: int | None = None) -> int:
    """
    Clear all unprocessed decisions (reset to pending state without action).

    This is mainly for testing/development. In production, processed
    decisions should remain for audit trail.

    Args:
        year: Optional year filter.

    Returns:
        Number of cleared pending changes.
    """
    qs = PendingChange.objects.filter(is_processed=False)
    if year:
        qs = qs.filter(record__year=year)

    # Just delete pending changes, restoring records to NORMAL
    count = 0
    for pending in qs:
        record = pending.record
        record.status = RecordStatus.NORMAL
        record.save(update_fields=["status"])
        pending.delete()
        count += 1

    return count


# =============================================================================
# NeuList Functions
# =============================================================================

def get_or_create_superadmin_state(user: User) -> SuperadminState:
    """
    Get or create SuperadminState for a user.

    Args:
        user: The Superadmin user.

    Returns:
        SuperadminState instance.
    """
    state, _ = SuperadminState.objects.get_or_create(user=user)
    return state


def _get_last_seen_id_for_year(state: SuperadminState, year: int) -> int:
    """
    Get the last_seen_id for a specific year.

    Uses the per-year tracking in last_seen_by_year if available.
    If the year key doesn't exist, returns 0 (meaning no records have been
    "seen" for this year yet).

    Note: We intentionally do NOT fall back to the legacy last_seen_id field,
    because that value may come from a different year and cause incorrect
    filtering (e.g., legacy last_seen_id=1000 from year 2025 would hide
    all records with id<=1000 in year 2026, even though IDs restart per year).

    Args:
        state: The SuperadminState instance.
        year: The year to get the last_seen_id for.

    Returns:
        The last_seen_id for the given year, or 0 if not yet tracked.
    """
    year_key = str(year)
    if state.last_seen_by_year and year_key in state.last_seen_by_year:
        return state.last_seen_by_year[year_key]
    # Year not yet tracked - return 0 (show all records as "new" for this year)
    return 0


def get_new_records(user: User, year: int) -> list[KarteiRecord]:
    """
    Get records that are "new" for the given Superadmin and year.

    A record is "new" if its ID is greater than the user's per-year last_seen_id.
    Mirrors VBA valid_NeuList.RefreshNeuList logic.

    Note: Since record IDs are only unique within a year (domain key is (year, id)),
    we now track last_seen_id per-year in last_seen_by_year JSONField.

    Args:
        user: The Superadmin user.
        year: The year to filter by (required).

    Returns:
        List of new KarteiRecord instances for the given year.
    """
    state = get_or_create_superadmin_state(user)
    last_seen_id = _get_last_seen_id_for_year(state, year)

    qs = KarteiRecord.objects.filter(year=year, id__gt=last_seen_id).order_by("id")

    return list(qs)


def get_new_records_count(user: User, year: int) -> int:
    """
    Get count of new records for a Superadmin and year.

    Args:
        user: The Superadmin user.
        year: The year to filter by (required).

    Returns:
        Count of new records for the given year.
    """
    state = get_or_create_superadmin_state(user)
    last_seen_id = _get_last_seen_id_for_year(state, year)

    return KarteiRecord.objects.filter(year=year, id__gt=last_seen_id).count()


def update_last_seen_id(user: User, year: int, max_id: int | None = None) -> None:
    """
    Update the last seen ID for a Superadmin for a specific year.

    If max_id is None, it's set to the current maximum ID for the given year.
    Updates both the per-year tracking (last_seen_by_year) and legacy last_seen_id.

    Args:
        user: The Superadmin user.
        year: The year to update last_seen_id for (required).
        max_id: The new last_seen_id value, or None to use max for the year.
    """
    from django.db.models import Max
    from django.utils import timezone

    state = get_or_create_superadmin_state(user)

    if max_id is None:
        result = KarteiRecord.objects.filter(year=year).aggregate(Max("id"))
        max_id = result["id__max"] or 0

    # Update per-year tracking
    if state.last_seen_by_year is None:
        state.last_seen_by_year = {}
    state.last_seen_by_year[str(year)] = max_id

    # Also update legacy field for compatibility/debugging
    state.last_seen_id = max_id
    state.last_seen_date = timezone.now()
    state.save(update_fields=["last_seen_id", "last_seen_by_year", "last_seen_date", "updated_at"])
