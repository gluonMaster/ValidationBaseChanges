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
    snapshot: dict[str, Any] = {}

    for field_name in TRACKED_FIELDS:
        value = getattr(record, field_name, None)

        # Convert non-JSON-serializable types
        if isinstance(value, date):
            snapshot[field_name] = value.isoformat()
        elif isinstance(value, Decimal):
            snapshot[field_name] = str(value) if value is not None else None
        else:
            snapshot[field_name] = value

    return snapshot


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
            # Apply changes from snapshot to record
            record = _apply_snapshot_to_record(record, pending.snapshot)
            
            # Update status to NORMAL
            record.status = RecordStatus.NORMAL
            record.save()

            # Mark pending as processed
            pending.is_processed = True
            pending.save(update_fields=["is_processed", "updated_at"])

            # Write APR entry to history (append to history_raw)
            _write_history_entry(record, "APR", user, combined_comment)

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


def _apply_snapshot_to_record(
    record: KarteiRecord,
    snapshot: dict[str, Any],
) -> KarteiRecord:
    """
    Apply a snapshot dictionary to a KarteiRecord.

    Converts string representations back to proper Django field types.

    Args:
        record: The record to update.
        snapshot: Dictionary of field_name -> value.

    Returns:
        The updated record (not saved).
    """
    for field_name, value in snapshot.items():
        if field_name not in TRACKED_FIELDS:
            continue

        field = record._meta.get_field(field_name)

        if value is not None:
            # DecimalField
            if field_name in ("price1", "price2") or field_name.startswith("month_"):
                if isinstance(value, str) and value:
                    value = Decimal(value)
                elif not value:
                    value = None
            # DateField
            elif field_name == "birthdate":
                if isinstance(value, str) and value:
                    from datetime import datetime as dt
                    value = dt.strptime(value, "%Y-%m-%d").date()

        setattr(record, field_name, value)

    return record


def _write_history_entry(
    record: KarteiRecord,
    entry_type: str,
    user: User | None,
    comment: str | None,
) -> None:
    """
    Append an entry to the record's history_raw field.

    Entry types:
    - APR: Approved by Superadmin
    - DCL: Declined by Superadmin
    - ADM: Admin comment on SAFE change

    Format matches VBA Export_HistoryBuilder:
    - APR: APR/@<comment>@/<date>||
    - DCL: DCL(<N>-><comment>)/@<date>@/||
    - ADM: ADM:<user>/@<comment>@/<date>||

    Args:
        record: The record to update history for.
        entry_type: "APR", "DCL", or "ADM".
        user: The user who made the decision.
        comment: Optional comment.
    """
    from django.utils import timezone

    now = timezone.now()
    date_str = now.strftime("%d.%m.%Y")
    
    user_name = user.username if user else "System"

    if entry_type == "APR":
        if comment:
            entry = f"APR:{user_name}/@{comment}@/{date_str}||"
        else:
            entry = f"APR:{user_name}/{date_str}||"
    elif entry_type == "DCL":
        # Count existing declines for numbering
        dcl_count = record.declined_changes.count()
        entry = f"DCL({dcl_count}->{comment or 'Abgelehnt'})/@{user_name}@/{date_str}||"
    elif entry_type == "ADM":
        # Admin comment on SAFE change (direct save without approval)
        if comment:
            entry = f"ADM:{user_name}/@{comment}@/{date_str}||"
        else:
            # No comment, no entry needed
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
) -> None:
    """
    Public wrapper to append an entry to the record's history_raw field.

    Entry types:
    - APR: Approved by Superadmin
    - DCL: Declined by Superadmin
    - ADM: Admin comment on SAFE change

    Format uses /@...@/ markers for new-format recognition:
    - APR: APR:<user>/@<comment>@/<date>||
    - DCL: DCL(<N>-><comment>)/@<user>@/<date>||
    - ADM: ADM:<user>/@<comment>@/<date>||

    Args:
        record: The record to update history for.
        entry_type: "APR", "DCL", or "ADM".
        user: The user who made the decision.
        comment: Optional comment (for ADM, entry is skipped if empty).
    """
    _write_history_entry(record, entry_type, user, comment)


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


def get_new_records(user: User, year: int | None = None) -> list[KarteiRecord]:
    """
    Get records that are "new" for the given Superadmin.

    A record is "new" if its ID is greater than the user's last_seen_id.
    Mirrors VBA valid_NeuList.RefreshNeuList logic.

    Args:
        user: The Superadmin user.
        year: Optional year filter.

    Returns:
        List of new KarteiRecord instances.
    """
    state = get_or_create_superadmin_state(user)
    last_seen_id = state.last_seen_id

    qs = KarteiRecord.objects.filter(id__gt=last_seen_id).order_by("id")
    if year:
        qs = qs.filter(year=year)

    return list(qs)


def get_new_records_count(user: User, year: int | None = None) -> int:
    """
    Get count of new records for a Superadmin.

    Args:
        user: The Superadmin user.
        year: Optional year filter.

    Returns:
        Count of new records.
    """
    state = get_or_create_superadmin_state(user)
    
    qs = KarteiRecord.objects.filter(id__gt=state.last_seen_id)
    if year:
        qs = qs.filter(year=year)
    
    return qs.count()


def update_last_seen_id(user: User, max_id: int | None = None) -> None:
    """
    Update the last seen ID for a Superadmin.

    If max_id is None, it's set to the current maximum ID in the database.

    Args:
        user: The Superadmin user.
        max_id: The new last_seen_id value, or None to use max.
    """
    from django.db.models import Max
    from django.utils import timezone

    state = get_or_create_superadmin_state(user)

    if max_id is None:
        result = KarteiRecord.objects.aggregate(Max("id"))
        max_id = result["id__max"] or 0

    state.last_seen_id = max_id
    state.last_seen_date = timezone.now()
    state.save(update_fields=["last_seen_id", "last_seen_date", "updated_at"])
