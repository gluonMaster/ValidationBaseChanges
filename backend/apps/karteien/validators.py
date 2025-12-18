"""
Validators for the karteien app.

This module contains validation logic for Kartei operations, analogous to
VBA modules in the legacy Excel/Access system:

- ExportProtection.bas: SEPA restrictions, past-months validation
- Export_ValidationKartei.bas: FamilyID duplicate checks

Key validation rules:
1. FamilyID + Parent uniqueness
2. Empty FamilyID with non-empty Parent is not allowed
3. SEPA restrictions for Operator role
4. Past-months restrictions for Operator role
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from django.core.exceptions import ValidationError as DjangoValidationError
from django.utils.translation import gettext_lazy as _

from .models import KarteiRecord, MONTH_FIELD_NAMES

if TYPE_CHECKING:
    from apps.accounts.models import User


# =============================================================================
# Validation Result Types
# =============================================================================

@dataclass
class ValidationResult:
    """Result of a validation check."""
    is_valid: bool
    errors: list[str]
    warnings: list[str]
    fixed_data: dict[str, Any] | None = None  # Data with fixes applied
    
    @classmethod
    def success(cls, fixed_data: dict[str, Any] | None = None) -> "ValidationResult":
        """Create a successful validation result."""
        return cls(is_valid=True, errors=[], warnings=[], fixed_data=fixed_data)
    
    @classmethod
    def failure(cls, errors: list[str]) -> "ValidationResult":
        """Create a failed validation result."""
        return cls(is_valid=False, errors=errors, warnings=[])
    
    def add_warning(self, warning: str) -> None:
        """Add a warning message."""
        self.warnings.append(warning)


# =============================================================================
# FamilyID Validation
# =============================================================================

def validate_family_id_parent_unique(
    family_id: str,
    parent_name: str,
    year: int,
    exclude_pk: int | None = None,
) -> ValidationResult:
    """
    Check for duplicate FamilyID with different Parent names.
    
    Mirrors VBA: Export_ValidationKartei.CheckDuplicateFamilyIDs
    
    Rule: If a FamilyID already exists in the year with a DIFFERENT parent name,
    it's a conflict that should be reported.
    
    Args:
        family_id: The FamilyID to check.
        parent_name: The parent name for the new/updated record.
        year: The year context.
        exclude_pk: Record pkid (surrogate PK) to exclude (for updates).
        
    Returns:
        ValidationResult with any errors found.
    """
    if not family_id:
        return ValidationResult.success()
    
    # Find other records with the same FamilyID
    qs = KarteiRecord.objects.filter(
        year=year,
        family_id__iexact=family_id,
    )
    
    if exclude_pk is not None:
        qs = qs.exclude(pk=exclude_pk)
    
    # Check if any existing record has a different parent name
    for record in qs:
        existing_parent = (record.parent_name or "").strip().lower()
        new_parent = (parent_name or "").strip().lower()
        
        # If parent names differ, it's a potential conflict
        if existing_parent and new_parent and existing_parent != new_parent:
            error_msg = (
                f"FamilyID '{family_id}' bereits vorhanden mit Eltern '{record.parent_name}'. "
                f"Neue Eingabe hat Eltern '{parent_name}'. "
                f"Bitte prüfen Sie, ob dies korrekt ist."
            )
            return ValidationResult.failure([error_msg])
    
    return ValidationResult.success()


def validate_family_id_not_empty_with_parent(
    family_id: str,
    parent_name: str,
) -> ValidationResult:
    """
    Check that FamilyID is not empty when Parent is filled.
    
    Mirrors VBA: Export_ValidationKartei.CheckEmptyFamilyIDWithParent
    
    Rule: If Parent name is provided, FamilyID must also be provided.
    
    Args:
        family_id: The FamilyID value.
        parent_name: The parent name value.
        
    Returns:
        ValidationResult with any errors found.
    """
    family_id_stripped = (family_id or "").strip()
    parent_name_stripped = (parent_name or "").strip()
    
    if parent_name_stripped and not family_id_stripped:
        error_msg = (
            f"FamilyID ist leer, aber Eltern '{parent_name}' ist ausgefüllt. "
            "Bitte geben Sie eine FamilyID ein."
        )
        return ValidationResult.failure([error_msg])
    
    return ValidationResult.success()


# =============================================================================
# SEPA Restrictions
# =============================================================================

def validate_sepa_restrictions(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
) -> ValidationResult:
    """
    Validate SEPA restrictions for Operator role.
    
    Mirrors VBA: ExportSyncKartei.ValidateOperatorSepaRestrictions
    
    Rule: Operators cannot modify month fields (month_1..month_12) for records
    with sepa_marker == "SEPA".
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        
    Returns:
        ValidationResult with any errors found.
    """
    # Admin has no SEPA restrictions
    if not user.has_sepa_restrictions:
        return ValidationResult.success()
    
    # Check if record has SEPA marker
    if not record.is_sepa:
        return ValidationResult.success()
    
    # Check if any month fields are being changed
    month_changes = [
        field for field in proposed_changes
        if field in MONTH_FIELD_NAMES
    ]
    
    if month_changes:
        error_msg = (
            "Als Operator können Sie Monatsfelder für SEPA-markierte "
            "Datensätze nicht ändern. "
            f"Betroffene Felder: {', '.join(month_changes)}"
        )
        return ValidationResult.failure([error_msg])
    
    return ValidationResult.success()


def filter_sepa_restricted_changes(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
) -> tuple[dict[str, Any], list[str]]:
    """
    Filter out changes that violate SEPA restrictions.
    
    Instead of blocking the save, this function removes restricted fields
    and returns the filtered changes along with warning messages.
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        
    Returns:
        Tuple of (filtered_changes, warning_messages).
    """
    if not user.has_sepa_restrictions or not record.is_sepa:
        return proposed_changes, []
    
    filtered = {}
    warnings = []
    
    for field, value in proposed_changes.items():
        if field in MONTH_FIELD_NAMES:
            warnings.append(
                f"Änderung von '{field}' wurde ignoriert (SEPA-Einschränkung)"
            )
        else:
            filtered[field] = value
    
    return filtered, warnings


# =============================================================================
# Past Months Restrictions
# =============================================================================

def get_allowed_months(
    data_year: int,
    current_date: date | None = None,
) -> tuple[set[str], str]:
    """
    Determine which month fields can be edited based on current date.
    
    Mirrors VBA: ExportProtection.ValidateAndFixPastMonths logic
    
    Rules:
    - If current year > data year: No months can be edited
    - If current year < data year: All months can be edited
    - If current year == data year: Only current month and future can be edited
    
    Args:
        data_year: The year of the record's data.
        current_date: Override for current date (defaults to today).
        
    Returns:
        Tuple of (set of allowed field names, restriction reason or empty string).
    """
    if current_date is None:
        current_date = date.today()
    
    now_year = current_date.year
    now_month = current_date.month
    
    # Data year is in the future - all months allowed
    if now_year < data_year:
        return set(MONTH_FIELD_NAMES), ""
    
    # Data year is in the past - no months allowed
    if now_year > data_year:
        return set(), f"Das Datenjahr {data_year} liegt in der Vergangenheit"
    
    # Same year - only current and future months allowed
    allowed = set()
    for month_num in range(now_month, 13):
        allowed.add(f"month_{month_num}")
    
    return allowed, ""


def validate_past_months_restrictions(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
    current_date: date | None = None,
) -> ValidationResult:
    """
    Validate past-months restrictions for Operator role.
    
    Mirrors VBA: ExportProtection.ValidateAndFixPastMonths
    
    Rule: Operators cannot modify month fields for past months.
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        current_date: Override for current date.
        
    Returns:
        ValidationResult with any errors found.
    """
    # Admin has no past-months restrictions
    if not user.has_past_months_restrictions:
        return ValidationResult.success()
    
    allowed_months, reason = get_allowed_months(record.year, current_date)
    
    # If all months are blocked due to past year
    if not allowed_months and reason:
        # Check if any month fields are being changed
        month_changes = [
            field for field in proposed_changes
            if field in MONTH_FIELD_NAMES
        ]
        
        if month_changes:
            return ValidationResult.failure([
                f"{reason}. Sie können keine Monatsfelder ändern."
            ])
        return ValidationResult.success()
    
    # Check for changes to disallowed months
    blocked_changes = []
    for field in proposed_changes:
        if field in MONTH_FIELD_NAMES and field not in allowed_months:
            # Extract month number for user-friendly message
            month_num = int(field.split("_")[1])
            blocked_changes.append(f"Monat {month_num}")
    
    if blocked_changes:
        return ValidationResult.failure([
            f"Als Operator können Sie vergangene Monate nicht ändern: {', '.join(blocked_changes)}"
        ])
    
    return ValidationResult.success()


def filter_past_months_changes(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
    current_date: date | None = None,
) -> tuple[dict[str, Any], list[str]]:
    """
    Filter out changes to past months for Operator.
    
    Instead of blocking, removes restricted month changes and returns warnings.
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        current_date: Override for current date.
        
    Returns:
        Tuple of (filtered_changes, warning_messages).
    """
    if not user.has_past_months_restrictions:
        return proposed_changes, []
    
    allowed_months, _ = get_allowed_months(record.year, current_date)
    
    filtered = {}
    warnings = []
    
    for field, value in proposed_changes.items():
        if field in MONTH_FIELD_NAMES and field not in allowed_months:
            month_num = int(field.split("_")[1])
            warnings.append(
                f"Änderung von Monat {month_num} wurde ignoriert (vergangener Monat)"
            )
        else:
            filtered[field] = value
    
    return filtered, warnings


# =============================================================================
# Combined Operator Validation
# =============================================================================

def validate_operator_changes(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
    current_date: date | None = None,
) -> ValidationResult:
    """
    Validate all Operator restrictions on a change.
    
    Combines SEPA and past-months validation.
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        current_date: Override for current date.
        
    Returns:
        Combined ValidationResult with all errors.
    """
    errors = []
    warnings = []
    
    # SEPA validation
    sepa_result = validate_sepa_restrictions(record, proposed_changes, user)
    if not sepa_result.is_valid:
        errors.extend(sepa_result.errors)
    warnings.extend(sepa_result.warnings)
    
    # Past months validation
    past_months_result = validate_past_months_restrictions(
        record, proposed_changes, user, current_date
    )
    if not past_months_result.is_valid:
        errors.extend(past_months_result.errors)
    warnings.extend(past_months_result.warnings)
    
    if errors:
        return ValidationResult.failure(errors)
    
    result = ValidationResult.success()
    for warning in warnings:
        result.add_warning(warning)
    return result


def apply_operator_filters(
    record: KarteiRecord,
    proposed_changes: dict[str, Any],
    user: "User",
    current_date: date | None = None,
) -> tuple[dict[str, Any], list[str]]:
    """
    Apply all Operator restrictions by filtering changes.
    
    Combines SEPA and past-months filtering.
    
    Args:
        record: The existing KarteiRecord being modified.
        proposed_changes: Dict of field_name -> new_value.
        user: The user making the change.
        current_date: Override for current date.
        
    Returns:
        Tuple of (filtered_changes, all_warning_messages).
    """
    all_warnings = []
    
    # Apply SEPA filter
    changes, sepa_warnings = filter_sepa_restricted_changes(
        record, proposed_changes, user
    )
    all_warnings.extend(sepa_warnings)
    
    # Apply past months filter
    changes, past_warnings = filter_past_months_changes(
        record, changes, user, current_date
    )
    all_warnings.extend(past_warnings)
    
    return changes, all_warnings


# =============================================================================
# Complete Record Validation
# =============================================================================

def validate_kartei_record(
    data: dict[str, Any],
    year: int,
    user: "User",
    existing_record: KarteiRecord | None = None,
    current_date: date | None = None,
) -> ValidationResult:
    """
    Complete validation for a Kartei record create/update operation.
    
    Runs all relevant validations:
    1. FamilyID + Parent uniqueness
    2. Empty FamilyID with non-empty Parent
    3. Operator restrictions (SEPA, past months)
    
    Args:
        data: The proposed record data.
        year: The year context.
        user: The user making the change.
        existing_record: Existing record for update, None for create.
        current_date: Override for current date.
        
    Returns:
        Combined ValidationResult.
    """
    errors = []
    warnings = []
    
    family_id = data.get("family_id", "")
    parent_name = data.get("parent_name", "")
    
    # FamilyID + Parent uniqueness
    exclude_pk = existing_record.pk if existing_record else None
    unique_result = validate_family_id_parent_unique(
        family_id, parent_name, year, exclude_pk
    )
    if not unique_result.is_valid:
        errors.extend(unique_result.errors)
    
    # Empty FamilyID with non-empty Parent
    empty_result = validate_family_id_not_empty_with_parent(family_id, parent_name)
    if not empty_result.is_valid:
        errors.extend(empty_result.errors)
    
    # Operator restrictions (only for updates)
    if existing_record and user.has_sepa_restrictions:
        # Build changes dict from data vs existing
        changes = {}
        for field in MONTH_FIELD_NAMES:
            new_val = data.get(field)
            old_val = getattr(existing_record, field, None)
            if new_val != old_val:
                changes[field] = new_val
        
        if changes:
            operator_result = validate_operator_changes(
                existing_record, changes, user, current_date
            )
            if not operator_result.is_valid:
                errors.extend(operator_result.errors)
            warnings.extend(operator_result.warnings)
    
    if errors:
        return ValidationResult.failure(errors)
    
    result = ValidationResult.success()
    for warning in warnings:
        result.add_warning(warning)
    return result


# =============================================================================
# Django Form/Serializer Integration
# =============================================================================

class KarteiValidationMixin:
    """
    Mixin for Django forms/serializers that adds Kartei validation.
    
    Usage in a ModelForm:
        class KarteiRecordForm(KarteiValidationMixin, forms.ModelForm):
            class Meta:
                model = KarteiRecord
                fields = [...]
    """
    
    def clean(self) -> dict[str, Any]:
        """Run Kartei-specific validations."""
        cleaned_data = super().clean()  # type: ignore
        
        # Get context (should be set by view)
        user = getattr(self, "user", None)
        year = cleaned_data.get("year") or getattr(self, "year", None)
        existing_record = getattr(self, "instance", None)
        
        if not user or not year:
            return cleaned_data
        
        # Run validation
        result = validate_kartei_record(
            data=cleaned_data,
            year=year,
            user=user,
            existing_record=existing_record if existing_record and existing_record.pk else None,
        )
        
        if not result.is_valid:
            for error in result.errors:
                raise DjangoValidationError(error)
        
        # Store warnings for view to display
        self._validation_warnings = result.warnings
        
        return cleaned_data
