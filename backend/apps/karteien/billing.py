"""
Billing calculation module for KarteiRecord monthly charges.

This module implements the automatic calculation of month_1..month_12 values
based on:
- Semester assignment (1st semester: months 1-6, 2nd semester: months 7-12)
- Subject type (Individual/Nachhilfe = per-hour pricing, otherwise per-month)
- Start month for each semester (months before start = 0.00)
- Discounts (family and record-level, percent and fixed)

Key functions:
- get_semester_month_ranges(year): Get month ranges for each semester
- is_individual_subject(name): Check if subject is Individual type
- is_nachhilfe_subject(name): Check if subject is Nachhilfe type
- round_money_up(value): Round money to 2 decimals, always up
- build_base_amounts(record): Calculate base amounts before discounts
- calculate_month_values(record, ...): Calculate final month values with discounts
- detect_meaningful_changes(original, cleaned_data): Detect billing-relevant changes
- recalculate_legacy_to_auto(record, touched_months): Convert LEGACY to AUTO mode
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .models import KarteiRecord


# =============================================================================
# Constants and Patterns
# =============================================================================

# Default semester boundaries (can be customized per year in future)
DEFAULT_SEMESTER_1_MONTHS = list(range(1, 7))   # [1, 2, 3, 4, 5, 6]
DEFAULT_SEMESTER_2_MONTHS = list(range(7, 13))  # [7, 8, 9, 10, 11, 12]

# Patterns for detecting subject types
# Individual: contains "ind." or starts with "vspe_" (case-insensitive)
INDIVIDUAL_PATTERN = re.compile(r'\bInd\.|VSpE_', re.IGNORECASE)
# Nachhilfe: contains "nachhilfe" or "nh" as standalone token
NACHHILFE_PATTERN = re.compile(r'\bNH\b|Nachhilfe', re.IGNORECASE)

# Decimal constants
ZERO = Decimal('0.00')
MAX_PERCENT_DISCOUNT = Decimal('0.99')


# =============================================================================
# Normalization Helpers
# =============================================================================

def _normalize_subject_name(name: str | None) -> str:
    """
    Normalize subject name for comparison: trim, collapse spaces, casefold.
    
    Args:
        name: Subject name to normalize.
        
    Returns:
        Normalized string, or empty string if input is None/empty.
    """
    if not name:
        return ''
    # Strip leading/trailing whitespace, collapse internal spaces, casefold
    return ' '.join(name.split()).casefold()


# =============================================================================
# Semester Configuration
# =============================================================================

def get_semester_month_ranges(year: int) -> tuple[list[int], list[int]]:
    """
    Get the month ranges for each semester of a given year.
    
    Currently returns default boundaries:
    - 1st semester: months 1-6
    - 2nd semester: months 7-12
    
    In the future, this function can be extended to read from a configuration
    table to support different boundaries per year (e.g., 1-7 and 8-12).
    
    Args:
        year: The year to get semester ranges for.
        
    Returns:
        Tuple of (semester_1_months, semester_2_months) as lists of integers.
    """
    # Future: query YearConfig table for custom boundaries
    # For now, use defaults
    return (DEFAULT_SEMESTER_1_MONTHS.copy(), DEFAULT_SEMESTER_2_MONTHS.copy())


def get_semester_for_month(month: int, year: int) -> int:
    """
    Determine which semester a month belongs to.
    
    Args:
        month: Month number (1-12).
        year: The year for semester configuration.
        
    Returns:
        1 for first semester, 2 for second semester.
        
    Raises:
        ValueError: If month is not in range 1-12.
    """
    if not 1 <= month <= 12:
        raise ValueError(f"Month must be 1-12, got {month}")
    
    sem1_months, sem2_months = get_semester_month_ranges(year)
    
    if month in sem1_months:
        return 1
    elif month in sem2_months:
        return 2
    else:
        # Should not happen with default config
        raise ValueError(f"Month {month} not in any semester for year {year}")


# =============================================================================
# Subject Type Detection
# =============================================================================

def is_individual_subject(name: str | None) -> bool:
    """
    Check if a subject name indicates Individual lessons (per-hour pricing).
    
    Individual subjects contain:
    - "Ind." (case-insensitive)
    - "VSpE_" prefix (case-insensitive)
    
    Args:
        name: Subject name to check.
        
    Returns:
        True if subject is Individual type, False otherwise.
    """
    if not name:
        return False
    return bool(INDIVIDUAL_PATTERN.search(name))


def is_nachhilfe_subject(name: str | None) -> bool:
    """
    Check if a subject name indicates Nachhilfe (tutoring, per-hour pricing).
    
    Nachhilfe subjects contain:
    - "Nachhilfe" (case-insensitive)
    - "NH" as a standalone word (case-insensitive)
    
    Args:
        name: Subject name to check.
        
    Returns:
        True if subject is Nachhilfe type, False otherwise.
    """
    if not name:
        return False
    return bool(NACHHILFE_PATTERN.search(name))


def is_per_hour_subject(name: str | None) -> bool:
    """
    Check if a subject uses per-hour pricing.
    
    This includes both Individual and Nachhilfe subjects.
    
    Args:
        name: Subject name to check.
        
    Returns:
        True if subject uses per-hour pricing, False otherwise.
    """
    return is_individual_subject(name) or is_nachhilfe_subject(name)


# =============================================================================
# Money Rounding Functions
# =============================================================================

def round_money_up(value: Decimal | float | str | None) -> Decimal:
    """
    Round a monetary value UP to 2 decimal places.
    
    Always rounds up (ceiling), as per business requirements.
    Examples:
    - 37.473 -> 37.48
    - 37.471 -> 37.48
    - 37.47  -> 37.47
    - 37.00  -> 37.00
    
    Args:
        value: The value to round. Can be Decimal, float, str, or None.
        
    Returns:
        Rounded Decimal with 2 decimal places.
        Returns 0.00 for None or empty values.
    """
    if value is None or value == '':
        return ZERO
    
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    
    # Round up (ceiling) to 2 decimal places
    # We use quantize with ROUND_CEILING
    return value.quantize(Decimal('0.01'), rounding=ROUND_CEILING)


def normalize_hours(value: Decimal | float | str | None) -> Decimal:
    """
    Normalize hours value to 2 decimal places with standard rounding.
    
    Uses ROUND_HALF_UP (standard rounding) for hours, not ceiling.
    Hours must be >= 0.
    
    Args:
        value: The hours value to normalize.
        
    Returns:
        Normalized Decimal with 2 decimal places.
        Returns 0.00 for None, empty, or negative values.
    """
    if value is None or value == '':
        return ZERO
    
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    
    if value < 0:
        return ZERO
    
    return value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def normalize_money(value: Decimal | float | str | None) -> Decimal:
    """
    Normalize a monetary value to 2 decimal places using ceiling rounding.
    
    Alias for round_money_up for clarity in form handling.
    
    Args:
        value: The value to normalize.
        
    Returns:
        Normalized Decimal with 2 decimal places.
    """
    return round_money_up(value)


# =============================================================================
# Data Classes for Calculation Results
# =============================================================================

@dataclass
class CalculationFlags:
    """Flags and warnings from month value calculation."""
    
    # Months where discount clamped negative value to zero
    clamped_to_zero_months: list[int] = field(default_factory=list)
    
    # Warning if total percent discount exceeded 99%
    percent_discount_exceeded: bool = False
    original_percent_sum: Decimal = field(default_factory=lambda: ZERO)
    
    # Months where Nachhilfe exemption applied (no discounts)
    nachhilfe_exempt_months: list[int] = field(default_factory=list)
    
    # Months where discounts_disabled flag applied
    discounts_disabled_months: list[int] = field(default_factory=list)
    
    # Months zeroed due to contract termination
    terminated_months: list[int] = field(default_factory=list)
    
    # Termination effective month (if contract is terminated)
    termination_from_month: int | None = None
    
    @property
    def has_warnings(self) -> bool:
        """Check if there are any warnings."""
        return (
            bool(self.clamped_to_zero_months) or
            self.percent_discount_exceeded
        )
    
    @property
    def requires_confirmation(self) -> bool:
        """Check if user confirmation is required."""
        return bool(self.clamped_to_zero_months)


# =============================================================================
# Base Amount Calculation
# =============================================================================

def get_subject_name_for_semester(record: "KarteiRecord", semester: int) -> str | None:
    """
    Get the subject name for a semester.
    
    Prefers catalog reference, falls back to legacy field.
    
    Args:
        record: The KarteiRecord instance.
        semester: 1 or 2.
        
    Returns:
        Subject name or None if not set.
    """
    if semester == 1:
        if record.subject1_ref_id:
            return record.subject1_ref.name
        return record.subject1 or None
    else:
        if record.subject2_ref_id:
            return record.subject2_ref.name
        return record.subject2 or None


def get_price_for_semester(record: "KarteiRecord", semester: int) -> Decimal | None:
    """
    Get the price amount for a semester.
    
    Prefers catalog reference, falls back to legacy field.
    
    Args:
        record: The KarteiRecord instance.
        semester: 1 or 2.
        
    Returns:
        Price amount or None if not set.
    """
    if semester == 1:
        if record.price1_ref_id:
            return record.price1_ref.amount
        return record.price1
    else:
        if record.price2_ref_id:
            return record.price2_ref.amount
        return record.price2


def get_start_month_for_semester(record: "KarteiRecord", semester: int) -> int:
    """
    Get the start month for billing in a semester.
    
    Args:
        record: The KarteiRecord instance.
        semester: 1 or 2.
        
    Returns:
        Start month number (1-6 for semester 1, 7-12 for semester 2).
    """
    if semester == 1:
        return record.start_month_1 or 1
    else:
        return record.start_month_2 or 7


def build_base_amounts(
    record: "KarteiRecord",
    *,
    apply_from_month_1: int | None = None,
    apply_from_month_2: int | None = None,
    apply_to_month_1: int | None = None,
    apply_to_month_2: int | None = None,
    hours_amounts: dict[str, Decimal | str] | None = None,
) -> dict[str, Decimal]:
    """
    Build base amounts for all months (before discounts).
    
    For each month:
    - Determine which semester it belongs to
    - Check if month is >= start_month for that semester
    - Calculate base amount based on subject type:
      - Regular: base = price_per_month
      - Individual/Nachhilfe: base = hours * price_per_hour
    
    Args:
        record: The KarteiRecord instance.
        apply_from_month_1: If set, only update months >= this value in semester 1.
            Months before this keep their current base_amounts.
        apply_from_month_2: If set, only update months >= this value in semester 2.
            Months before this keep their current base_amounts.
        apply_to_month_1: If set, only update months <= this value in semester 1.
            Months after this keep their current base_amounts.
        apply_to_month_2: If set, only update months <= this value in semester 2.
            Months after this keep their current base_amounts.
        hours_amounts: Optional hours data to use instead of record.hours_amounts.
            Format: {"month_1": "2.00", "month_2": "1.50", ...}
    
    Returns:
        Dict mapping "month_1".."month_12" to Decimal base amounts.
    """
    result: dict[str, Decimal] = {}
    
    # Get existing base_amounts for partial updates
    existing_bases = getattr(record, 'base_amounts', None) or {}
    
    # Get hours data
    hours_data = hours_amounts if hours_amounts is not None else (
        getattr(record, 'hours_amounts', None) or {}
    )
    
    # Get semester ranges
    sem1_months, sem2_months = get_semester_month_ranges(record.year)
    
    for month_num in range(1, 13):
        field_key = f"month_{month_num}"
        
        # Determine semester
        semester = get_semester_for_month(month_num, record.year)
        
        # Check if we should skip this month (partial update - before apply_from)
        apply_from = apply_from_month_1 if semester == 1 else apply_from_month_2
        if apply_from is not None and month_num < apply_from:
            # Keep existing base
            existing_value = existing_bases.get(field_key)
            if existing_value is not None:
                result[field_key] = Decimal(str(existing_value))
            else:
                result[field_key] = ZERO
            continue
        
        # Check if we should skip this month (partial update - after apply_to)
        apply_to = apply_to_month_1 if semester == 1 else apply_to_month_2
        if apply_to is not None and month_num > apply_to:
            # Keep existing base
            existing_value = existing_bases.get(field_key)
            if existing_value is not None:
                result[field_key] = Decimal(str(existing_value))
            else:
                result[field_key] = ZERO
            continue
        
        # Get semester parameters
        subject_name = get_subject_name_for_semester(record, semester)
        price = get_price_for_semester(record, semester)
        start_month = get_start_month_for_semester(record, semester)
        
        # Month before start_month: 0.00
        if month_num < start_month:
            result[field_key] = ZERO
            continue
        
        # No subject or price: 0.00
        if not subject_name or price is None:
            result[field_key] = ZERO
            continue
        
        # Calculate base based on subject type
        if is_per_hour_subject(subject_name):
            # Per-hour pricing: base = hours * price_per_hour
            hours_value = hours_data.get(field_key, ZERO)
            if hours_value is None or hours_value == '':
                hours_value = ZERO
            hours_decimal = normalize_hours(hours_value)
            base = hours_decimal * Decimal(str(price))
        else:
            # Per-month pricing: base = price
            base = Decimal(str(price))
        
        result[field_key] = round_money_up(base)
    
    return result


# =============================================================================
# Discount Application
# =============================================================================

def collect_discounts_for_month(
    month: int,
    family_discounts: list,
    record_discounts: list,
) -> tuple[Decimal, Decimal]:
    """
    Collect all applicable discounts for a specific month.
    
    Args:
        month: The month number (1-12).
        family_discounts: List of FamilyDiscount instances.
        record_discounts: List of RecordDiscount instances.
        
    Returns:
        Tuple of (total_percent_sum, total_fixed_sum).
        Percent sum is a fraction (e.g., 0.25 = 25%).
    """
    from apps.catalog.models import DiscountKind
    
    percent_sum = ZERO
    fixed_sum = ZERO
    
    all_discounts = list(family_discounts) + list(record_discounts)
    
    for discount_assignment in all_discounts:
        applicable_months = discount_assignment.get_applicable_months()
        
        if month not in applicable_months:
            continue
        
        discount = discount_assignment.discount
        if discount.kind == DiscountKind.PERCENT:
            percent_sum += discount.value
        else:  # FIXED
            fixed_sum += discount.value
    
    return (percent_sum, fixed_sum)


def calculate_month_values(
    record: "KarteiRecord",
    family_discounts: list | None = None,
    record_discounts: list | None = None,
    *,
    base_amounts: dict[str, Decimal] | None = None,
) -> tuple[dict[str, Decimal], CalculationFlags]:
    """
    Calculate final month values after applying discounts.
    
    Discount application rules:
    1. Percent discounts are summed (clamped to 99% max)
    2. Fixed discounts are summed and subtracted after percent
    3. Fixed discounts only apply if base > 0
    4. Nachhilfe subjects: discounts never apply
    5. If record.discounts_disabled: discounts never apply
    6. Negative results are clamped to 0 and flagged
    
    Args:
        record: The KarteiRecord instance.
        family_discounts: List of FamilyDiscount instances for this family/year.
            If None, will be fetched from database.
        record_discounts: List of RecordDiscount instances for this record.
            If None, will be fetched from database.
        base_amounts: Pre-calculated base amounts. If None, will be calculated.
        
    Returns:
        Tuple of (month_values dict, CalculationFlags).
    """
    from apps.catalog.models import FamilyDiscount, RecordDiscount
    
    flags = CalculationFlags()
    result: dict[str, Decimal] = {}
    
    # Get base amounts
    if base_amounts is None:
        base_amounts = build_base_amounts(record)
    
    # Fetch discounts if not provided
    if family_discounts is None:
        family_discounts = list(
            FamilyDiscount.objects.filter(
                year=record.year,
                family_id=record.family_id,
            ).select_related('discount')
        )
    
    if record_discounts is None:
        record_discounts = list(
            record.record_discounts.select_related('discount')
        ) if record.pk else []
    
    # Check if discounts are globally disabled for this record
    discounts_disabled = getattr(record, 'discounts_disabled', False)
    discounts_disabled_months = getattr(record, 'discounts_disabled_months', None) or []
    
    # Check for contract termination
    is_terminated = getattr(record, 'is_contract_terminated', False)
    terminated_from_month = getattr(record, 'contract_terminated_from_month', None)
    if is_terminated and terminated_from_month is not None:
        flags.termination_from_month = terminated_from_month
    
    # Get semester ranges
    sem1_months, sem2_months = get_semester_month_ranges(record.year)
    
    for month_num in range(1, 13):
        field_key = f"month_{month_num}"
        base = base_amounts.get(field_key, ZERO)
        
        # Check contract termination first - months >= terminated_from_month are zero
        if is_terminated and terminated_from_month is not None:
            if month_num >= terminated_from_month:
                result[field_key] = ZERO
                flags.terminated_months.append(month_num)
                continue
        
        # Determine subject for this month
        semester = get_semester_for_month(month_num, record.year)
        subject_name = get_subject_name_for_semester(record, semester)
        
        # Check discount exemptions
        apply_discounts = True
        
        # Check discounts_disabled flag with month-specific logic
        # - discounts_disabled=True + empty months list → disabled for ALL months
        # - discounts_disabled=True + non-empty months list → disabled only for listed months
        if discounts_disabled:
            if not discounts_disabled_months:
                # Empty list = all months disabled
                apply_discounts = False
                flags.discounts_disabled_months.append(month_num)
            elif month_num in discounts_disabled_months:
                # Month is in the disabled list
                apply_discounts = False
                flags.discounts_disabled_months.append(month_num)
        
        if apply_discounts and is_nachhilfe_subject(subject_name):
            apply_discounts = False
            flags.nachhilfe_exempt_months.append(month_num)
        
        if not apply_discounts or base == ZERO:
            # No discounts: final = base
            result[field_key] = round_money_up(base)
            continue
        
        # Collect discounts for this month
        percent_sum, fixed_sum = collect_discounts_for_month(
            month_num, family_discounts, record_discounts
        )
        
        # Clamp percent sum to 99%
        if percent_sum > MAX_PERCENT_DISCOUNT:
            flags.percent_discount_exceeded = True
            flags.original_percent_sum = percent_sum
            percent_sum = MAX_PERCENT_DISCOUNT
        
        # Apply percent discount
        after_percent = base * (Decimal('1') - percent_sum)
        
        # Apply fixed discount (only if base > 0, which we already checked)
        after_fixed = after_percent - fixed_sum
        
        # Clamp negative to zero
        if after_fixed < ZERO:
            flags.clamped_to_zero_months.append(month_num)
            after_fixed = ZERO
        
        result[field_key] = round_money_up(after_fixed)
    
    return (result, flags)


# =============================================================================
# High-level Calculation Functions
# =============================================================================

def get_month_breakdown(
    record: "KarteiRecord",
    month: int,
    family_discounts: list | None = None,
    record_discounts: list | None = None,
) -> dict:
    """
    Get detailed breakdown of month charge calculation.
    
    Returns a dict with all calculation details for UI display.
    For LEGACY or OVERRIDE mode, returns minimal info with available=False.
    
    Args:
        record: The KarteiRecord instance.
        month: Month number (1-12).
        family_discounts: Optional list of FamilyDiscount instances.
        record_discounts: Optional list of RecordDiscount instances.
        
    Returns:
        Dict with breakdown data suitable for JSON serialization.
    """
    from apps.catalog.models import FamilyDiscount, RecordDiscount, DiscountKind
    
    result: dict = {
        'month': month,
        'available': False,
        'reason': None,
        'source_mode': record.months_mode,
    }
    
    # Check modes that cannot provide breakdown
    if record.months_mode == 'LEGACY':
        result['reason'] = 'LEGACY_NO_DATA'
        result['final'] = str(getattr(record, f'month_{month}', None) or '0.00')
        return result
    
    if record.months_mode == 'OVERRIDE':
        result['reason'] = 'OVERRIDE_MODE'
        result['final'] = str(getattr(record, f'month_{month}', None) or '0.00')
        return result
    
    # AUTO mode - full breakdown available
    result['available'] = True
    
    # Determine semester
    semester = get_semester_for_month(month, record.year)
    result['semester'] = semester
    
    # Get subject info
    subject_name = get_subject_name_for_semester(record, semester)
    result['subject_name'] = subject_name or ''
    
    # Determine billing kind
    is_hourly = is_per_hour_subject(subject_name)
    result['billing_kind'] = 'HOURLY' if is_hourly else 'MONTHLY'
    
    # Get base amount
    base_amounts = getattr(record, 'base_amounts', None) or {}
    field_key = f'month_{month}'
    base = Decimal(str(base_amounts.get(field_key, '0.00') or '0.00'))
    result['base'] = str(base)
    
    # For HOURLY, get hours and calculate unit price
    if is_hourly:
        hours_amounts = getattr(record, 'hours_amounts', None) or {}
        hours = Decimal(str(hours_amounts.get(field_key, '0.00') or '0.00'))
        result['hours'] = str(hours)
        
        # Calculate unit price (price per hour used)
        if hours > 0:
            unit_price = base / hours
            result['unit_price'] = str(round_money_up(unit_price))
        else:
            # Get price from reference
            price = get_price_for_semester(record, semester)
            result['unit_price'] = str(price) if price else None
    
    # Discount info
    discounts_disabled = getattr(record, 'discounts_disabled', False)
    discounts_disabled_months = getattr(record, 'discounts_disabled_months', None) or []
    result['discounts_disabled'] = discounts_disabled
    
    # Check if Nachhilfe
    is_nh = is_nachhilfe_subject(subject_name)
    
    # Determine if discounts are skipped for this specific month
    discounts_skipped_for_month = False
    if discounts_disabled:
        if not discounts_disabled_months:
            # Empty list = all months disabled
            discounts_skipped_for_month = True
            result['discounts_skipped_reason'] = 'DISABLED'
        elif month in discounts_disabled_months:
            # This month is in the disabled list
            discounts_skipped_for_month = True
            result['discounts_skipped_reason'] = 'DISABLED_MONTH'
        else:
            result['discounts_skipped_reason'] = None
    elif is_nh:
        discounts_skipped_for_month = True
        result['discounts_skipped_reason'] = 'NACHHILFE'
    else:
        result['discounts_skipped_reason'] = None
    
    # Fetch discounts if not provided
    if family_discounts is None:
        family_discounts = list(
            FamilyDiscount.objects.filter(
                year=record.year,
                family_id=record.family_id,
            ).select_related('discount')
        )
    
    if record_discounts is None:
        record_discounts = list(
            record.record_discounts.select_related('discount')
        ) if record.pk else []
    
    # Collect applicable discounts
    percent_discounts_applied = []
    fixed_discounts_applied = []
    percent_sum = ZERO
    fixed_sum = ZERO
    
    if not discounts_skipped_for_month and base > 0:
        all_discounts = list(family_discounts) + list(record_discounts)
        
        for discount_assignment in all_discounts:
            applicable_months = discount_assignment.get_applicable_months()
            
            if month not in applicable_months:
                continue
            
            discount = discount_assignment.discount
            discount_info = {
                'id': discount.pk,
                'description': discount.description,
                'value': str(discount.value),
                'kind': discount.kind,
                'months_display': discount_assignment.months_display(),
            }
            
            if discount.kind == DiscountKind.PERCENT:
                percent_discounts_applied.append(discount_info)
                percent_sum += discount.value
            else:
                fixed_discounts_applied.append(discount_info)
                fixed_sum += discount.value
    
    result['percent_discounts_applied'] = percent_discounts_applied
    result['fixed_discounts_applied'] = fixed_discounts_applied
    result['percent_sum'] = str(percent_sum)
    result['fixed_sum'] = str(fixed_sum)
    
    # Calculation steps
    percent_clamped = False
    original_percent_sum = percent_sum
    if percent_sum > MAX_PERCENT_DISCOUNT:
        percent_clamped = True
        percent_sum = MAX_PERCENT_DISCOUNT
    result['percent_clamped'] = percent_clamped
    if percent_clamped:
        result['original_percent_sum'] = str(original_percent_sum)
    
    # Skip discounts for Nachhilfe or if disabled for this month
    if discounts_skipped_for_month or base == ZERO:
        result['after_percent'] = str(base)
        result['after_fixed'] = str(base)
        result['final'] = str(round_money_up(base))
        result['clamped_to_zero'] = False
        result['fixed_applied'] = False
        return result
    
    # Apply percent discount
    after_percent = base * (Decimal('1') - percent_sum)
    result['after_percent'] = str(round_money_up(after_percent))
    
    # Apply fixed discount
    result['fixed_applied'] = base > 0
    after_fixed = after_percent - fixed_sum
    
    # Clamp to zero
    clamped_to_zero = after_fixed < ZERO
    if clamped_to_zero:
        after_fixed = ZERO
    result['clamped_to_zero'] = clamped_to_zero
    result['after_fixed'] = str(round_money_up(after_fixed))
    result['final'] = str(round_money_up(after_fixed))
    
    return result


def recalculate_record_months(
    record: "KarteiRecord",
    *,
    apply_from_month_1: int | None = None,
    apply_from_month_2: int | None = None,
    apply_to_month_1: int | None = None,
    apply_to_month_2: int | None = None,
    hours_amounts: dict[str, Decimal | str] | None = None,
    touched_months: set[int] | None = None,
) -> CalculationFlags:
    """
    Recalculate all month values for a record.
    
    This is the main entry point for billing calculation. It:
    1. Builds base amounts
    2. Applies discounts
    3. Updates record.base_amounts and record.month_* fields
    
    Args:
        record: The KarteiRecord instance to update.
        apply_from_month_1: If set, only update months >= this in semester 1.
        apply_from_month_2: If set, only update months >= this in semester 2.
        apply_to_month_1: If set, only update months <= this in semester 1.
        apply_to_month_2: If set, only update months <= this in semester 2.
        hours_amounts: Optional hours data to use.
        touched_months: If set, only update these specific months (for LEGACY->AUTO).
            Other months will keep their current values.
        
    Returns:
        CalculationFlags with any warnings.
    """
    # Build base amounts
    base_amounts = build_base_amounts(
        record,
        apply_from_month_1=apply_from_month_1,
        apply_from_month_2=apply_from_month_2,
        apply_to_month_1=apply_to_month_1,
        apply_to_month_2=apply_to_month_2,
        hours_amounts=hours_amounts,
    )
    
    # Store base amounts on record
    record.base_amounts = {k: str(v) for k, v in base_amounts.items()}
    
    # Calculate final values with discounts
    month_values, flags = calculate_month_values(
        record, base_amounts=base_amounts
    )
    
    # Update record month fields
    for month_num in range(1, 13):
        field_key = f"month_{month_num}"
        
        # If touched_months is specified, only update those months
        if touched_months is not None and month_num not in touched_months:
            # Keep current value for untouched months
            continue
        
        value = month_values.get(field_key, ZERO)
        setattr(record, field_key, value)
    
    return flags


# =============================================================================
# LEGACY to AUTO Conversion Helpers
# =============================================================================

def _parse_apply_from_month(value) -> int | None:
    """
    Safely parse apply_from_month value from form data.
    
    Args:
        value: Can be string ("3", "7", ""), int, or None.
        
    Returns:
        Parsed int or None if empty/invalid.
    """
    if value is None or value == '':
        return None
    try:
        return int(value)
    except (ValueError, TypeError):
        return None


def detect_meaningful_changes(
    original: "KarteiRecord",
    cleaned_data: dict,
    hours_amounts: dict[str, str] | None = None,
) -> tuple[bool, set[int]]:
    """
    Detect if there are "meaningful" changes that should trigger LEGACY->AUTO conversion.
    
    Meaningful changes are those that affect billing calculations:
    - price*_ref or start_month_* changes
    - subject*_ref changes (may change pricing type)
    - discounts_disabled changes
    - contract_type/status changes (is_monthly_contract, is_contract_terminated)
    - hours input for hourly subjects (touched months)
    
    When price*_ref changes, only months from apply_from_month_* are marked as touched,
    preserving legacy values for earlier months.
    
    When start_month_* changes, only months that changed their billing status
    (from 0.00 to charged or vice versa) are marked as touched.
    
    Args:
        original: The original KarteiRecord from database.
        cleaned_data: Form cleaned_data with proposed changes.
        hours_amounts: Hours data from form (optional).
        
    Returns:
        Tuple of (has_meaningful_changes, touched_months_set).
        touched_months_set contains month numbers (1-12) affected by the changes.
    """
    has_changes = False
    touched_months: set[int] = set()
    
    sem1_months = set(range(1, 7))
    sem2_months = set(range(7, 13))
    
    # Parse apply_from_month values from form data
    apply_from_1 = _parse_apply_from_month(cleaned_data.get('apply_from_month_1'))
    apply_from_2 = _parse_apply_from_month(cleaned_data.get('apply_from_month_2'))
    
    # Get new start months for reference
    new_start_1 = cleaned_data.get('start_month_1') or 1
    new_start_2 = cleaned_data.get('start_month_2') or 7
    old_start_1 = original.start_month_1 or 1
    old_start_2 = original.start_month_2 or 7
    
    # Check price1_ref change - only touch months from apply_from_month_1
    # Exception: linking ref to existing legacy value with same amount is NOT meaningful
    new_price1_ref = cleaned_data.get('price1_ref')
    new_price1_ref_id = new_price1_ref.id if new_price1_ref else None
    if new_price1_ref_id != original.price1_ref_id:
        # Check if this is just linking to matching legacy value
        is_linking_to_same_price1 = (
            original.price1_ref_id is None
            and new_price1_ref is not None
            and original.price1 is not None
            and new_price1_ref.amount == original.price1
        )
        if not is_linking_to_same_price1:
            has_changes = True
            # Only touch months from apply_from_month_1 to 6
            start_month = apply_from_1 if apply_from_1 else 1
            touched_months.update(range(start_month, 7))
    
    # Check price2_ref change - only touch months from apply_from_month_2
    # Exception: linking ref to existing legacy value with same amount is NOT meaningful
    new_price2_ref = cleaned_data.get('price2_ref')
    new_price2_ref_id = new_price2_ref.id if new_price2_ref else None
    if new_price2_ref_id != original.price2_ref_id:
        # Check if this is just linking to matching legacy value
        is_linking_to_same_price2 = (
            original.price2_ref_id is None
            and new_price2_ref is not None
            and original.price2 is not None
            and new_price2_ref.amount == original.price2
        )
        if not is_linking_to_same_price2:
            has_changes = True
            # Only touch months from apply_from_month_2 to 12
            start_month = apply_from_2 if apply_from_2 else 7
            touched_months.update(range(start_month, 13))
    
    # Check subject1_ref change (may change pricing type - hourly vs monthly)
    # Touch months from current start_month onwards (subject change affects calculation type)
    # Exception: linking ref to existing legacy value with same name is NOT meaningful
    new_subject1_ref = cleaned_data.get('subject1_ref')
    new_subject1_ref_id = new_subject1_ref.id if new_subject1_ref else None
    if new_subject1_ref_id != original.subject1_ref_id:
        # Check if this is just linking to matching legacy value
        is_linking_to_same_subject1 = (
            original.subject1_ref_id is None
            and new_subject1_ref is not None
            and original.subject1
            and _normalize_subject_name(new_subject1_ref.name) == _normalize_subject_name(original.subject1)
        )
        if not is_linking_to_same_subject1:
            has_changes = True
            # Touch months from start_month_1 to 6
            effective_start = min(old_start_1, new_start_1)
            touched_months.update(range(effective_start, 7))
    
    # Check subject2_ref change
    # Exception: linking ref to existing legacy value with same name is NOT meaningful
    new_subject2_ref = cleaned_data.get('subject2_ref')
    new_subject2_ref_id = new_subject2_ref.id if new_subject2_ref else None
    if new_subject2_ref_id != original.subject2_ref_id:
        # Check if this is just linking to matching legacy value
        is_linking_to_same_subject2 = (
            original.subject2_ref_id is None
            and new_subject2_ref is not None
            and original.subject2
            and _normalize_subject_name(new_subject2_ref.name) == _normalize_subject_name(original.subject2)
        )
        if not is_linking_to_same_subject2:
            has_changes = True
            # Touch months from start_month_2 to 12
            effective_start = min(old_start_2, new_start_2)
            touched_months.update(range(effective_start, 13))
    
    # Check start_month_1 change - only touch months that changed status
    if new_start_1 != old_start_1:
        has_changes = True
        # Only months between old and new start change their billing status
        # e.g., start 1->3: months 1,2 go from charged to 0.00 (or vice versa)
        min_start = min(old_start_1, new_start_1)
        max_start = max(old_start_1, new_start_1)
        # Touch months from min_start to (max_start - 1)
        touched_months.update(range(min_start, max_start))
    
    # Check start_month_2 change - only touch months that changed status
    if new_start_2 != old_start_2:
        has_changes = True
        min_start = min(old_start_2, new_start_2)
        max_start = max(old_start_2, new_start_2)
        touched_months.update(range(min_start, max_start))
    
    # Check discounts_disabled change
    new_discounts_disabled = cleaned_data.get('discounts_disabled', False)
    if new_discounts_disabled != original.discounts_disabled:
        has_changes = True
        touched_months.update(range(1, 13))
    
    # Check contract type change (is_monthly_contract)
    # Derived from contract_type form field
    contract_type = cleaned_data.get('contract_type', 'yearly')
    new_is_monthly = (contract_type == 'monthly')
    if new_is_monthly != original.is_monthly_contract:
        has_changes = True
        touched_months.update(range(1, 13))
    
    # Check contract status change (is_contract_terminated)
    contract_status = cleaned_data.get('contract_status', 'active')
    new_is_terminated = (contract_status == 'terminated')
    if new_is_terminated != original.is_contract_terminated:
        has_changes = True
        touched_months.update(range(1, 13))
    
    # Check hours changes for per-hour subjects
    if hours_amounts:
        original_hours = original.hours_amounts or {}
        for month_num in range(1, 13):
            field_key = f"month_{month_num}"
            new_hours = hours_amounts.get(field_key, '0.00')
            old_hours = original_hours.get(field_key, '0.00')
            
            # Normalize for comparison
            try:
                new_val = Decimal(str(new_hours or '0.00'))
                old_val = Decimal(str(old_hours or '0.00'))
                if new_val != old_val:
                    has_changes = True
                    touched_months.add(month_num)
            except (ValueError, TypeError):
                pass
    
    return has_changes, touched_months


def recalculate_legacy_to_auto(
    record: "KarteiRecord",
    touched_months: set[int],
    hours_amounts: dict[str, str] | None = None,
) -> CalculationFlags:
    """
    Convert a LEGACY record to AUTO mode with partial month updates.
    
    Only touched months are recalculated; untouched months keep their legacy values.
    
    Args:
        record: The KarteiRecord instance (will be modified).
        touched_months: Set of month numbers (1-12) to recalculate.
        hours_amounts: Hours data for per-hour subjects.
        
    Returns:
        CalculationFlags with any warnings.
    """
    # Store original month values for preservation
    original_month_values = {}
    for month_num in range(1, 13):
        field_key = f"month_{month_num}"
        original_month_values[field_key] = getattr(record, field_key)
    
    # Recalculate only touched months
    flags = recalculate_record_months(
        record,
        hours_amounts=hours_amounts,
        touched_months=touched_months,
    )
    
    # Set mode to AUTO
    record.months_mode = 'AUTO'
    
    return flags


# =============================================================================
# Month Mismatch Detection (for highlighting suspicious values)
# =============================================================================

def get_month_mismatches(record: "KarteiRecord") -> set[int]:
    """
    Detect months where the stored value doesn't match expected calculation.
    
    This is used to highlight "suspicious" month values with a red border in UI.
    
    Logic:
    1. OVERRIDE mode: no mismatches (user intentionally set values manually).
    2. For per-hour subjects (Ind., Nachhilfe, VSpE_): skip price comparison
       since the price is per hour and Monatswert can be any valid amount.
    3. For other months:
       - AUTO mode: use stored base_amounts and recalculate with discounts
       - LEGACY mode: rebuild base_amounts from current prices/hours and compare
    4. Compare expected vs actual values (both normalized to 2 decimals).
    
    Args:
        record: The KarteiRecord instance to check.
        
    Returns:
        Set of month numbers (1-12) that have mismatched values.
    """
    # OVERRIDE mode: never highlight (intentional manual values)
    if record.months_mode == 'OVERRIDE':
        return set()
    
    mismatches: set[int] = set()
    
    # Determine expected values based on mode
    if record.months_mode == 'AUTO':
        # For AUTO: use stored base_amounts (preserves historical prices per month)
        stored_bases = getattr(record, 'base_amounts', None) or {}
        base_amounts = {}
        for month_num in range(1, 13):
            field_key = f"month_{month_num}"
            raw_val = stored_bases.get(field_key, '0.00')
            try:
                base_amounts[field_key] = Decimal(str(raw_val or '0.00'))
            except (ValueError, TypeError):
                base_amounts[field_key] = ZERO
    else:
        # For LEGACY: rebuild base_amounts from current data as if it were AUTO
        hours_amounts = getattr(record, 'hours_amounts', None) or {}
        base_amounts = build_base_amounts(record, hours_amounts=hours_amounts)
    
    # Calculate expected month values
    expected_values, _flags = calculate_month_values(record, base_amounts=base_amounts)
    
    # Compare expected vs actual for each month
    for month_num in range(1, 13):
        field_key = f"month_{month_num}"
        
        # Determine semester and subject for this month
        semester = get_semester_for_month(month_num, record.year)
        subject_name = get_subject_name_for_semester(record, semester)
        
        # Skip per-hour subjects (Ind., Nachhilfe, VSpE_)
        # Their monthly value can be anything based on hours worked
        if is_per_hour_subject(subject_name):
            continue
        
        # Get expected value
        expected = expected_values.get(field_key, ZERO)
        if not isinstance(expected, Decimal):
            expected = Decimal(str(expected or '0.00'))
        expected = expected.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        
        # Get actual stored value
        actual = getattr(record, field_key, None)
        if actual is None:
            actual = ZERO
        else:
            try:
                actual = Decimal(str(actual))
            except (ValueError, TypeError):
                actual = ZERO
        actual = actual.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        
        # Compare
        if expected != actual:
            mismatches.add(month_num)
    
    return mismatches
