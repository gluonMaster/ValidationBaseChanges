"""
FamilyID generation and reservation services.

This module provides centralized functions for working with FamilyIDs:
- Generating the next available FamilyID (considering both existing records and reservations)
- Listing active reservations

FamilyID format: "1. <number>" (prefix.number), global across all years.
"""

from __future__ import annotations

import re
from typing import List

from apps.karteien.models import KarteiRecord

from .models import FamilyIdReservation


# Pattern for FamilyID format: "1. <number>"
FAMILY_ID_PATTERN = re.compile(r'^1\.\s*(\d+)$')

# Maximum valid number in the "correct" 4-digit range (1-9999)
# Values with 5+ digits (e.g., "1. 31431") are treated as erroneous
MAX_VALID_FAMILY_NUMBER = 9999


def get_next_family_id(prefix: str = "1") -> str:
    """
    Generate the next globally unique FamilyID.
    
    Scans all existing family_id values across all years that match
    the pattern "1. <number>" AND all active reservations (is_used=False),
    then returns the next available number.
    
    Only considers "valid" FamilyIDs with numbers <= MAX_VALID_FAMILY_NUMBER.
    Values with 5+ digits are excluded from the maximum calculation.
    
    Args:
        prefix: The prefix for the FamilyID (default "1").
    
    Returns:
        str: Next FamilyID in format "{prefix}. <number>".
    """
    pattern = re.compile(rf'^{re.escape(prefix)}\.\s*(\d+)$')
    
    # Collect all used numbers from KarteiRecord
    all_family_ids = KarteiRecord.objects.values_list('family_id', flat=True).distinct()
    used_numbers = set()
    max_number = 0
    
    for fid in all_family_ids:
        if not fid:
            continue
        match = pattern.match(str(fid).strip())
        if match:
            number = int(match.group(1))
            # Only consider "valid" FamilyIDs (4 digits or less)
            if number <= MAX_VALID_FAMILY_NUMBER:
                used_numbers.add(number)
                if number > max_number:
                    max_number = number
    
    # Collect all reserved numbers (active reservations only)
    reserved_family_ids = list_reserved_family_ids()
    for fid in reserved_family_ids:
        match = pattern.match(str(fid).strip())
        if match:
            number = int(match.group(1))
            if number <= MAX_VALID_FAMILY_NUMBER:
                used_numbers.add(number)
                if number > max_number:
                    max_number = number
    
    # Find the next available number (max + 1)
    # This ensures we always return a number greater than all used/reserved
    next_number = max_number + 1
    
    # Double-check it's not in use (shouldn't happen with max+1, but safety first)
    while next_number in used_numbers:
        next_number += 1
    
    return f"{prefix}. {next_number}"


def list_reserved_family_ids() -> List[str]:
    """
    Get list of all active (not yet used) reserved FamilyIDs.
    
    Returns:
        List of family_id strings from active reservations.
    """
    return list(
        FamilyIdReservation.objects
        .filter(is_used=False)
        .values_list('family_id', flat=True)
    )


def is_family_id_available(family_id: str) -> bool:
    """
    Check if a family_id is available (not used in records, not reserved).
    
    Args:
        family_id: The FamilyID to check.
    
    Returns:
        True if the family_id is available, False otherwise.
    """
    # Check if used in any KarteiRecord
    if KarteiRecord.objects.filter(family_id=family_id).exists():
        return False
    
    # Check if reserved (active reservation)
    if FamilyIdReservation.objects.filter(family_id=family_id, is_used=False).exists():
        return False
    
    return True


def get_family_id_status(family_id: str) -> str:
    """
    Get the status of a family_id.
    
    Args:
        family_id: The FamilyID to check.
    
    Returns:
        One of: 'available', 'used', 'reserved'
    """
    if KarteiRecord.objects.filter(family_id=family_id).exists():
        return 'used'
    
    if FamilyIdReservation.objects.filter(family_id=family_id, is_used=False).exists():
        return 'reserved'
    
    return 'available'


def get_active_reservations() -> List[FamilyIdReservation]:
    """
    Get all active (unused) reservations ordered by family_id.
    
    Returns:
        List of FamilyIdReservation objects that have not been used.
    """
    return list(
        FamilyIdReservation.objects
        .filter(is_used=False)
        .order_by('family_id')
    )


def mark_reservation_as_used(family_id: str) -> bool:
    """
    Mark a reservation as used.
    
    Args:
        family_id: The FamilyID to mark as used.
    
    Returns:
        True if the reservation was found and marked, False otherwise.
    
    Raises:
        ValueError: If the reservation is already used or does not exist.
    """
    from django.utils import timezone
    
    try:
        reservation = FamilyIdReservation.objects.get(
            family_id=family_id,
            is_used=False
        )
        reservation.is_used = True
        reservation.used_at = timezone.now()
        reservation.save(update_fields=['is_used', 'used_at'])
        return True
    except FamilyIdReservation.DoesNotExist:
        raise ValueError(
            f"Reservierung für FamilyID '{family_id}' nicht gefunden oder bereits verwendet."
        )
