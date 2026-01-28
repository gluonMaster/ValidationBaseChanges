"""
Services for the history app.

This module contains:
- History parsing functions (parse_raw_history)
- History synchronization (sync_history_from_raw)
- History data structures (HistoryEventData)

Implements parsing of the VBA history format from AZ/Value52:
- New format: [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||
- Legacy format: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY ||
- Decline format: DCL(<N>-><comment>)||

See Export_HistoryBuilder.bas and Export_HistoryParser.bas for VBA implementation.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING

from django.utils import timezone

if TYPE_CHECKING:
    from apps.karteien.models import KarteiRecord


# =============================================================================
# Constants: History Format Tags and Delimiters
# =============================================================================

# Session separator (between history entries)
HD_SESSION = "||"

# Field separator within session
HD_FIELD = ";"

# Old->New value separator
HD_VALUE = "->"

# Comment delimiters
HD_COMMENT_START = "/@"
HD_COMMENT_END = "@/"

# Retroactive change prefix
HD_RUCK_PREFIX = "RUCK:"

# Decline tag
HD_DECLINE_TAG = "DCL"

# Field tags mapping (VBA tag -> Django field name)
TAG_TO_FIELD: dict[str, str] = {
    "FID": "family_id",
    "PAR": "parent_name",
    "CHD": "child_name",
    "DOB": "birthdate",
    "BDT": "birthdate",  # Alternative tag
    "ADR": "address",
    "TEL": "phone",
    "PHN": "phone",  # Alternative tag
    "MOB": "mobile",
    "EML": "email",
    "SB1": "subject1",
    "PR1": "price1",
    "SB2": "subject2",
    "PR2": "price2",
    "M01": "month_1",
    "M02": "month_2",
    "M03": "month_3",
    "M04": "month_4",
    "M05": "month_5",
    "M06": "month_6",
    "M07": "month_7",
    "M08": "month_8",
    "M09": "month_9",
    "M10": "month_10",
    "M11": "month_11",
    "M12": "month_12",
    "EX1": "extra1",
    "EX2": "extra2",
    "EX3": "extra3",
}

# Legacy format month pattern: "Mnt.N" or "Monat N"
LEGACY_MONTH_PATTERN = re.compile(r"(?:Mnt\.?|Monat\s*)(\d{1,2})", re.IGNORECASE)

# Date pattern: DD.MM.YYYY
DATE_PATTERN = re.compile(r"(\d{1,2})\.(\d{1,2})\.(\d{4})")


# =============================================================================
# Data Structures
# =============================================================================

@dataclass
class HistoryEventData:
    """
    Parsed history event data structure.
    
    This is a plain data class (not a Django model) used as intermediate
    representation when parsing raw history strings.
    """
    
    event_time: datetime | None = None
    event_type: str = "CHANGE"  # CHANGE, CREATE, APPROVE, DECLINE, IMPORT, RUCK
    changes: dict[str, dict[str, str | None]] = field(default_factory=dict)
    comment: str = ""
    raw_fragment: str = ""
    is_ruck: bool = False
    decline_number: int | None = None
    decline_comment: str = ""


# =============================================================================
# Format Detection
# =============================================================================

def is_new_format(segment: str) -> bool:
    """
    Detect if a history segment uses the new format.
    
    New format indicators:
    - Contains -> (value separator)
    - Contains /@ or @/ (comment delimiters)
    - Starts with DCL( or RUCK: followed by tag
    - Starts with APR: or ADM: (approve/admin comment)
    - Contains field tags like M01(, ADR(, SB1( etc.
    
    Args:
        segment: A single history session segment (between || separators)
        
    Returns:
        True if new format, False if legacy format
    """
    # Check for new format markers
    if HD_VALUE in segment:
        return True
    
    if HD_COMMENT_START in segment or HD_COMMENT_END in segment:
        return True
    
    # Check for DCL( at start
    if segment.startswith("DCL("):
        return True
    
    # Check for APR: or ADM: prefixes (approve/admin comment)
    if segment.startswith("APR:") or segment.startswith("ADM:"):
        return True
    
    # Check for RUCK: prefix with new format after it
    if segment.startswith(HD_RUCK_PREFIX):
        after_ruck = segment[len(HD_RUCK_PREFIX):].strip()
        if HD_VALUE in after_ruck:
            return True
    
    # Check for new format field tags (M01(, ADR(, etc.)
    new_tag_pattern = re.compile(
        r"(M\d{2}|FID|PAR|CHD|DOB|BDT|ADR|TEL|PHN|MOB|EML|SB1|PR1|SB2|PR2|EX[123])\(",
        re.IGNORECASE,
    )
    if new_tag_pattern.search(segment):
        return True
    
    return False


# =============================================================================
# Date Parsing
# =============================================================================

def parse_date(date_str: str) -> datetime | None:
    """
    Parse a date string in DD.MM.YYYY format.
    
    Args:
        date_str: Date string to parse
        
    Returns:
        datetime object or None if parsing fails
    """
    match = DATE_PATTERN.search(date_str)
    if match:
        try:
            day = int(match.group(1))
            month = int(match.group(2))
            year = int(match.group(3))
            return datetime(year, month, day, tzinfo=timezone.get_current_timezone())
        except (ValueError, TypeError):
            pass
    return None


# =============================================================================
# New Format Parser
# =============================================================================

def parse_new_format_segment(segment: str) -> HistoryEventData:
    """
    Parse a history segment in new format.
    
    New format: [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>
                DCL(<N>-><comment>)/@<user>@/<date>
                APR:<username>/[TAG(old->new);...]/@<comment>@/<date>
                ADM:<username>/[TAG(old->new);...]/@<comment>@/<date>
    
    Args:
        segment: A single history session segment
        
    Returns:
        HistoryEventData with parsed values
    """
    event = HistoryEventData(raw_fragment=segment)
    
    working = segment.strip()
    
    # Check for RUCK: prefix
    if working.startswith(HD_RUCK_PREFIX):
        event.is_ruck = True
        event.event_type = "RUCK"
        working = working[len(HD_RUCK_PREFIX):].strip()
    
    # Check for APR: approve entry
    if working.startswith("APR:"):
        event.event_type = "APPROVE"
        # Remove "APR:<username>/" prefix to get the rest
        # Format: APR:username/[TAG(old->new);...]/@comment@/date||
        prefix_match = re.match(r"APR:([^/]+)/", working)
        if prefix_match:
            working = working[prefix_match.end():]
        # Continue to parse comment, date, and field changes below
    
    # Check for ADM: admin comment entry
    elif working.startswith("ADM:"):
        event.event_type = "CHANGE"  # Keep as CHANGE but include field diffs
        # Remove "ADM:<username>/" prefix
        prefix_match = re.match(r"ADM:([^/]+)/", working)
        if prefix_match:
            working = working[prefix_match.end():]
        # Continue to parse comment, date, and field changes below
    
    # Check for DCL( decline entry
    elif working.startswith("DCL("):
        event.event_type = "DECLINE"
        
        # Parse DCL(N->comment)
        dcl_pattern = re.compile(r"DCL\((\d+)->([^)]*)\)", re.IGNORECASE)
        dcl_match = dcl_pattern.search(working)
        
        if dcl_match:
            event.decline_number = int(dcl_match.group(1))
            event.decline_comment = dcl_match.group(2)
            event.comment = event.decline_comment
        
        # Parse date from the tail (after DCL(...) there may be /@user@/date)
        event.event_time = parse_date(working)
        
        return event
    
    # Extract date from the end (DD.MM.YYYY)
    event.event_time = parse_date(working)
    
    # Extract comment (/@...@/)
    comment_pattern = re.compile(r"/@(.*)@/", re.DOTALL)
    comment_match = comment_pattern.search(working)
    if comment_match:
        event.comment = comment_match.group(1).strip()
        # Remove comment from working string for field parsing
        working = comment_pattern.sub("", working)
    
    # Remove date from end if present
    working = DATE_PATTERN.sub("", working).strip()
    
    # Parse field changes: TAG(OLD->NEW);TAG(OLD->NEW);...
    field_pattern = re.compile(
        r"([A-Z][A-Z0-9]{1,2})\(([^)]*?)->([^)]*?)\)",
        re.IGNORECASE,
    )
    
    for match in field_pattern.finditer(working):
        tag = match.group(1).upper()
        old_value = match.group(2).strip()
        new_value = match.group(3).strip()
        
        # Convert tag to Django field name
        field_name = TAG_TO_FIELD.get(tag)
        if field_name:
            event.changes[field_name] = {"old": old_value, "new": new_value}
    
    return event


# =============================================================================
# Legacy Format Parser
# =============================================================================

def parse_legacy_format_segment(segment: str) -> HistoryEventData:
    """
    Parse a history segment in legacy format.
    
    Legacy formats:
    - Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    - Address: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    - Ruck: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    - Decl_N: Was(); Is(comment).
    
    Args:
        segment: A single history session segment
        
    Returns:
        HistoryEventData with parsed values
    """
    event = HistoryEventData(raw_fragment=segment)
    
    working = segment.strip()
    
    # Check for Ruck: prefix
    if working.lower().startswith("ruck:"):
        event.is_ruck = True
        event.event_type = "RUCK"
        working = working[5:].strip()
    
    # Check for Decl_ prefix
    decl_pattern = re.compile(r"^Decl_(\d+):", re.IGNORECASE)
    decl_match = decl_pattern.match(working)
    if decl_match:
        event.event_type = "DECLINE"
        event.decline_number = int(decl_match.group(1))
        working = decl_pattern.sub("", working).strip()
    
    # Extract date from the end
    event.event_time = parse_date(working)
    
    # Extract legacy comment (/.../)
    legacy_comment_pattern = re.compile(r"/([^/]+)/")
    comment_match = legacy_comment_pattern.search(working)
    if comment_match:
        event.comment = comment_match.group(1).strip()
    
    # Parse War/Ist or Was/Is pattern
    war_ist_pattern = re.compile(
        r"(?:War|Was)\s*\(([^)]*)\)\s*;\s*(?:Ist|Is)\s*\(([^)]*)\)",
        re.IGNORECASE,
    )
    
    for match in war_ist_pattern.finditer(working):
        old_value = match.group(1).strip()
        new_value = match.group(2).strip()
        
        # Try to determine field from context before the War/Ist
        context_before = working[:match.start()]
        
        # Check for month pattern
        month_match = LEGACY_MONTH_PATTERN.search(context_before)
        if month_match:
            month_num = int(month_match.group(1))
            if 1 <= month_num <= 12:
                field_name = f"month_{month_num}"
                event.changes[field_name] = {"old": old_value, "new": new_value}
                continue
        
        # Check for other field indicators
        field_indicators = {
            "address": "address",
            "adresse": "address",
            "subject1": "subject1",
            "fach1": "subject1",
            "subject2": "subject2",
            "fach2": "subject2",
            "parent": "parent_name",
            "eltern": "parent_name",
            "child": "child_name",
            "kind": "child_name",
            "phone": "phone",
            "telefon": "phone",
            "mobile": "mobile",
            "handy": "mobile",
            "email": "email",
        }
        
        context_lower = context_before.lower()
        for indicator, field_name in field_indicators.items():
            if indicator in context_lower:
                event.changes[field_name] = {"old": old_value, "new": new_value}
                break
    
    return event


# =============================================================================
# Main Parsing Functions
# =============================================================================

def parse_raw_history(raw: str) -> list[HistoryEventData]:
    """
    Parse a raw history string (AZ/Value52) into a list of structured events.
    
    Supports both new format (TAG->based) and legacy format (War/Ist based).
    Sessions are separated by "||".
    
    Args:
        raw: Raw history string from KarteiRecord.history_raw
        
    Returns:
        List of HistoryEventData objects, one per session
        
    Example:
        >>> raw = "M01(100->200);M02(50->75)/@Monthly update@/01.12.2025||"
        >>> events = parse_raw_history(raw)
        >>> len(events)
        1
        >>> events[0].changes
        {'month_1': {'old': '100', 'new': '200'}, 'month_2': {'old': '50', 'new': '75'}}
    """
    if not raw or not raw.strip():
        return []
    
    result: list[HistoryEventData] = []
    
    # Split by session separator
    segments = raw.split(HD_SESSION)
    
    for segment in segments:
        segment = segment.strip()
        
        if not segment:
            continue
        
        # Detect format and parse accordingly
        if is_new_format(segment):
            event = parse_new_format_segment(segment)
        else:
            event = parse_legacy_format_segment(segment)
        
        # Only add if we got meaningful data:
        # - Has changes, OR
        # - Has comment/decline_comment (comment-only events), OR
        # - Is a significant event type (DECLINE, CREATE, IMPORT, APPROVE, RUCK)
        has_meaningful_content = (
            event.changes
            or event.comment
            or event.decline_comment
            or event.event_type in ("DECLINE", "CREATE", "IMPORT", "APPROVE", "RUCK")
        )
        if has_meaningful_content:
            result.append(event)
    
    return result


def sync_history_from_raw(record: "KarteiRecord") -> list:
    """
    Synchronize history events from a KarteiRecord's raw history.
    
    Reads the record's history_raw field (AZ/Value52), parses it,
    and creates HistoryEvent objects for any events not already in the database.
    
    This provides one-way compatibility: importing old history from AZ
    into normalized events. It does not modify the existing history_raw string.
    
    Args:
        record: KarteiRecord instance to sync history for
        
    Returns:
        List of newly created HistoryEvent objects
        
    Note:
        - Existing events are not updated, only missing events are created.
        - Events are matched by raw_history_fragment to avoid duplicates.
        - If user information is not available, user field is left null.
    """
    # Import here to avoid circular imports
    from apps.history.models import EventType, HistoryEvent
    
    # Get raw history from record
    raw_history = getattr(record, "history_raw", "") or ""
    
    if not raw_history.strip():
        return []
    
    # Parse raw history
    parsed_events = parse_raw_history(raw_history)
    
    if not parsed_events:
        return []
    
    # Get existing raw fragments to avoid duplicates
    existing_fragments = set(
        HistoryEvent.objects.filter(record=record)
        .exclude(raw_history_fragment="")
        .values_list("raw_history_fragment", flat=True)
    )
    
    created_events: list[HistoryEvent] = []
    
    for parsed in parsed_events:
        # Skip if this fragment already exists
        if parsed.raw_fragment and parsed.raw_fragment in existing_fragments:
            continue
        
        # Map event type
        event_type_map = {
            "CHANGE": EventType.CHANGE,
            "CREATE": EventType.CREATE,
            "APPROVE": EventType.APPROVE,
            "DECLINE": EventType.DECLINE,
            "IMPORT": EventType.IMPORT,
            "RUCK": EventType.RUCK,
        }
        event_type = event_type_map.get(parsed.event_type, EventType.CHANGE)
        
        # Use parsed time or current time
        event_time = parsed.event_time or timezone.now()
        
        # Build comment
        comment = parsed.comment
        if parsed.decline_comment and not comment:
            comment = parsed.decline_comment
        
        # Create HistoryEvent
        history_event = HistoryEvent.objects.create(
            record=record,
            user=None,  # No user info in legacy data
            event_time=event_time,
            event_type=event_type,
            changes=parsed.changes,
            comment=comment,
            raw_history_fragment=parsed.raw_fragment,
        )
        
        created_events.append(history_event)
    
    return created_events


# =============================================================================
# Utility Functions
# =============================================================================

def build_changes_dict(
    old_values: dict[str, str | None],
    new_values: dict[str, str | None],
) -> dict[str, dict[str, str | None]]:
    """
    Build a changes dictionary from old and new value dictionaries.
    
    Only includes fields where values differ.
    
    Args:
        old_values: Dictionary of field_name -> old_value
        new_values: Dictionary of field_name -> new_value
        
    Returns:
        Dictionary of changed fields: {field_name: {"old": old, "new": new}}
    """
    changes: dict[str, dict[str, str | None]] = {}
    
    all_fields = set(old_values.keys()) | set(new_values.keys())
    
    for field_name in all_fields:
        old_val = old_values.get(field_name)
        new_val = new_values.get(field_name)
        
        # Normalize None and empty string
        old_normalized = old_val if old_val else ""
        new_normalized = new_val if new_val else ""
        
        if old_normalized != new_normalized:
            changes[field_name] = {"old": old_val, "new": new_val}
    
    return changes


def get_field_display_name(field_name: str) -> str:
    """
    Get a human-readable display name for a field.
    
    Args:
        field_name: Django model field name
        
    Returns:
        Human-readable display name
    """
    display_names = {
        "family_id": "Family ID",
        "parent_name": "Parent",
        "child_name": "Child",
        "birthdate": "Birthdate",
        "address": "Address",
        "phone": "Phone",
        "mobile": "Mobile",
        "email": "Email",
        "subject1": "Subject 1",
        "price1": "Price 1",
        "subject2": "Subject 2",
        "price2": "Price 2",
        "extra1": "Extra 1",
        "extra2": "Extra 2",
        "extra3": "Extra 3",
    }
    
    # Handle month fields
    if field_name.startswith("month_"):
        month_num = field_name.replace("month_", "")
        return f"Month {month_num}"
    
    return display_names.get(field_name, field_name.replace("_", " ").title())
