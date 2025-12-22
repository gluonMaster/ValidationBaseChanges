"""
Import services for legacy_import.

This module contains the core import logic for migrating data from Access
databases to Django models. It handles:

- Mapping Access fields (Value1..Value52) to Django model fields
- Detecting and skipping marker rows ("Zahlung" family separators)
- Creating/updating KarteiRecord, PendingChange, DeclinedChange
- Collecting import statistics

See PROMPT_07_LEGACY_IMPORT.md for detailed requirements.
See DOMAIN_MODEL.md Section 8 (Legacy Import Mapping) for field mapping.
"""

from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, time
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Any

from django.db import transaction
from django.utils import timezone

if TYPE_CHECKING:
    from apps.legacy_import.access_client import RowDict


logger = logging.getLogger(__name__)


# =============================================================================
# Constants: Field Mapping
# =============================================================================

# Mapping from Access Value fields to Django model fields
# Based on DOMAIN_MODEL.md Section 8.2 (Access Value Fields → Django Fields)
ACCESS_TO_DJANGO_FIELD_MAP: dict[str, str] = {
    "Value1": "family_id",       # A - FamilyID
    "Value2": "parent_name",     # B - Parent
    # Value3 - visual/extra (skipped)
    "Value4": "child_name",      # D - Child
    "Value5": "birthdate",       # E - Birthdate
    "Value6": "address",         # F - Address
    "Value7": "phone",           # G - Phone
    "Value8": "mobile",          # H - Mobile
    "Value9": "email",           # I - Email
    "Value10": "subject1",       # J - Subject1
    "Value11": "teacher1_legacy_name",  # K - Teacher 1st semester (legacy text)
    # Value12 - skipped (intermediate)
    "Value13": "price1",         # M - Price1
    "Value14": "contract_type_raw",  # N - Contract type marker (raw text, may contain 'O/V')
    "Value15": "subject2",       # O - Subject2
    "Value16": "teacher2_legacy_name",  # P - Teacher 2nd semester (legacy text)
    # Value17 - skipped
    "Value18": "price2",         # R - Price2
    # Value19 - skipped
    "Value20": "contract_status_raw",  # T - Contract status marker (raw text, may contain 'KN')
    # Months: Value21-Value32 -> month_1..month_12
    "Value21": "month_1",
    "Value22": "month_2",
    "Value23": "month_3",
    "Value24": "month_4",
    "Value25": "month_5",
    "Value26": "month_6",
    "Value27": "month_7",
    "Value28": "month_8",
    "Value29": "month_9",
    "Value30": "month_10",
    "Value31": "month_11",
    "Value32": "month_12",
    # Value33-Value36 - skipped
    "Value37": "extra1",         # AK - Extra1
    "Value38": "extra2",         # AL - Extra2
    "Value39": "extra3",         # AM - Extra3
    # Value40-Value46 - skipped
    "Value47": "sepa_marker",    # AU - SepaMarker
    # Value48 is ID field in Access, handled separately
    "Value49": "last_change_role",  # AW - LastChangeRole
    "Value50": "last_change_date",  # AX - LastChangeDate
    "Value51": "last_change_time",  # AY - LastChangeTime
    "Value52": "history_raw",       # AZ - History
}

# Fields that are only set during patch import (--patch-fields mode)
PATCH_ONLY_FIELDS: tuple[str, ...] = (
    "teacher1_legacy_name",
    "teacher2_legacy_name",
    "contract_type_raw",
    "is_monthly_contract",
    "contract_status_raw",
    "is_contract_terminated",
    "sepa_marker",
)

# Green color value for marker rows (Excel interior color for family separator)
# This is the typical RGB value for green marker rows
MARKER_GREEN_COLOR = 5287936  # RGB(0, 176, 80) as Long

# Alternative marker colors (may vary by Excel version)
MARKER_GREEN_COLORS = {
    5287936,   # Standard green
    5296274,   # Alternative green
    32768,     # Dark green
    65280,     # Bright green
}


# =============================================================================
# Contract Type / Status Detection (Pure Functions)
# =============================================================================

def detect_is_monthly_contract(contract_type_raw: str) -> bool:
    """
    Detect if the contract is a monthly contract (Monatsvertrag).
    
    The raw string may contain 'O/V' anywhere in the text, possibly adjacent
    to numbers or other characters. Case-insensitive search.
    
    Examples:
        - "O/V" -> True
        - "12O/V" -> True
        - "o/v45" -> True
        - "Something O/V else" -> True
        - "OV" (no slash) -> False
        - "" -> False
    
    Args:
        contract_type_raw: Raw text from Access Value14.
        
    Returns:
        True if 'O/V' substring is found (case-insensitive).
    """
    if not contract_type_raw:
        return False
    return "o/v" in contract_type_raw.lower()


def detect_is_contract_terminated(contract_status_raw: str) -> bool:
    """
    Detect if the contract is terminated (gekündigt).
    
    Looks for 'KN' as a separate token (word boundary). The token must be
    separated by whitespace or be at the start/end of the string.
    
    Examples:
        - "KN" -> True
        - "KN something" -> True
        - "text KN" -> True
        - "a KN b" -> True
        - "UNKNOWN" -> False (KN is part of a larger word)
        - "AKN" -> False
        - "" -> False
    
    Args:
        contract_status_raw: Raw text from Access Value20.
        
    Returns:
        True if 'KN' token is found (word boundary, case-insensitive).
    """
    if not contract_status_raw:
        return False
    # Regex: word boundary + KN + word boundary, case-insensitive
    return bool(re.search(r"(?:^|\s)KN(?:\s|$)", contract_status_raw, re.IGNORECASE))


# =============================================================================
# Data Structures
# =============================================================================

@dataclass
class ImportStats:
    """
    Statistics collected during an import operation.
    
    This dataclass tracks counts of various outcomes during import,
    allowing for detailed reporting and analysis.
    """
    # Row counts
    total_rows: int = 0
    
    # Record operations
    created_records: int = 0
    updated_records: int = 0
    skipped_records: int = 0
    
    # Pending/Declined
    pending_created: int = 0
    pending_updated: int = 0
    declined_created: int = 0
    
    # Special cases
    marker_skipped: int = 0
    
    # Issues detected
    familyid_issues: int = 0
    conflicts_year_id: int = 0
    parse_errors: int = 0
    
    # Details for reporting
    marker_row_ids: list[int] = field(default_factory=list)
    conflict_ids: list[tuple[int, int]] = field(default_factory=list)  # (year, id)
    error_details: list[dict[str, Any]] = field(default_factory=list)
    
    def merge(self, other: "ImportStats") -> "ImportStats":
        """Merge another ImportStats into this one."""
        self.total_rows += other.total_rows
        self.created_records += other.created_records
        self.updated_records += other.updated_records
        self.skipped_records += other.skipped_records
        self.pending_created += other.pending_created
        self.pending_updated += other.pending_updated
        self.declined_created += other.declined_created
        self.marker_skipped += other.marker_skipped
        self.familyid_issues += other.familyid_issues
        self.conflicts_year_id += other.conflicts_year_id
        self.parse_errors += other.parse_errors
        self.marker_row_ids.extend(other.marker_row_ids)
        self.conflict_ids.extend(other.conflict_ids)
        self.error_details.extend(other.error_details)
        return self
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "total_rows": self.total_rows,
            "created_records": self.created_records,
            "updated_records": self.updated_records,
            "skipped_records": self.skipped_records,
            "pending_created": self.pending_created,
            "pending_updated": self.pending_updated,
            "declined_created": self.declined_created,
            "marker_skipped": self.marker_skipped,
            "familyid_issues": self.familyid_issues,
            "conflicts_year_id": self.conflicts_year_id,
            "parse_errors": self.parse_errors,
            "marker_row_ids_sample": self.marker_row_ids[:10],  # First 10 for brevity
            "conflict_ids_sample": self.conflict_ids[:10],
            "error_count": len(self.error_details),
        }


@dataclass
class FamilyIdIssue:
    """
    Represents a detected FamilyID inconsistency.
    """
    issue_type: str  # SAME_FAMILY_DIFFERENT_FAMILYIDS or SAME_FAMILYID_DIFFERENT_FAMILIES
    year: int
    record_id: int
    family_id: str
    parent: str
    email: str
    family_key: str
    details: str


# =============================================================================
# Helper Functions: Value Conversion
# =============================================================================

def _parse_decimal(value: Any) -> Decimal | None:
    """
    Parse a value as Decimal.
    
    Args:
        value: Raw value from Access (string, int, float, or None).
        
    Returns:
        Decimal value or None if empty/invalid.
    """
    if value is None or value == "":
        return None
    
    if isinstance(value, Decimal):
        return value
    
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    
    # String parsing
    try:
        # Handle German number format (comma as decimal separator)
        str_value = str(value).strip()
        str_value = str_value.replace(",", ".")
        
        # Remove currency symbols and spaces
        str_value = re.sub(r"[€$\s]", "", str_value)
        
        if not str_value:
            return None
        
        return Decimal(str_value)
    except (InvalidOperation, ValueError):
        return None


def _parse_date(value: Any) -> date | None:
    """
    Parse a value as date.
    
    Args:
        value: Raw value from Access.
        
    Returns:
        date object or None if empty/invalid.
    """
    if value is None or value == "":
        return None
    
    if isinstance(value, date):
        return value
    
    # Try to parse string date
    str_value = str(value).strip()
    
    # German format: DD.MM.YYYY
    match = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", str_value)
    if match:
        try:
            day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
            return date(year, month, day)
        except ValueError:
            pass
    
    # ISO format: YYYY-MM-DD
    match = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", str_value)
    if match:
        try:
            year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
            return date(year, month, day)
        except ValueError:
            pass
    
    return None


def _parse_time(value: Any) -> time | None:
    """
    Parse a value as time.
    
    Args:
        value: Raw value from Access.
        
    Returns:
        time object or None if empty/invalid.
    """
    if value is None or value == "":
        return None
    
    if isinstance(value, time):
        return value
    
    str_value = str(value).strip()
    
    # HH:MM:SS or HH:MM
    match = re.match(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", str_value)
    if match:
        try:
            hour = int(match.group(1))
            minute = int(match.group(2))
            second = int(match.group(3)) if match.group(3) else 0
            return time(hour, minute, second)
        except ValueError:
            pass
    
    return None


def _clean_string(value: Any) -> str:
    """
    Clean a string value.
    
    Args:
        value: Raw value from Access.
        
    Returns:
        Cleaned string (trimmed, normalized whitespace).
    """
    if value is None:
        return ""
    
    str_value = str(value)
    
    # Normalize Unicode
    str_value = unicodedata.normalize("NFC", str_value)
    
    # Strip and normalize whitespace
    str_value = " ".join(str_value.split())
    
    return str_value


# =============================================================================
# Marker Row Detection
# =============================================================================

def is_marker_row(row: "RowDict") -> bool:
    """
    Detect if a row is a marker row (family separator).
    
    Marker rows are visual grouping rows in Excel that:
    - Have FamilyID in A (Value1)
    - Have Parent in B (Value2)
    - Have " Zahlung" (with leading space) in D (Value4)
    - Have empty/null values in most other fields
    - Are typically colored green
    
    These rows should NOT be imported into Django.
    
    Args:
        row: RowDict from Access.
        
    Returns:
        True if this is a marker row to be skipped.
    """
    child_value = row.get("Value4")
    
    if child_value is None:
        return False
    
    # Check for "Zahlung" marker (with or without leading space)
    child_str = str(child_value).strip()
    
    if child_str.lower() == "zahlung":
        return True
    
    # Also check with leading space preserved
    if str(child_value).lstrip().lower() == "zahlung":
        return True
    
    # Additional check: if Value4 is exactly " Zahlung" (with space)
    if str(child_value) == " Zahlung":
        return True
    
    # Check for green color (if interior color data is available)
    interior_color_1 = row.get("InteriorColor1")
    if interior_color_1 and interior_color_1 in MARKER_GREEN_COLORS:
        # If green colored and Value4 looks like "Zahlung", it's a marker
        if "zahlung" in str(child_value).lower():
            return True
    
    return False


# =============================================================================
# Field Value Extraction
# =============================================================================

def extract_record_id(row: "RowDict") -> int | None:
    """
    Extract the record ID from a row.
    
    The ID is in the dedicated "ID" field, not Value48.
    
    Args:
        row: RowDict from Access.
        
    Returns:
        Integer ID or None if not present.
    """
    id_value = row.get("ID")
    
    if id_value is None:
        return None
    
    try:
        return int(id_value)
    except (ValueError, TypeError):
        return None


def extract_field_value(row: "RowDict", access_field: str, django_field: str) -> Any:
    """
    Extract and convert a field value from Access to Django format.
    
    Args:
        row: RowDict from Access.
        access_field: Access field name (e.g., "Value1").
        django_field: Django model field name (e.g., "family_id").
        
    Returns:
        Converted value appropriate for the Django field.
    """
    raw_value = row.get(access_field)
    
    # Determine conversion based on Django field type
    decimal_fields = {
        "price1", "price2",
        "month_1", "month_2", "month_3", "month_4",
        "month_5", "month_6", "month_7", "month_8",
        "month_9", "month_10", "month_11", "month_12",
    }
    
    date_fields = {"birthdate", "last_change_date"}
    time_fields = {"last_change_time"}
    
    if django_field in decimal_fields:
        return _parse_decimal(raw_value)
    elif django_field in date_fields:
        return _parse_date(raw_value)
    elif django_field in time_fields:
        return _parse_time(raw_value)
    else:
        return _clean_string(raw_value)


def row_to_model_data(row: "RowDict", year: int) -> dict[str, Any]:
    """
    Convert an Access row to Django model field values.
    
    Args:
        row: RowDict from Access.
        year: Year for the record.
        
    Returns:
        Dictionary of Django model field names to values.
    """
    data: dict[str, Any] = {"year": year}
    
    # Extract ID
    record_id = extract_record_id(row)
    if record_id is not None:
        data["id"] = record_id
    
    # Map all known fields
    for access_field, django_field in ACCESS_TO_DJANGO_FIELD_MAP.items():
        value = extract_field_value(row, access_field, django_field)
        data[django_field] = value
    
    # Compute derived boolean flags from raw contract fields
    contract_type_raw = data.get("contract_type_raw", "")
    contract_status_raw = data.get("contract_status_raw", "")
    data["is_monthly_contract"] = detect_is_monthly_contract(contract_type_raw)
    data["is_contract_terminated"] = detect_is_contract_terminated(contract_status_raw)
    
    return data


# =============================================================================
# FamilyID Analysis
# =============================================================================

def normalize_for_family_key(value: str | None) -> str:
    """
    Normalize a string for use in family key matching.
    
    Args:
        value: String to normalize.
        
    Returns:
        Normalized lowercase string.
    """
    if not value:
        return ""
    
    normalized = str(value).strip().lower()
    
    # Remove invisible/control characters
    normalized = "".join(c for c in normalized if unicodedata.category(c)[0] != "C")
    
    # Normalize Unicode
    normalized = unicodedata.normalize("NFKC", normalized)
    
    # Collapse multiple spaces
    normalized = " ".join(normalized.split())
    
    return normalized


def compute_family_key(parent_name: str | None, email: str | None) -> str:
    """
    Compute a family key for matching families across different FamilyIDs.
    
    Args:
        parent_name: Parent name (Value2).
        email: Email address (Value9).
        
    Returns:
        Family key string.
    """
    parent_normalized = normalize_for_family_key(parent_name)
    email_normalized = normalize_for_family_key(email)
    
    return f"{parent_normalized}|{email_normalized}"


def analyze_familyid_issues(
    rows: list["RowDict"],
    existing_records: dict[tuple[int, int], dict[str, Any]],
    year: int,
) -> tuple[list[FamilyIdIssue], dict[str, str]]:
    """
    Analyze FamilyID consistency issues.
    
    Detects:
    - Same family (by family_key) with different FamilyIDs
    - Same FamilyID used for different families
    
    Args:
        rows: List of RowDict from Access.
        existing_records: Dict of (year, id) -> record data for existing records.
        year: Import year.
        
    Returns:
        Tuple of (list of issues, mapping for auto-merge if applicable).
    """
    issues: list[FamilyIdIssue] = []
    
    # Build mappings
    family_key_to_family_ids: dict[str, set[str]] = {}
    family_id_to_family_keys: dict[str, set[str]] = {}
    
    # Track row details for issue reporting
    family_key_details: dict[str, list[tuple[int, str, str, str]]] = {}  # key -> [(id, fid, parent, email)]
    
    # Process new rows
    for row in rows:
        if is_marker_row(row):
            continue
        
        record_id = extract_record_id(row)
        if record_id is None:
            continue
        
        family_id = _clean_string(row.get("Value1"))
        parent_name = _clean_string(row.get("Value2"))
        email = _clean_string(row.get("Value9"))
        family_key = compute_family_key(parent_name, email)
        
        if not family_key.strip("|"):  # Skip if both parent and email are empty
            continue
        
        # Update mappings
        if family_key not in family_key_to_family_ids:
            family_key_to_family_ids[family_key] = set()
            family_key_details[family_key] = []
        family_key_to_family_ids[family_key].add(family_id)
        family_key_details[family_key].append((record_id, family_id, parent_name, email))
        
        if family_id not in family_id_to_family_keys:
            family_id_to_family_keys[family_id] = set()
        family_id_to_family_keys[family_id].add(family_key)
    
    # Also include existing records in the analysis
    for (rec_year, rec_id), rec_data in existing_records.items():
        family_id = rec_data.get("family_id", "")
        parent_name = rec_data.get("parent_name", "")
        email = rec_data.get("email", "")
        family_key = compute_family_key(parent_name, email)
        
        if not family_key.strip("|"):
            continue
        
        if family_key not in family_key_to_family_ids:
            family_key_to_family_ids[family_key] = set()
            family_key_details[family_key] = []
        family_key_to_family_ids[family_key].add(family_id)
        
        if family_id not in family_id_to_family_keys:
            family_id_to_family_keys[family_id] = set()
        family_id_to_family_keys[family_id].add(family_key)
    
    # Detect issues
    
    # Issue type 1: Same family, different FamilyIDs
    for family_key, family_ids in family_key_to_family_ids.items():
        if len(family_ids) > 1:
            details = family_key_details.get(family_key, [])
            for record_id, fid, parent, email in details:
                issues.append(FamilyIdIssue(
                    issue_type="SAME_FAMILY_DIFFERENT_FAMILYIDS",
                    year=year,
                    record_id=record_id,
                    family_id=fid,
                    parent=parent,
                    email=email,
                    family_key=family_key,
                    details=f"Family has multiple FamilyIDs: {', '.join(sorted(family_ids))}",
                ))
    
    # Issue type 2: Same FamilyID, different families
    for family_id, family_keys in family_id_to_family_keys.items():
        if len(family_keys) > 1:
            # Get sample records for this FamilyID
            for fk in family_keys:
                details = family_key_details.get(fk, [])
                for record_id, fid, parent, email in details:
                    if fid == family_id:
                        issues.append(FamilyIdIssue(
                            issue_type="SAME_FAMILYID_DIFFERENT_FAMILIES",
                            year=year,
                            record_id=record_id,
                            family_id=family_id,
                            parent=parent,
                            email=email,
                            family_key=fk,
                            details=f"FamilyID used by {len(family_keys)} different families",
                        ))
    
    # Auto-merge mapping (for SAME_FAMILY_DIFFERENT_FAMILYIDS)
    # Choose canonical FamilyID: first alphabetically or most common
    auto_merge_map: dict[str, str] = {}
    
    for family_key, family_ids in family_key_to_family_ids.items():
        if len(family_ids) > 1:
            # Choose canonical: first non-empty, alphabetically sorted
            sorted_ids = sorted(fid for fid in family_ids if fid)
            if sorted_ids:
                canonical = sorted_ids[0]
                for fid in family_ids:
                    if fid and fid != canonical:
                        auto_merge_map[fid] = canonical
    
    return issues, auto_merge_map


# =============================================================================
# Import Functions
# =============================================================================

def import_tbl_kartei(
    rows: list["RowDict"],
    year: int,
    dry_run: bool = False,
    update_existing: bool = True,
    familyid_mapping: dict[str, str] | None = None,
) -> ImportStats:
    """
    Import rows from tblKartei into KarteiRecord.
    
    Args:
        rows: List of RowDict from Access tblKartei.
        year: Year for the records.
        dry_run: If True, don't write to database, just collect stats.
        update_existing: If True, update existing records; if False, skip them.
        familyid_mapping: Optional mapping to normalize FamilyIDs.
        
    Returns:
        ImportStats with counts of operations performed.
    """
    from apps.karteien.models import KarteiRecord, RecordStatus
    
    stats = ImportStats()
    
    # Track seen record IDs to detect duplicates within this import batch
    seen_ids: set[int] = set()
    
    for row in rows:
        stats.total_rows += 1
        
        # Skip marker rows
        if is_marker_row(row):
            record_id = extract_record_id(row)
            stats.marker_skipped += 1
            if record_id is not None:
                stats.marker_row_ids.append(record_id)
            logger.debug("Skipping marker row: ID=%s", record_id)
            continue
        
        # Extract record ID
        record_id = extract_record_id(row)
        if record_id is None:
            stats.skipped_records += 1
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "missing_id",
                "row_sample": {k: v for k, v in list(row.items())[:5]},
            })
            logger.warning("Row without ID, skipping")
            continue
        
        # Check for duplicate (year, id) within this import batch
        if record_id in seen_ids:
            stats.conflicts_year_id += 1
            stats.conflict_ids.append((year, record_id))
            stats.error_details.append({
                "type": "duplicate_year_id",
                "year": year,
                "id": record_id,
                "row_sample": {
                    "family_id": row.get("Value1"),
                    "parent": row.get("Value2"),
                    "child": row.get("Value4"),
                },
            })
            logger.warning(
                "Duplicate (year=%d, id=%d) in import batch, skipping row",
                year, record_id
            )
            stats.skipped_records += 1
            continue
        seen_ids.add(record_id)
        
        # Convert row to model data
        model_data = row_to_model_data(row, year)
        
        # Apply FamilyID mapping if provided
        if familyid_mapping and model_data.get("family_id"):
            original_fid = model_data["family_id"]
            if original_fid in familyid_mapping:
                model_data["family_id"] = familyid_mapping[original_fid]
                logger.debug(
                    "Remapped FamilyID for ID=%d: %s -> %s",
                    record_id, original_fid, model_data["family_id"]
                )
        
        if dry_run:
            # Check if record exists
            exists = KarteiRecord.objects.filter(year=year, id=record_id).exists()
            if exists:
                stats.updated_records += 1
            else:
                stats.created_records += 1
            continue
        
        # Create or update record
        try:
            with transaction.atomic():
                existing = KarteiRecord.objects.filter(year=year, id=record_id).first()
                
                if existing:
                    if not update_existing:
                        stats.skipped_records += 1
                        continue
                    
                    # Update existing record
                    for field_name, value in model_data.items():
                        if field_name not in ("id", "year"):  # Don't overwrite key fields
                            setattr(existing, field_name, value)
                    existing.save()
                    stats.updated_records += 1
                    logger.debug("Updated KarteiRecord: year=%d, id=%d", year, record_id)
                else:
                    # Create new record
                    # Set status to NORMAL for main table import
                    model_data["status"] = RecordStatus.NORMAL
                    record = KarteiRecord(**model_data)
                    record.save()
                    stats.created_records += 1
                    logger.debug("Created KarteiRecord: year=%d, id=%d", year, record_id)
                    
        except Exception as e:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "save_error",
                "year": year,
                "id": record_id,
                "error": str(e),
            })
            logger.exception("Error saving KarteiRecord year=%d, id=%d", year, record_id)
    
    return stats


def import_pre_tbl_kartei(
    rows: list["RowDict"],
    year: int,
    dry_run: bool = False,
    familyid_mapping: dict[str, str] | None = None,
) -> ImportStats:
    """
    Import rows from pre_tblKartei into PendingChange.
    
    This creates/updates KarteiRecord with status=PENDING and creates
    corresponding PendingChange entries.
    
    Args:
        rows: List of RowDict from Access pre_tblKartei.
        year: Year for the records.
        dry_run: If True, don't write to database.
        familyid_mapping: Optional mapping to normalize FamilyIDs.
        
    Returns:
        ImportStats with counts of operations performed.
    """
    from apps.approvals.models import PendingChange
    from apps.karteien.models import KarteiRecord, RecordStatus, TRACKED_FIELDS
    
    stats = ImportStats()
    
    for row in rows:
        stats.total_rows += 1
        
        # Skip marker rows
        if is_marker_row(row):
            record_id = extract_record_id(row)
            stats.marker_skipped += 1
            if record_id is not None:
                stats.marker_row_ids.append(record_id)
            continue
        
        # Extract record ID
        record_id = extract_record_id(row)
        if record_id is None:
            stats.skipped_records += 1
            stats.parse_errors += 1
            continue
        
        # Convert row to model data
        model_data = row_to_model_data(row, year)
        
        # Apply FamilyID mapping if provided
        if familyid_mapping and model_data.get("family_id"):
            original_fid = model_data["family_id"]
            if original_fid in familyid_mapping:
                model_data["family_id"] = familyid_mapping[original_fid]
        
        if dry_run:
            # Check if pending change exists
            exists = KarteiRecord.objects.filter(year=year, id=record_id).exists()
            if exists:
                pending_exists = PendingChange.objects.filter(record__year=year, record__id=record_id).exists()
                if pending_exists:
                    stats.pending_updated += 1
                else:
                    stats.pending_created += 1
            else:
                stats.created_records += 1
                stats.pending_created += 1
            continue
        
        try:
            with transaction.atomic():
                # Find or create KarteiRecord
                record, created = KarteiRecord.objects.get_or_create(
                    year=year,
                    id=record_id,
                    defaults={
                        **{k: v for k, v in model_data.items() if k not in ("id", "year")},
                        "status": RecordStatus.PENDING,
                    }
                )
                
                if created:
                    stats.created_records += 1
                else:
                    # Update status to PENDING
                    record.status = RecordStatus.PENDING
                    record.save(update_fields=["status"])
                
                # Build snapshot from tracked fields
                snapshot: dict[str, Any] = {}
                for field_name in TRACKED_FIELDS:
                    if field_name in model_data:
                        value = model_data[field_name]
                        # Convert to JSON-serializable format
                        if isinstance(value, Decimal):
                            snapshot[field_name] = str(value)
                        elif isinstance(value, (date, time)):
                            snapshot[field_name] = value.isoformat()
                        else:
                            snapshot[field_name] = value
                
                # Create or update PendingChange
                pending, pending_created = PendingChange.objects.update_or_create(
                    record=record,
                    defaults={
                        "snapshot": snapshot,
                        "is_processed": False,
                    }
                )
                
                if pending_created:
                    stats.pending_created += 1
                else:
                    stats.pending_updated += 1
                
                logger.debug("Created/updated PendingChange for year=%d, id=%d", year, record_id)
                
        except Exception as e:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "pending_save_error",
                "year": year,
                "id": record_id,
                "error": str(e),
            })
            logger.exception("Error saving PendingChange year=%d, id=%d", year, record_id)
    
    return stats


def import_decl_tbl_kartei(
    rows: list["RowDict"],
    year: int,
    dry_run: bool = False,
    familyid_mapping: dict[str, str] | None = None,
) -> ImportStats:
    """
    Import rows from decl_tblKartei into DeclinedChange.
    
    This creates/updates KarteiRecord with status=DECLINED and creates
    corresponding DeclinedChange entries.
    
    Args:
        rows: List of RowDict from Access decl_tblKartei.
        year: Year for the records.
        dry_run: If True, don't write to database.
        familyid_mapping: Optional mapping to normalize FamilyIDs.
        
    Returns:
        ImportStats with counts of operations performed.
    """
    from apps.approvals.models import DeclinedChange
    from apps.karteien.models import KarteiRecord, RecordStatus, TRACKED_FIELDS
    
    stats = ImportStats()
    
    for row in rows:
        stats.total_rows += 1
        
        # Skip marker rows
        if is_marker_row(row):
            record_id = extract_record_id(row)
            stats.marker_skipped += 1
            if record_id is not None:
                stats.marker_row_ids.append(record_id)
            continue
        
        # Extract record ID
        record_id = extract_record_id(row)
        if record_id is None:
            stats.skipped_records += 1
            stats.parse_errors += 1
            continue
        
        # Convert row to model data
        model_data = row_to_model_data(row, year)
        
        # Apply FamilyID mapping if provided
        if familyid_mapping and model_data.get("family_id"):
            original_fid = model_data["family_id"]
            if original_fid in familyid_mapping:
                model_data["family_id"] = familyid_mapping[original_fid]
        
        if dry_run:
            stats.declined_created += 1
            continue
        
        try:
            with transaction.atomic():
                # Find or create KarteiRecord
                record, created = KarteiRecord.objects.get_or_create(
                    year=year,
                    id=record_id,
                    defaults={
                        **{k: v for k, v in model_data.items() if k not in ("id", "year")},
                        "status": RecordStatus.DECLINED,
                    }
                )
                
                if created:
                    stats.created_records += 1
                else:
                    # Update status to DECLINED
                    record.status = RecordStatus.DECLINED
                    record.save(update_fields=["status"])
                
                # Build snapshot from tracked fields
                snapshot: dict[str, Any] = {}
                for field_name in TRACKED_FIELDS:
                    if field_name in model_data:
                        value = model_data[field_name]
                        if isinstance(value, Decimal):
                            snapshot[field_name] = str(value)
                        elif isinstance(value, (date, time)):
                            snapshot[field_name] = value.isoformat()
                        else:
                            snapshot[field_name] = value
                
                # Create DeclinedChange (always create new, as there can be multiple declines)
                DeclinedChange.objects.create(
                    record=record,
                    snapshot=snapshot,
                    decline_reason="Imported from legacy Access database",
                )
                
                stats.declined_created += 1
                logger.debug("Created DeclinedChange for year=%d, id=%d", year, record_id)
                
        except Exception as e:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "declined_save_error",
                "year": year,
                "id": record_id,
                "error": str(e),
            })
            logger.exception("Error saving DeclinedChange year=%d, id=%d", year, record_id)
    
    return stats


def sync_history_for_records(
    year: int,
    record_ids: list[int] | None = None,
    dry_run: bool = False,
) -> int:
    """
    Synchronize history from history_raw for imported records.
    
    Args:
        year: Year of records to sync.
        record_ids: Optional list of specific record IDs. If None, sync all for year.
        dry_run: If True, don't create history events.
        
    Returns:
        Number of history events created.
    """
    from apps.history.services import sync_history_from_raw
    from apps.karteien.models import KarteiRecord
    
    queryset = KarteiRecord.objects.filter(year=year)
    if record_ids:
        queryset = queryset.filter(id__in=record_ids)
    
    # Only sync records that have history_raw
    queryset = queryset.exclude(history_raw="").exclude(history_raw__isnull=True)
    
    total_events = 0
    
    for record in queryset.iterator():
        if dry_run:
            # Just count potential events
            from apps.history.services import parse_raw_history
            events = parse_raw_history(record.history_raw)
            total_events += len(events)
        else:
            created_events = sync_history_from_raw(record)
            total_events += len(created_events)
    
    return total_events


# =============================================================================
# Patch Import (Update only specific fields without full re-import)
# =============================================================================

@dataclass
class PatchStats:
    """
    Statistics collected during a patch import operation.
    """
    total_rows: int = 0
    records_found: int = 0
    records_updated: int = 0
    records_not_found: int = 0
    marker_skipped: int = 0
    parse_errors: int = 0
    error_details: list[dict[str, Any]] = field(default_factory=list)
    not_found_ids: list[int] = field(default_factory=list)
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "total_rows": self.total_rows,
            "records_found": self.records_found,
            "records_updated": self.records_updated,
            "records_not_found": self.records_not_found,
            "marker_skipped": self.marker_skipped,
            "parse_errors": self.parse_errors,
            "not_found_ids_sample": self.not_found_ids[:20],  # First 20 for brevity
            "error_count": len(self.error_details),
        }


def patch_tbl_kartei(
    rows: list["RowDict"],
    year: int,
    dry_run: bool = False,
) -> PatchStats:
    """
    Patch existing KarteiRecord entries with additional fields from Access.
    
    This function updates ONLY the patch-specific fields (teacher names,
    contract type/status, sepa_marker) without modifying other fields.
    It is used for adding missing data to already-imported records.
    
    Lookup is by domain key (year, id), NOT by pkid.
    
    Updated fields:
        - teacher1_legacy_name (from Value11)
        - teacher2_legacy_name (from Value16)
        - contract_type_raw (from Value14)
        - is_monthly_contract (computed from contract_type_raw)
        - contract_status_raw (from Value20)
        - is_contract_terminated (computed from contract_status_raw)
        - sepa_marker (from Value47)
    
    Args:
        rows: List of RowDict from Access tblKartei.
        year: Year for the records.
        dry_run: If True, don't write to database, just collect stats.
        
    Returns:
        PatchStats with counts of operations performed.
    """
    from apps.karteien.models import KarteiRecord
    
    stats = PatchStats()
    
    for row in rows:
        stats.total_rows += 1
        
        # Skip marker rows
        if is_marker_row(row):
            stats.marker_skipped += 1
            logger.debug("Patch: Skipping marker row")
            continue
        
        # Extract record ID
        record_id = extract_record_id(row)
        if record_id is None:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "missing_id",
                "row_sample": {k: v for k, v in list(row.items())[:5]},
            })
            logger.warning("Patch: Row without ID, skipping")
            continue
        
        # Find existing record by (year, id)
        try:
            existing = KarteiRecord.objects.filter(year=year, id=record_id).first()
        except Exception as e:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "db_lookup_error",
                "year": year,
                "id": record_id,
                "error": str(e),
            })
            logger.exception("Patch: DB lookup error for year=%d, id=%d", year, record_id)
            continue
        
        if existing is None:
            stats.records_not_found += 1
            stats.not_found_ids.append(record_id)
            logger.debug("Patch: Record not found year=%d, id=%d", year, record_id)
            continue
        
        stats.records_found += 1
        
        # Extract only patch-specific fields
        teacher1 = _clean_string(row.get("Value11"))
        teacher2 = _clean_string(row.get("Value16"))
        contract_type_raw = _clean_string(row.get("Value14"))
        contract_status_raw = _clean_string(row.get("Value20"))
        sepa_marker = _clean_string(row.get("Value47"))
        
        # Compute derived flags
        is_monthly = detect_is_monthly_contract(contract_type_raw)
        is_terminated = detect_is_contract_terminated(contract_status_raw)
        
        if dry_run:
            stats.records_updated += 1
            continue
        
        # Update only patch fields
        try:
            with transaction.atomic():
                existing.teacher1_legacy_name = teacher1
                existing.teacher2_legacy_name = teacher2
                existing.contract_type_raw = contract_type_raw
                existing.is_monthly_contract = is_monthly
                existing.contract_status_raw = contract_status_raw
                existing.is_contract_terminated = is_terminated
                existing.sepa_marker = sepa_marker
                
                existing.save(update_fields=[
                    "teacher1_legacy_name",
                    "teacher2_legacy_name",
                    "contract_type_raw",
                    "is_monthly_contract",
                    "contract_status_raw",
                    "is_contract_terminated",
                    "sepa_marker",
                ])
                
                stats.records_updated += 1
                logger.debug("Patch: Updated record year=%d, id=%d", year, record_id)
                
        except Exception as e:
            stats.parse_errors += 1
            stats.error_details.append({
                "type": "patch_save_error",
                "year": year,
                "id": record_id,
                "error": str(e),
            })
            logger.exception("Patch: Error saving record year=%d, id=%d", year, record_id)
    
    return stats
