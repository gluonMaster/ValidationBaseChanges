"""
Models for the karteien app.

This module contains:
- KarteiRecord — основная модель записи картотеки семей/детей
- Поля: FamilyID, Parent, Child, Birthdate, Address, Phone, Mobile, Email
- Предметы и цены: Subject1, Price1, Subject2, Price2, Extra1-3
- Месячные данные: months 1-12
- Служебные поля: year, status, history

Mapping:
- Excel sheet: Kartei (KindElternDaten_XX_Admin.xlsm)
- Access table: tblKartei (KindElternDaten_XX_front.accdb)

See DOMAIN_MODEL.md for full field mapping.
"""

from __future__ import annotations

from decimal import Decimal
from typing import ClassVar

from django.db import models


# =============================================================================
# Constants: Month Field Names
# =============================================================================

# Month field names (1-indexed to match Excel columns U-AF / indices 21-32)
MONTH_FIELD_NAMES: tuple[str, ...] = tuple(f"month_{i}" for i in range(1, 13))

# Mapping from month number (1-12) to field name
MONTH_NUM_TO_FIELD: dict[int, str] = {i: f"month_{i}" for i in range(1, 13)}

# Mapping from Excel column index (21-32) to month number (1-12)
EXCEL_COL_TO_MONTH: dict[int, int] = {20 + i: i for i in range(1, 13)}

# Tracked fields for history and risk classification
# Based on DOMAIN_MODEL.md Section 4
TRACKED_FIELDS: tuple[str, ...] = (
    "family_id",      # A (1)
    "parent_name",    # B (2)
    "child_name",     # D (4)
    "birthdate",      # E (5)
    "address",        # F (6)
    "phone",          # G (7)
    "mobile",         # H (8)
    "email",          # I (9)
    "subject1",       # J (10)
    "price1",         # M (13)
    "subject2",       # O (15)
    "price2",         # R (18)
    # Months U-AF (21-32)
    *MONTH_FIELD_NAMES,
    "extra1",         # AK (37)
    "extra2",         # AL (38)
    "extra3",         # AM (39)
)

# History field tags for building/parsing history strings
# Used by Export_HistoryBuilder logic
HISTORY_FIELD_TAGS: dict[str, str] = {
    "family_id": "FID",
    "parent_name": "PAR",
    "child_name": "CHD",
    "birthdate": "BDT",
    "address": "ADR",
    "phone": "PHN",
    "mobile": "MOB",
    "email": "EML",
    "subject1": "SB1",
    "price1": "PR1",
    "subject2": "SB2",
    "price2": "PR2",
    "month_1": "M01",
    "month_2": "M02",
    "month_3": "M03",
    "month_4": "M04",
    "month_5": "M05",
    "month_6": "M06",
    "month_7": "M07",
    "month_8": "M08",
    "month_9": "M09",
    "month_10": "M10",
    "month_11": "M11",
    "month_12": "M12",
    "extra1": "EX1",
    "extra2": "EX2",
    "extra3": "EX3",
}


# =============================================================================
# Choices / Enums
# =============================================================================

class RecordStatus(models.TextChoices):
    """
    Status of a KarteiRecord row.
    
    Based on Kartei!BA column behavior in Excel:
    - NORMAL: standard record, only in tblKartei
    - PENDING: record exists in pre_tblKartei, awaiting Superadmin decision
    - DECLINED: record exists in decl_tblKartei, rejected by Superadmin
    """
    NORMAL = "", "Normal"
    PENDING = "PENDING", "Pending Approval"
    DECLINED = "DECLINED", "Declined"


class UserRole(models.TextChoices):
    """
    Role of user who made last change.
    Stored in LastChangeRole / Value49.
    """
    ADMIN = "Admin", "Admin"
    OPERATOR = "Operator", "Operator"
    SUPERADMIN = "Superadmin", "Superadmin"


# =============================================================================
# KarteiRecord Model
# =============================================================================

class KarteiRecord(models.Model):
    """
    Main registry record for a family/child entry.
    
    Maps to:
    - Excel sheet: Kartei (columns A-AZ)
    - Access table: tblKartei (fields ID, Value1-Value52)
    
    The unique identifier is (year, id) where:
    - id corresponds to Kartei!AV and tblKartei.ID
    - year separates records from different yearly databases
    """
    
    # -------------------------------------------------------------------------
    # Primary Key & Year
    # -------------------------------------------------------------------------
    
    id = models.PositiveIntegerField(
        primary_key=True,
        help_text="Unique record ID. Excel: AV (48), Access: ID field.",
    )
    
    year = models.PositiveSmallIntegerField(
        db_index=True,
        help_text="Year of the record (2024, 2025, etc.). Replaces separate yearly Access files.",
    )
    
    # -------------------------------------------------------------------------
    # Basic Family/Child Information (Section 1.2 in DOMAIN_MODEL.md)
    # -------------------------------------------------------------------------
    
    family_id = models.CharField(
        max_length=50,
        db_index=True,
        help_text="Family identifier. Excel: A (1), Access: Value1. Groups related records.",
    )
    
    parent_name = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Parent name (Eltern). Excel: B (2), Access: Value2.",
    )
    
    child_name = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Child name (Kind). Excel: D (4), Access: Value4.",
    )
    
    birthdate = models.DateField(
        null=True,
        blank=True,
        help_text="Child's birthdate. Excel: E (5), Access: Value5.",
    )
    
    address = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text="Address. Excel: F (6), Access: Value6.",
    )
    
    phone = models.CharField(
        max_length=50,
        blank=True,
        default="",
        help_text="Phone number (as text to preserve leading zeros). Excel: G (7), Access: Value7.",
    )
    
    mobile = models.CharField(
        max_length=50,
        blank=True,
        default="",
        help_text="Mobile number (as text). Excel: H (8), Access: Value8.",
    )
    
    email = models.EmailField(
        max_length=255,
        blank=True,
        default="",
        help_text="Email address. Excel: I (9), Access: Value9.",
    )
    
    # -------------------------------------------------------------------------
    # Subjects & Prices (Section 1.3 in DOMAIN_MODEL.md)
    # -------------------------------------------------------------------------
    
    subject1 = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Primary subject 1. Excel: J (10), Access: Value10.",
    )
    
    price1 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Price for subject 1. Excel: M (13), Access: Value13.",
    )
    
    subject2 = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Subject 2. Excel: O (15), Access: Value15.",
    )
    
    price2 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Price for subject 2. Excel: R (18), Access: Value18.",
    )
    
    extra1 = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Extra subject 1. Excel: AK (37), Access: Value37.",
    )
    
    extra2 = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Extra subject 2. Excel: AL (38), Access: Value38.",
    )
    
    extra3 = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Extra subject 3. Excel: AM (39), Access: Value39.",
    )
    
    # -------------------------------------------------------------------------
    # Monthly Fields (Section 1.4 in DOMAIN_MODEL.md)
    # Columns U-AF (indices 21-32), 12 months
    # -------------------------------------------------------------------------
    
    month_1 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 1 value (January). Excel: U (21), Access: Value21.",
    )
    
    month_2 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 2 value (February). Excel: V (22), Access: Value22.",
    )
    
    month_3 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 3 value (March). Excel: W (23), Access: Value23.",
    )
    
    month_4 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 4 value (April). Excel: X (24), Access: Value24.",
    )
    
    month_5 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 5 value (May). Excel: Y (25), Access: Value25.",
    )
    
    month_6 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 6 value (June). Excel: Z (26), Access: Value26.",
    )
    
    month_7 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 7 value (July). Excel: AA (27), Access: Value27.",
    )
    
    month_8 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 8 value (August). Excel: AB (28), Access: Value28.",
    )
    
    month_9 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 9 value (September). Excel: AC (29), Access: Value29.",
    )
    
    month_10 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 10 value (October). Excel: AD (30), Access: Value30.",
    )
    
    month_11 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 11 value (November). Excel: AE (31), Access: Value31.",
    )
    
    month_12 = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Month 12 value (December). Excel: AF (32), Access: Value32.",
    )
    
    # -------------------------------------------------------------------------
    # Technical / Service Fields (Section 1.5 in DOMAIN_MODEL.md)
    # -------------------------------------------------------------------------
    
    sepa_marker = models.CharField(
        max_length=20,
        blank=True,
        default="",
        help_text="SEPA marker (e.g., 'SEPA'). Excel: AU (47), Access: Value47. "
                  "Rows with SEPA have restrictions for Operator role.",
    )
    
    status = models.CharField(
        max_length=20,
        choices=RecordStatus.choices,
        default=RecordStatus.NORMAL,
        db_index=True,
        help_text="Record status: '', 'PENDING', or 'DECLINED'. Excel: BA (53). "
                  "Derived from presence in pre_tblKartei/decl_tblKartei.",
    )
    
    last_change_role = models.CharField(
        max_length=20,
        choices=UserRole.choices,
        blank=True,
        default="",
        help_text="Role of user who made last change. Excel: AW (49), Access: Value49.",
    )
    
    last_change_date = models.DateField(
        null=True,
        blank=True,
        help_text="Date of last sync/change. Excel: AX (50), Access: Value50.",
    )
    
    last_change_time = models.TimeField(
        null=True,
        blank=True,
        help_text="Time of last sync/change. Excel: AY (51), Access: Value51.",
    )
    
    history_raw = models.TextField(
        blank=True,
        default="",
        help_text="Raw history string (legacy format from Export_HistoryBuilder). "
                  "Excel: AZ (52), Access: Value52 (Memo field). "
                  "Sessions separated by '||'. Normalized history in history app.",
    )
    
    # -------------------------------------------------------------------------
    # Timestamps (Django-managed)
    # -------------------------------------------------------------------------
    
    created_at = models.DateTimeField(
        auto_now_add=True,
        help_text="Timestamp when record was created in Django.",
    )
    
    updated_at = models.DateTimeField(
        auto_now=True,
        help_text="Timestamp when record was last updated in Django.",
    )
    
    # -------------------------------------------------------------------------
    # Class-level constants for external access
    # -------------------------------------------------------------------------
    
    MONTH_FIELDS: ClassVar[tuple[str, ...]] = MONTH_FIELD_NAMES
    TRACKED_FIELDS: ClassVar[tuple[str, ...]] = TRACKED_FIELDS
    
    class Meta:
        db_table = "karteien_record"
        ordering = ["year", "family_id", "id"]
        verbose_name = "Kartei Record"
        verbose_name_plural = "Kartei Records"
        constraints = [
            models.UniqueConstraint(
                fields=["year", "id"],
                name="unique_kartei_year_id",
            ),
        ]
        indexes = [
            models.Index(fields=["year", "status"], name="idx_kartei_year_status"),
            models.Index(fields=["family_id", "year"], name="idx_kartei_family_year"),
        ]
    
    def __str__(self) -> str:
        return f"{self.family_id} – {self.parent_name} – {self.child_name} ({self.year})"
    
    # -------------------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------------------
    
    def get_month_value(self, month_num: int) -> Decimal | None:
        """
        Get value for a specific month (1-12).
        
        Args:
            month_num: Month number from 1 to 12.
            
        Returns:
            The month's value or None if not set.
            
        Raises:
            ValueError: If month_num is not in range 1-12.
        """
        if not 1 <= month_num <= 12:
            raise ValueError(f"month_num must be 1-12, got {month_num}")
        return getattr(self, f"month_{month_num}")
    
    def set_month_value(self, month_num: int, value: Decimal | None) -> None:
        """
        Set value for a specific month (1-12).
        
        Args:
            month_num: Month number from 1 to 12.
            value: The value to set (Decimal or None).
            
        Raises:
            ValueError: If month_num is not in range 1-12.
        """
        if not 1 <= month_num <= 12:
            raise ValueError(f"month_num must be 1-12, got {month_num}")
        setattr(self, f"month_{month_num}", value)
    
    def get_all_months(self) -> dict[int, Decimal | None]:
        """
        Get all month values as a dictionary.
        
        Returns:
            Dict mapping month number (1-12) to value.
        """
        return {i: self.get_month_value(i) for i in range(1, 13)}
    
    @property
    def is_pending(self) -> bool:
        """Check if record is in PENDING status."""
        return self.status == RecordStatus.PENDING
    
    @property
    def is_declined(self) -> bool:
        """Check if record is in DECLINED status."""
        return self.status == RecordStatus.DECLINED
    
    @property
    def is_sepa(self) -> bool:
        """Check if record has SEPA marker (restricts Operator edits)."""
        return self.sepa_marker.upper() == "SEPA"
