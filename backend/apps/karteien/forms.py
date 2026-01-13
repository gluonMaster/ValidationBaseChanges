"""
Forms for the karteien app.

This module contains:
- KarteiRecordForm: Form for creating/editing KarteiRecord entries
- KarteiRecordFilterForm: Form for filtering the record list
- MonthsOverrideForm: Emergency override form for month values

Forms include validation for:
- FamilyID + Parent uniqueness
- Empty FamilyID with non-empty Parent
- SEPA restrictions (Operator)
- Past-months restrictions (Operator)
- Catalog references consistency (subject/teacher/price)
- AUTO mode billing calculations
"""

from __future__ import annotations

import re
from datetime import date
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from django import forms
from django.core.exceptions import ValidationError

from .billing import (
    is_individual_subject,
    is_nachhilfe_subject,
    is_per_hour_subject,
    round_money_up,
    normalize_hours,
    recalculate_record_months,
    get_semester_for_month,
    get_subject_name_for_semester,
    build_base_amounts,
    calculate_month_values,
    detect_meaningful_changes,
    recalculate_legacy_to_auto,
    ZERO,
)
from .models import KarteiRecord, MONTH_FIELD_NAMES, RecordStatus, MonthsMode
from .validators import (
    validate_family_id_parent_unique,
    validate_family_id_not_empty_with_parent,
    validate_sepa_restrictions,
    validate_past_months_restrictions,
    get_allowed_months,
)

if TYPE_CHECKING:
    from apps.accounts.models import User


# Semester month ranges
SEMESTER_1_MONTHS = (1, 2, 3, 4, 5, 6)
SEMESTER_2_MONTHS = (7, 8, 9, 10, 11, 12)

# Start month choices
START_MONTH_1_CHOICES = [(i, f"Monat {i}") for i in range(1, 7)]
START_MONTH_2_CHOICES = [(i, f"Monat {i}") for i in range(7, 13)]

# Contract type choices (form-only field)
CONTRACT_TYPE_CHOICES = [
    ("yearly", "Jährlich"),
    ("monthly", "Monatlich (O/V)"),
]

# Contract status choices (form-only field)
CONTRACT_STATUS_CHOICES = [
    ("active", "Aktiv"),
    ("terminated", "Gekündigt (KN)"),
]


# =============================================================================
# Contract Raw Field Helpers
# =============================================================================

def add_ov_marker(raw_value: str) -> str:
    """
    Add 'O/V' marker to contract_type_raw if not present.
    
    Careful insertion: appends with space if raw is not empty.
    """
    raw_value = raw_value or ""
    # Check if O/V already present (case-insensitive)
    if re.search(r"O/V", raw_value, re.IGNORECASE):
        return raw_value
    # Append O/V
    if raw_value.strip():
        return raw_value.strip() + " O/V"
    return "O/V"


def remove_ov_marker(raw_value: str) -> str:
    """
    Remove all 'O/V' occurrences from contract_type_raw.
    
    Case-insensitive removal, then normalize whitespace.
    O/V can appear directly attached to numbers, so we remove it carefully.
    """
    raw_value = raw_value or ""
    # Remove all O/V occurrences (case-insensitive)
    result = re.sub(r"O/V", "", raw_value, flags=re.IGNORECASE)
    # Normalize whitespace: collapse multiple spaces, strip
    result = re.sub(r"\s+", " ", result).strip()
    return result


def add_kn_marker(raw_value: str) -> str:
    """
    Add 'KN' token to contract_status_raw if not present.
    
    KN is always added as a separate token (space-separated).
    """
    raw_value = raw_value or ""
    # Check if KN already present as separate token
    if re.search(r"(^|\s)KN(\s|$)", raw_value, re.IGNORECASE):
        return raw_value
    # Append KN
    if raw_value.strip():
        return raw_value.strip() + " KN"
    return "KN"


def remove_kn_marker(raw_value: str) -> str:
    """
    Remove 'KN' token from contract_status_raw (only as separate token).
    
    Uses word boundary matching to only remove KN when it's a standalone token.
    """
    raw_value = raw_value or ""
    # Remove KN only as a separate token (beginning/end or surrounded by whitespace)
    result = re.sub(r"(^|\s)KN(\s|$)", r"\1\2", raw_value, flags=re.IGNORECASE)
    # Normalize whitespace
    result = re.sub(r"\s+", " ", result).strip()
    return result


# =============================================================================
# Filter Form
# =============================================================================

class KarteiRecordFilterForm(forms.Form):
    """
    Form for filtering the KarteiRecord list.
    """
    
    year = forms.IntegerField(
        required=False,
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    
    family_id = forms.CharField(
        required=False,
        max_length=50,
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "placeholder": "FamilyID",
        }),
    )
    
    parent = forms.CharField(
        required=False,
        max_length=255,
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "placeholder": "Eltern",
        }),
    )
    
    child = forms.CharField(
        required=False,
        max_length=255,
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "placeholder": "Kind",
        }),
    )
    
    status = forms.ChoiceField(
        required=False,
        choices=[
            ("", "Alle Status"),
            ("NORMAL", "Normal"),
            ("PENDING", "Wartend"),
            ("DECLINED", "Abgelehnt"),
        ],
        widget=forms.Select(attrs={"class": "form-select"}),
    )


# =============================================================================
# Record Form
# =============================================================================

class KarteiRecordForm(forms.ModelForm):
    """
    Form for creating and editing KarteiRecord entries.
    
    Includes validation for:
    - FamilyID + Parent uniqueness
    - Empty FamilyID with non-empty Parent
    - SEPA restrictions (for Operator role)
    - Past-months restrictions (for Operator role)
    - Catalog references consistency (subject/teacher/price)
    - AUTO mode billing calculations with hours input
    
    The form must be initialized with:
    - user: The user making the change
    - year: The year context for the record
    """
    
    class Meta:
        model = KarteiRecord
        fields = [
            # Basic info
            "family_id",
            "parent_name",
            "child_name",
            "birthdate",
            "address",
            "phone",
            "mobile",
            "email",
            # Catalog references (1st semester / Fach 1)
            "subject1_ref",
            "teacher1_ref",
            "price1_ref",
            "start_month_1",
            # Catalog references (2nd semester / Fach 2)
            "subject2_ref",
            "teacher2_ref",
            "price2_ref",
            "start_month_2",
            # Legacy subjects/prices (hidden, synced from refs)
            "subject1",
            "price1",
            "subject2",
            "price2",
            # Extras (free text, no catalog refs)
            "extra1",
            "extra2",
            "extra3",
            # Months
            "month_1",
            "month_2",
            "month_3",
            "month_4",
            "month_5",
            "month_6",
            "month_7",
            "month_8",
            "month_9",
            "month_10",
            "month_11",
            "month_12",
            # SEPA marker
            "sepa_marker",
            # Billing mode
            "discounts_disabled",
        ]
        widgets = {
            "family_id": forms.TextInput(attrs={"class": "form-control"}),
            "parent_name": forms.TextInput(attrs={"class": "form-control"}),
            "child_name": forms.TextInput(attrs={"class": "form-control"}),
            "birthdate": forms.DateInput(
                attrs={"class": "form-control", "type": "date"},
                format="%Y-%m-%d",
            ),
            "address": forms.TextInput(attrs={"class": "form-control"}),
            "phone": forms.TextInput(attrs={"class": "form-control"}),
            "mobile": forms.TextInput(attrs={"class": "form-control"}),
            "email": forms.EmailInput(attrs={"class": "form-control"}),
            # Catalog reference selects
            "subject1_ref": forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
            "teacher1_ref": forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
            "price1_ref": forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
            "start_month_1": forms.Select(attrs={"class": "form-select"}, choices=START_MONTH_1_CHOICES),
            "subject2_ref": forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
            "teacher2_ref": forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
            "price2_ref": forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
            "start_month_2": forms.Select(attrs={"class": "form-select"}, choices=START_MONTH_2_CHOICES),
            # Legacy fields (hidden)
            "subject1": forms.HiddenInput(),
            "price1": forms.HiddenInput(),
            "subject2": forms.HiddenInput(),
            "price2": forms.HiddenInput(),
            # Extras
            "extra1": forms.TextInput(attrs={"class": "form-control"}),
            "extra2": forms.TextInput(attrs={"class": "form-control"}),
            "extra3": forms.TextInput(attrs={"class": "form-control"}),
            "sepa_marker": forms.Select(attrs={"class": "form-select"}),
            "discounts_disabled": forms.CheckboxInput(attrs={"class": "form-check-input"}),
        }
        # Month widgets
        for i in range(1, 13):
            widgets[f"month_{i}"] = forms.NumberInput(
                attrs={"class": "form-control", "step": "0.01"}
            )
    
    # Comment field for risky changes (Notitzen)
    comment = forms.CharField(
        required=False,
        max_length=1000,
        widget=forms.Textarea(attrs={
            "class": "form-control",
            "rows": 3,
            "placeholder": "Kommentar zur Änderung (optional)",
        }),
        help_text="Kommentar wird bei riskanten Änderungen in der Historie gespeichert.",
    )
    
    # Contract type/status fields (form-only, Admin only)
    # These control is_monthly_contract and is_contract_terminated boolean flags
    # and update the raw fields (contract_type_raw, contract_status_raw)
    contract_type = forms.ChoiceField(
        required=False,
        choices=CONTRACT_TYPE_CHOICES,
        initial="yearly",
        widget=forms.Select(attrs={"class": "form-select"}),
        help_text="Jährlich (Standard) oder Monatlich (O/V-Marker).",
    )
    
    contract_status = forms.ChoiceField(
        required=False,
        choices=CONTRACT_STATUS_CHOICES,
        initial="active",
        widget=forms.Select(attrs={"class": "form-select"}),
        help_text="Aktiv (Standard) oder Gekündigt (KN-Marker).",
    )
    
    # Hours fields for Individual/Nachhilfe subjects (non-model fields)
    hours_month_1 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
        help_text="Akademische Stunden (UE)",
    )
    hours_month_2 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_3 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_4 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_5 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_6 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_7 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_8 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_9 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_10 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_11 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    hours_month_12 = forms.DecimalField(
        required=False, max_digits=6, decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01", "min": "0"}),
    )
    
    # Price change application month (for edit mode when price changes)
    apply_from_month_1 = forms.ChoiceField(
        required=False,
        choices=[('', '---')] + START_MONTH_1_CHOICES,
        widget=forms.Select(attrs={"class": "form-select"}),
        help_text="Ab welchem Monat soll die neue Preis 1 gelten?",
    )
    apply_from_month_2 = forms.ChoiceField(
        required=False,
        choices=[('', '---')] + START_MONTH_2_CHOICES,
        widget=forms.Select(attrs={"class": "form-select"}),
        help_text="Ab welchem Monat soll die neue Preis 2 gelten?",
    )
    
    # Confirmation checkbox for negative discount clamping
    confirm_zero_clamp = forms.BooleanField(
        required=False,
        widget=forms.CheckboxInput(attrs={"class": "form-check-input"}),
        help_text="Ich bestätige, dass die auf 0 geklemmten Werte korrekt sind.",
    )
    
    def __init__(self, *args, **kwargs) -> None:
        """
        Initialize the form with user and year context.
        
        Args:
            user: The user making the change (required).
            year: The year context for the record (required).
        """
        self.user: "User" = kwargs.pop("user", None)
        self.year: int = kwargs.pop("year", date.today().year)
        
        super().__init__(*args, **kwargs)
        
        # Store original price refs for detecting changes
        self._original_price1_ref_id = None
        self._original_price2_ref_id = None
        if self.instance and self.instance.pk:
            self._original_price1_ref_id = self.instance.price1_ref_id
            self._original_price2_ref_id = self.instance.price2_ref_id
        
        # Configure catalog reference fields
        self._configure_catalog_fields()
        
        # Configure SEPA marker as choice field
        self._configure_sepa_marker_field()
        
        # Prefill refs from legacy fields if possible
        self._prefill_refs_from_legacy()
        
        # Configure hours fields from record
        self._configure_hours_fields()
        
        # Configure contract type/status fields (Admin only)
        self._configure_contract_fields()
        
        # Configure AUTO mode restrictions
        self._configure_auto_mode()
        
        # Apply restrictions for Operator
        self._apply_operator_restrictions()
    
    def _configure_hours_fields(self) -> None:
        """
        Prefill hours fields from record.hours_amounts.
        """
        if not self.instance or not self.instance.pk:
            return
        
        hours_data = self.instance.hours_amounts or {}
        for i in range(1, 13):
            field_key = f"month_{i}"
            hours_field = f"hours_month_{i}"
            hours_value = hours_data.get(field_key)
            if hours_value is not None:
                try:
                    self.initial[hours_field] = Decimal(str(hours_value))
                except (ValueError, TypeError):
                    pass
    
    def _configure_contract_fields(self) -> None:
        """
        Configure contract type/status fields based on record's boolean flags.
        
        For edit mode: initialize from is_monthly_contract / is_contract_terminated
        For create mode: default to yearly + active
        
        These fields are only shown to Admin role users in the template.
        """
        if self.instance and self.instance.pk:
            # Edit mode: initialize from record's boolean flags
            self.initial["contract_type"] = (
                "monthly" if self.instance.is_monthly_contract else "yearly"
            )
            self.initial["contract_status"] = (
                "terminated" if self.instance.is_contract_terminated else "active"
            )
        else:
            # Create mode: defaults (yearly + active)
            self.initial["contract_type"] = "yearly"
            self.initial["contract_status"] = "active"
    
    def _configure_auto_mode(self) -> None:
        """
        Configure form fields for AUTO mode.
        
        In AUTO mode:
        - Month fields are disabled (read-only)
        - Hours fields are enabled only for Individual/Nachhilfe months
        - Price change fields are shown for edit mode
        
        In LEGACY mode:
        - Month fields remain enabled (they are inherited from legacy import)
        - No automatic calculation happens unless a "meaningful" change is made
        - Upon meaningful change, record converts to AUTO mode (see _process_auto_mode)
        """
        is_auto_mode = False
        is_legacy_mode = False
        
        if self.instance and self.instance.pk:
            is_auto_mode = self.instance.months_mode == MonthsMode.AUTO
            is_legacy_mode = self.instance.months_mode == MonthsMode.LEGACY
        else:
            # New records will be AUTO by default
            is_auto_mode = True
        
        # Store mode flags for later use
        self._is_auto_mode = is_auto_mode
        self._is_legacy_mode = is_legacy_mode
        
        if is_auto_mode:
            # Disable month fields for AUTO mode (values are calculated)
            for field_name in MONTH_FIELD_NAMES:
                self.fields[field_name].disabled = True
                self.fields[field_name].widget.attrs['readonly'] = True
                self.fields[field_name].widget.attrs['class'] = 'form-control bg-light'
        elif is_legacy_mode:
            # LEGACY mode: month fields stay enabled but show informative styling
            # No disabling - legacy values are preserved until meaningful change
            for field_name in MONTH_FIELD_NAMES:
                self.fields[field_name].widget.attrs['class'] = 'form-control'
        
        # Determine which months need hours input
        self._hours_required_months = self._get_hourly_months()
    
    def _get_hourly_months(self) -> list[int]:
        """
        Get list of months that use per-hour pricing.
        
        Returns:
            List of month numbers (1-12) that have Individual/Nachhilfe subjects.
        """
        hourly_months = []
        
        # Get subject names for each semester
        subject1_name = None
        subject2_name = None
        
        # Check submitted data first, fall back to instance
        subject1_ref = self.data.get('subject1_ref') if self.data else None
        subject2_ref = self.data.get('subject2_ref') if self.data else None
        
        if subject1_ref:
            from apps.catalog.models import Subject
            try:
                subject1_name = Subject.objects.get(pk=subject1_ref).name
            except Subject.DoesNotExist:
                pass
        elif self.instance:
            if self.instance.subject1_ref_id:
                subject1_name = self.instance.subject1_ref.name
            else:
                subject1_name = self.instance.subject1
        
        if subject2_ref:
            from apps.catalog.models import Subject
            try:
                subject2_name = Subject.objects.get(pk=subject2_ref).name
            except Subject.DoesNotExist:
                pass
        elif self.instance:
            if self.instance.subject2_ref_id:
                subject2_name = self.instance.subject2_ref.name
            else:
                subject2_name = self.instance.subject2
        
        # Check each semester
        if is_per_hour_subject(subject1_name):
            hourly_months.extend(SEMESTER_1_MONTHS)
        if is_per_hour_subject(subject2_name):
            hourly_months.extend(SEMESTER_2_MONTHS)
        
        return hourly_months
    
    def _configure_catalog_fields(self) -> None:
        """
        Configure catalog reference fields with filtered querysets for the current year.
        """
        from apps.catalog.models import Subject, Teacher, PriceOption, TeachingAssignment
        
        # Get active subjects
        active_subjects = Subject.objects.filter(is_active=True).order_by("name")
        
        # For teacher and price, we need to filter by year
        # Initially show all active teachers/prices for the year
        # JS will further filter based on selected subject
        
        # Get teachers who have assignments in this year
        teacher_ids = TeachingAssignment.objects.filter(
            year=self.year, is_active=True
        ).values_list("teacher_id", flat=True).distinct()
        active_teachers = Teacher.objects.filter(
            id__in=teacher_ids, is_active=True
        ).order_by("last_name", "first_name")
        
        # Get prices for this year
        active_prices = PriceOption.objects.filter(
            year=self.year, is_active=True
        ).select_related("subject").order_by("subject__name", "amount")
        
        # Set querysets
        self.fields["subject1_ref"].queryset = active_subjects
        self.fields["subject1_ref"].required = False
        self.fields["subject2_ref"].queryset = active_subjects
        self.fields["subject2_ref"].required = False
        
        self.fields["teacher1_ref"].queryset = active_teachers
        self.fields["teacher1_ref"].required = False
        self.fields["teacher2_ref"].queryset = active_teachers
        self.fields["teacher2_ref"].required = False
        
        self.fields["price1_ref"].queryset = active_prices
        self.fields["price1_ref"].required = False
        self.fields["price2_ref"].queryset = active_prices
        self.fields["price2_ref"].required = False
        
        # Start month choices
        self.fields["start_month_1"].choices = START_MONTH_1_CHOICES
        self.fields["start_month_2"].choices = START_MONTH_2_CHOICES
    
    def _configure_sepa_marker_field(self) -> None:
        """
        Configure SEPA-Marker field as a choice field.
        
        Builds choices from:
        - Empty option first: ("", "— nicht gesetzt —")
        - "SEPA" is always included
        - All distinct non-empty sepa_marker values from database
        - Current instance value (if non-standard)
        """
        # Get distinct non-empty SEPA marker values from database
        sepa_values = (
            KarteiRecord.objects
            .exclude(sepa_marker__isnull=True)
            .exclude(sepa_marker="")
            .values_list("sepa_marker", flat=True)
            .distinct()
        )
        sepa_set = set(sepa_values)
        
        # Always ensure "SEPA" is available as an option
        sepa_set.add("SEPA")
        
        # If editing, ensure current value is in choices
        if self.instance and self.instance.pk and self.instance.sepa_marker:
            sepa_set.add(self.instance.sepa_marker)
        
        # Build choices: empty option first, then sorted values
        sepa_choices = [("", "— nicht gesetzt —")]
        sepa_choices += [(v, v) for v in sorted(sepa_set)]
        
        # Override the field as a ChoiceField
        self.fields["sepa_marker"] = forms.ChoiceField(
            required=False,
            choices=sepa_choices,
            widget=forms.Select(attrs={"class": "form-select"}),
            label="SEPA-Marker",
        )
        
        # Set initial value from instance
        if self.instance and self.instance.pk:
            self.initial["sepa_marker"] = self.instance.sepa_marker or ""
    
    def _prefill_refs_from_legacy(self) -> None:
        """
        Try to prefill *_ref fields from legacy subject1/subject2 values.
        
        Only applies to existing records where ref is NULL but legacy field has value.
        Uses normalized string matching for subjects (case-insensitive, whitespace-normalized).
        Uses filter() for prices to handle 0 or >1 matches gracefully.
        """
        from apps.catalog.models import Subject, PriceOption
        
        if not self.instance or not self.instance.pk:
            return
        
        def normalize_name(name: str) -> str:
            """Normalize subject name for matching: strip, collapse whitespace, casefold."""
            if not name:
                return ""
            # Strip leading/trailing whitespace
            name = name.strip()
            # Collapse multiple whitespace to single space
            name = re.sub(r"\s+", " ", name)
            # Case-insensitive comparison via casefold
            return name.casefold()
        
        def find_subject_by_legacy_name(legacy_name: str) -> int | None:
            """
            Find subject by legacy name using normalized matching.
            Returns subject id if exactly one match found, None otherwise.
            """
            if not legacy_name:
                return None
            
            normalized_legacy = normalize_name(legacy_name)
            if not normalized_legacy:
                return None
            
            # Build map of active subjects by normalized name
            active_subjects = Subject.objects.filter(is_active=True)
            matches = []
            for subj in active_subjects:
                if normalize_name(subj.name) == normalized_legacy:
                    matches.append(subj)
            
            # Only return if exactly one match (unambiguous)
            if len(matches) == 1:
                return matches[0].id
            return None
        
        # Subject 1
        if not self.instance.subject1_ref_id and self.instance.subject1:
            subject_id = find_subject_by_legacy_name(self.instance.subject1)
            if subject_id:
                self.initial["subject1_ref"] = subject_id
        
        # Subject 2
        if not self.instance.subject2_ref_id and self.instance.subject2:
            subject_id = find_subject_by_legacy_name(self.instance.subject2)
            if subject_id:
                self.initial["subject2_ref"] = subject_id
        
        # Price 1 - try to match by amount and subject using filter()
        if not self.instance.price1_ref_id and self.instance.price1 is not None:
            subject_ref = self.initial.get("subject1_ref") or (
                self.instance.subject1_ref_id
            )
            if subject_ref:
                prices = PriceOption.objects.filter(
                    year=self.instance.year,
                    subject_id=subject_ref,
                    amount=self.instance.price1,
                    is_active=True,
                )
                # Only set if exactly one match found
                if prices.count() == 1:
                    self.initial["price1_ref"] = prices.first().id
        
        # Price 2
        if not self.instance.price2_ref_id and self.instance.price2 is not None:
            subject_ref = self.initial.get("subject2_ref") or (
                self.instance.subject2_ref_id
            )
            if subject_ref:
                prices = PriceOption.objects.filter(
                    year=self.instance.year,
                    subject_id=subject_ref,
                    amount=self.instance.price2,
                    is_active=True,
                )
                # Only set if exactly one match found
                if prices.count() == 1:
                    self.initial["price2_ref"] = prices.first().id
    
    def _apply_operator_restrictions(self) -> None:
        """
        Apply field restrictions for Operator role.
        
        - Disable month fields for past months
        - Disable month fields for SEPA records
        """
        if not self.user:
            return
        
        is_update = self.instance and self.instance.pk
        
        # Past months restrictions
        if self.user.has_past_months_restrictions and is_update:
            allowed_months, reason = get_allowed_months(self.instance.year)
            
            for field_name in MONTH_FIELD_NAMES:
                if field_name not in allowed_months:
                    self.fields[field_name].disabled = True
                    self.fields[field_name].help_text = "Vergangener Monat - nicht änderbar"
        
        # SEPA restrictions
        if self.user.has_sepa_restrictions and is_update:
            if self.instance.is_sepa:
                for field_name in MONTH_FIELD_NAMES:
                    self.fields[field_name].disabled = True
                    self.fields[field_name].help_text = "SEPA-Datensatz - nicht änderbar"
    
    def clean(self) -> dict[str, Any]:
        """
        Run all Kartei-specific validations and billing calculations.
        """
        cleaned_data = super().clean()
        
        family_id = cleaned_data.get("family_id", "")
        parent_name = cleaned_data.get("parent_name", "")
        
        # Validate FamilyID + Parent uniqueness
        exclude_pk = self.instance.pk if self.instance else None
        unique_result = validate_family_id_parent_unique(
            family_id, parent_name, self.year, exclude_pk
        )
        if not unique_result.is_valid:
            raise ValidationError(unique_result.errors[0])
        
        # Validate empty FamilyID with non-empty Parent
        empty_result = validate_family_id_not_empty_with_parent(family_id, parent_name)
        if not empty_result.is_valid:
            raise ValidationError(empty_result.errors[0])
        
        # Validate catalog references consistency
        self._validate_catalog_refs(cleaned_data)
        
        # Normalize money fields
        self._normalize_money_fields(cleaned_data)
        
        # Process AUTO mode calculations
        self._process_auto_mode(cleaned_data)
        
        # Operator restrictions (additional check on form level)
        if self.user and self.instance and self.instance.pk:
            # Build changes dict
            changes = {}
            for field_name in MONTH_FIELD_NAMES:
                old_val = getattr(self.instance, field_name, None)
                new_val = cleaned_data.get(field_name)
                if old_val != new_val:
                    changes[field_name] = new_val
            
            if changes:
                # SEPA restrictions
                if self.user.has_sepa_restrictions:
                    sepa_result = validate_sepa_restrictions(
                        self.instance, changes, self.user
                    )
                    if not sepa_result.is_valid:
                        raise ValidationError(sepa_result.errors[0])
                
                # Past months restrictions
                if self.user.has_past_months_restrictions:
                    past_result = validate_past_months_restrictions(
                        self.instance, changes, self.user
                    )
                    if not past_result.is_valid:
                        raise ValidationError(past_result.errors[0])
        
        # Sync legacy fields from refs
        self._sync_legacy_from_refs(cleaned_data)
        
        return cleaned_data
    
    def _normalize_money_fields(self, cleaned_data: dict[str, Any]) -> None:
        """
        Normalize all money fields to 2 decimal places with ceiling rounding.
        
        Also normalizes hours fields with standard rounding.
        """
        # Normalize month values
        for field_name in MONTH_FIELD_NAMES:
            value = cleaned_data.get(field_name)
            cleaned_data[field_name] = round_money_up(value) if value is not None else ZERO
        
        # Normalize price fields
        for price_field in ['price1', 'price2']:
            value = cleaned_data.get(price_field)
            if value is not None:
                cleaned_data[price_field] = round_money_up(value)
        
        # Normalize hours fields
        for i in range(1, 13):
            hours_field = f"hours_month_{i}"
            value = cleaned_data.get(hours_field)
            cleaned_data[hours_field] = normalize_hours(value) if value is not None else ZERO
    
    def _process_auto_mode(self, cleaned_data: dict[str, Any]) -> None:
        """
        Process AUTO mode billing calculations.
        
        For AUTO mode records:
        1. Build hours_amounts from form data
        2. Check for price changes and require apply_from_month
        3. Calculate base_amounts and month values with discounts
        4. Validate discount clamping confirmation
        
        For LEGACY mode records:
        1. Detect if there are "meaningful" changes that affect billing
        2. If yes, mark record for LEGACY->AUTO conversion with touched_months
        3. If no, leave record as LEGACY without changing months
        """
        is_auto_mode = getattr(self, '_is_auto_mode', False)
        is_legacy_mode = getattr(self, '_is_legacy_mode', False)
        is_edit = self.instance and self.instance.pk
        
        if not is_edit:
            # New records default to AUTO
            is_auto_mode = True
        
        # Build hours_amounts from form
        hours_amounts = {}
        for i in range(1, 13):
            hours_field = f"hours_month_{i}"
            hours_value = cleaned_data.get(hours_field, ZERO)
            hours_amounts[f"month_{i}"] = str(normalize_hours(hours_value))
        
        # Store hours on instance
        self._billing_hours_amounts = hours_amounts
        
        # Initialize flags
        self._has_meaningful_changes = False
        self._touched_months = set()
        self._should_convert_to_auto = False
        
        if is_legacy_mode and is_edit:
            # LEGACY mode: detect meaningful changes
            original = KarteiRecord.objects.get(pk=self.instance.pk)
            has_changes, touched_months = detect_meaningful_changes(
                original, cleaned_data, hours_amounts
            )
            
            self._has_meaningful_changes = has_changes
            self._touched_months = touched_months
            self._should_convert_to_auto = has_changes
            
            if not has_changes:
                # No meaningful changes - don't process billing, keep LEGACY
                self._apply_from_month_1 = None
                self._apply_from_month_2 = None
                return
            
            # Meaningful changes detected - will convert to AUTO
            # Fall through to AUTO mode processing for calculation validation
            is_auto_mode = True
        
        if not is_auto_mode:
            return
        
        # Check for price changes (edit mode only)
        apply_from_1 = None
        apply_from_2 = None
        
        if is_edit:
            # Check if price1 changed
            new_price1_ref = cleaned_data.get('price1_ref')
            new_price1_ref_id = new_price1_ref.id if new_price1_ref else None
            
            if new_price1_ref_id != self._original_price1_ref_id:
                # Price 1 changed - require apply_from_month_1
                apply_from_str = cleaned_data.get('apply_from_month_1')
                if not apply_from_str:
                    raise ValidationError({
                        'apply_from_month_1': 'Bitte wählen Sie, ab welchem Monat der neue Preis 1 gelten soll.'
                    })
                try:
                    apply_from_1 = int(apply_from_str)
                except (ValueError, TypeError):
                    raise ValidationError({
                        'apply_from_month_1': 'Ungültiger Monat.'
                    })
            
            # Check if price2 changed
            new_price2_ref = cleaned_data.get('price2_ref')
            new_price2_ref_id = new_price2_ref.id if new_price2_ref else None
            
            if new_price2_ref_id != self._original_price2_ref_id:
                # Price 2 changed - require apply_from_month_2
                apply_from_str = cleaned_data.get('apply_from_month_2')
                if not apply_from_str:
                    raise ValidationError({
                        'apply_from_month_2': 'Bitte wählen Sie, ab welchem Monat der neue Preis 2 gelten soll.'
                    })
                try:
                    apply_from_2 = int(apply_from_str)
                except (ValueError, TypeError):
                    raise ValidationError({
                        'apply_from_month_2': 'Ungültiger Monat.'
                    })
        
        # Store apply_from values for view to use
        self._apply_from_month_1 = apply_from_1
        self._apply_from_month_2 = apply_from_2
        
        # Calculate preliminary values to check for clamping
        # Actual calculation will be done in the view with the saved instance
        # Create temp record for both create and edit
        if is_edit:
            # Edit: load existing record and apply form changes
            temp_record = KarteiRecord.objects.get(pk=self.instance.pk)
        else:
            # Create: build unsaved record from form data
            temp_record = KarteiRecord()
            temp_record.year = self.year
        
        # Apply form changes to temp record
        temp_record.hours_amounts = hours_amounts
        for field in ['family_id', 'subject1_ref', 'price1_ref', 'subject2_ref', 'price2_ref',
                      'start_month_1', 'start_month_2', 'discounts_disabled']:
            value = cleaned_data.get(field)
            if field.endswith('_ref') and value:
                setattr(temp_record, f"{field}_id", value.id)
                setattr(temp_record, field, value)
            else:
                setattr(temp_record, field, value)
        
        # Build base amounts
        base_amounts = build_base_amounts(
            temp_record,
            apply_from_month_1=apply_from_1,
            apply_from_month_2=apply_from_2,
            hours_amounts=hours_amounts,
        )
        
        # Calculate with discounts
        month_values, flags = calculate_month_values(
            temp_record,
            base_amounts=base_amounts,
        )
        
        # Store calculated values for view (always, even before validation)
        self._calculated_month_values = month_values
        self._calculated_base_amounts = base_amounts
        self._calculation_flags = flags
        self._clamped_to_zero_months = list(flags.clamped_to_zero_months) if flags.clamped_to_zero_months else []
        
        # Check if confirmation is required
        if flags.requires_confirmation:
            confirm = cleaned_data.get('confirm_zero_clamp', False)
            if not confirm:
                months_str = ", ".join(str(m) for m in flags.clamped_to_zero_months)
                raise ValidationError({
                    'confirm_zero_clamp': (
                        f"Die Rabatte führen in Monat(en) {months_str} zu negativen Werten, "
                        f"die auf 0 geklemmt werden. Bitte bestätigen Sie dies."
                    )
                })
    
    def _validate_catalog_refs(self, cleaned_data: dict[str, Any]) -> None:
        """
        Validate catalog reference consistency:
        - price*_ref must belong to subject*_ref and current year
        - teacher*_ref must have TeachingAssignment for subject*_ref and year
        """
        from apps.catalog.models import TeachingAssignment
        
        # Validate Semester 1 (subject1/teacher1/price1)
        subject1_ref = cleaned_data.get("subject1_ref")
        teacher1_ref = cleaned_data.get("teacher1_ref")
        price1_ref = cleaned_data.get("price1_ref")
        
        if price1_ref:
            # Price must match year
            if price1_ref.year != self.year:
                raise ValidationError({
                    "price1_ref": f"Preis gehört zu Jahr {price1_ref.year}, nicht {self.year}."
                })
            # If subject is selected, price must match it
            if subject1_ref and price1_ref.subject_id != subject1_ref.id:
                raise ValidationError({
                    "price1_ref": f"Preis gehört zu Fach '{price1_ref.subject.name}', nicht '{subject1_ref.name}'."
                })
        
        if teacher1_ref and subject1_ref:
            # Check TeachingAssignment exists
            assignment_exists = TeachingAssignment.objects.filter(
                year=self.year,
                subject=subject1_ref,
                teacher=teacher1_ref,
                is_active=True,
            ).exists()
            if not assignment_exists:
                raise ValidationError({
                    "teacher1_ref": f"Lehrer '{teacher1_ref}' ist dem Fach '{subject1_ref.name}' im Jahr {self.year} nicht zugeordnet."
                })
        
        # Validate Semester 2 (subject2/teacher2/price2)
        subject2_ref = cleaned_data.get("subject2_ref")
        teacher2_ref = cleaned_data.get("teacher2_ref")
        price2_ref = cleaned_data.get("price2_ref")
        
        if price2_ref:
            # Price must match year
            if price2_ref.year != self.year:
                raise ValidationError({
                    "price2_ref": f"Preis gehört zu Jahr {price2_ref.year}, nicht {self.year}."
                })
            # If subject is selected, price must match it
            if subject2_ref and price2_ref.subject_id != subject2_ref.id:
                raise ValidationError({
                    "price2_ref": f"Preis gehört zu Fach '{price2_ref.subject.name}', nicht '{subject2_ref.name}'."
                })
        
        if teacher2_ref and subject2_ref:
            # Check TeachingAssignment exists
            assignment_exists = TeachingAssignment.objects.filter(
                year=self.year,
                subject=subject2_ref,
                teacher=teacher2_ref,
                is_active=True,
            ).exists()
            if not assignment_exists:
                raise ValidationError({
                    "teacher2_ref": f"Lehrer '{teacher2_ref}' ist dem Fach '{subject2_ref.name}' im Jahr {self.year} nicht zugeordnet."
                })
        
        # Validate start_month ranges
        start_month_1 = cleaned_data.get("start_month_1", 1)
        start_month_2 = cleaned_data.get("start_month_2", 7)
        
        if start_month_1 is not None and start_month_1 not in SEMESTER_1_MONTHS:
            raise ValidationError({
                "start_month_1": f"Startmonat 1. HJ muss zwischen 1 und 6 liegen."
            })
        
        if start_month_2 is not None and start_month_2 not in SEMESTER_2_MONTHS:
            raise ValidationError({
                "start_month_2": f"Startmonat 2. HJ muss zwischen 7 und 12 liegen."
            })
    
    def _sync_legacy_from_refs(self, cleaned_data: dict[str, Any]) -> None:
        """
        Sync legacy subject1/subject2/price1/price2 fields from catalog refs.
        
        Only syncs if ref is selected; does NOT clear legacy values if ref is empty.
        """
        # Subject 1
        subject1_ref = cleaned_data.get("subject1_ref")
        if subject1_ref:
            cleaned_data["subject1"] = subject1_ref.name
        
        # Price 1
        price1_ref = cleaned_data.get("price1_ref")
        if price1_ref:
            cleaned_data["price1"] = price1_ref.amount
        
        # Subject 2
        subject2_ref = cleaned_data.get("subject2_ref")
        if subject2_ref:
            cleaned_data["subject2"] = subject2_ref.name
        
        # Price 2
        price2_ref = cleaned_data.get("price2_ref")
        if price2_ref:
            cleaned_data["price2"] = price2_ref.amount
    
    def get_comment(self) -> str:
        """Get the comment for history/pending changes."""
        return self.cleaned_data.get("comment", "")
    
    def get_hourly_months(self) -> list[int]:
        """
        Get list of months that require hours input.
        
        Used by templates to show/hide hours input fields.
        """
        return getattr(self, '_hours_required_months', [])
    
    def is_auto_mode(self) -> bool:
        """Check if form is in AUTO mode."""
        if self.instance and self.instance.pk:
            return self.instance.months_mode == MonthsMode.AUTO
        return True  # New records are AUTO by default
    
    def get_billing_data(self) -> dict[str, Any]:
        """
        Get billing calculation data for the view.
        
        Returns dict with:
        - hours_amounts: dict of hours per month
        - apply_from_month_1: starting month for price1 change
        - apply_from_month_2: starting month for price2 change
        - calculated_month_values: dict of final month values (edit mode)
        - calculated_base_amounts: dict of base amounts (edit mode)
        - flags: CalculationFlags (edit mode)
        - should_convert_to_auto: True if LEGACY record should become AUTO
        - touched_months: set of months affected by meaningful changes
        """
        return {
            'hours_amounts': getattr(self, '_billing_hours_amounts', {}),
            'apply_from_month_1': getattr(self, '_apply_from_month_1', None),
            'apply_from_month_2': getattr(self, '_apply_from_month_2', None),
            'calculated_month_values': getattr(self, '_calculated_month_values', None),
            'calculated_base_amounts': getattr(self, '_calculated_base_amounts', None),
            'flags': getattr(self, '_calculation_flags', None),
            'should_convert_to_auto': getattr(self, '_should_convert_to_auto', False),
            'touched_months': getattr(self, '_touched_months', set()),
        }
    
    def save(self, commit: bool = True) -> KarteiRecord:
        """
        Save the record with contract type/status updates.
        
        Updates both the boolean flags and raw text fields for contract type/status
        if the user is an Admin.
        """
        record = super().save(commit=False)
        
        # Process contract fields (Admin only)
        if self.user and self.user.is_admin_role:
            self._apply_contract_fields(record)
        
        if commit:
            record.save()
        
        return record
    
    def _apply_contract_fields(self, record: KarteiRecord) -> None:
        """
        Apply contract type/status changes to the record.
        
        Updates:
        - is_monthly_contract and contract_type_raw based on contract_type field
        - is_contract_terminated and contract_status_raw based on contract_status field
        """
        contract_type = self.cleaned_data.get("contract_type", "yearly")
        contract_status = self.cleaned_data.get("contract_status", "active")
        
        # Process contract type
        if contract_type == "monthly":
            record.is_monthly_contract = True
            record.contract_type_raw = add_ov_marker(record.contract_type_raw)
        else:  # yearly
            record.is_monthly_contract = False
            record.contract_type_raw = remove_ov_marker(record.contract_type_raw)
        
        # Process contract status
        if contract_status == "terminated":
            record.is_contract_terminated = True
            record.contract_status_raw = add_kn_marker(record.contract_status_raw)
        else:  # active
            record.is_contract_terminated = False
            record.contract_status_raw = remove_kn_marker(record.contract_status_raw)


# =============================================================================
# Inline Forms for Months
# =============================================================================

class MonthsInlineForm(forms.Form):
    """
    Inline form for editing just the month fields.
    
    Used in quick-edit scenarios where only months need updating.
    """
    
    month_1 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_2 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_3 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_4 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_5 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_6 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_7 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_8 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_9 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_10 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_11 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )
    month_12 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control form-control-sm", "step": "0.01"}),
    )


# =============================================================================
# Months Override Form (Emergency Admin-only tool)
# =============================================================================

class MonthsOverrideForm(forms.Form):
    """
    Emergency override form for manually setting month values.
    
    This form is used for admin-only emergency corrections.
    It requires confirmation and sets the record to OVERRIDE mode.
    Changes go through the approvals workflow.
    """
    
    # Override reason (required)
    reason = forms.CharField(
        required=True,
        max_length=1000,
        widget=forms.Textarea(attrs={
            "class": "form-control",
            "rows": 3,
            "placeholder": "Grund für die manuelle Änderung (erforderlich)",
        }),
        help_text="Bitte geben Sie den Grund für diese Änderung an.",
    )
    
    # Confirmation checkbox
    confirm_override = forms.BooleanField(
        required=True,
        widget=forms.CheckboxInput(attrs={"class": "form-check-input"}),
        label="Ich verstehe die Konsequenzen",
        help_text="Diese Änderung überschreibt die automatische Berechnung.",
    )
    
    # Month fields
    month_1 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_2 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_3 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_4 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_5 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_6 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_7 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_8 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_9 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_10 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_11 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    month_12 = forms.DecimalField(
        required=False,
        max_digits=10,
        decimal_places=2,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    
    def __init__(self, *args, record: KarteiRecord = None, **kwargs):
        """Initialize with current record values."""
        super().__init__(*args, **kwargs)
        
        self.record = record
        
        if record:
            # Prefill with current month values
            for i in range(1, 13):
                value = getattr(record, f"month_{i}")
                if value is not None:
                    self.initial[f"month_{i}"] = value
    
    def clean(self) -> dict[str, Any]:
        """Normalize all month values."""
        cleaned_data = super().clean()
        
        # Normalize month values
        for i in range(1, 13):
            field_name = f"month_{i}"
            value = cleaned_data.get(field_name)
            cleaned_data[field_name] = round_money_up(value) if value is not None else ZERO
        
        return cleaned_data
    
    def get_month_changes(self) -> dict[str, Decimal]:
        """Get dict of month values from cleaned data."""
        return {
            f"month_{i}": self.cleaned_data.get(f"month_{i}", ZERO)
            for i in range(1, 13)
        }
