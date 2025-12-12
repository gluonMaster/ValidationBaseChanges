"""
Forms for the karteien app.

This module contains:
- KarteiRecordForm: Form for creating/editing KarteiRecord entries
- KarteiRecordFilterForm: Form for filtering the record list

Forms include validation for:
- FamilyID + Parent uniqueness
- Empty FamilyID with non-empty Parent
- SEPA restrictions (Operator)
- Past-months restrictions (Operator)
"""

from __future__ import annotations

from datetime import date
from typing import TYPE_CHECKING, Any

from django import forms
from django.core.exceptions import ValidationError

from .models import KarteiRecord, MONTH_FIELD_NAMES, RecordStatus
from .validators import (
    validate_family_id_parent_unique,
    validate_family_id_not_empty_with_parent,
    validate_sepa_restrictions,
    validate_past_months_restrictions,
    get_allowed_months,
)

if TYPE_CHECKING:
    from apps.accounts.models import User


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
            # Subjects and prices
            "subject1",
            "price1",
            "subject2",
            "price2",
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
            "subject1": forms.TextInput(attrs={"class": "form-control"}),
            "price1": forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
            "subject2": forms.TextInput(attrs={"class": "form-control"}),
            "price2": forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
            "extra1": forms.TextInput(attrs={"class": "form-control"}),
            "extra2": forms.TextInput(attrs={"class": "form-control"}),
            "extra3": forms.TextInput(attrs={"class": "form-control"}),
            "sepa_marker": forms.TextInput(attrs={"class": "form-control"}),
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
        
        # Apply restrictions for Operator
        self._apply_operator_restrictions()
    
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
        Run all Kartei-specific validations.
        """
        cleaned_data = super().clean()
        
        family_id = cleaned_data.get("family_id", "")
        parent_name = cleaned_data.get("parent_name", "")
        
        # Validate FamilyID + Parent uniqueness
        exclude_id = self.instance.pk if self.instance else None
        unique_result = validate_family_id_parent_unique(
            family_id, parent_name, self.year, exclude_id
        )
        if not unique_result.is_valid:
            raise ValidationError(unique_result.errors[0])
        
        # Validate empty FamilyID with non-empty Parent
        empty_result = validate_family_id_not_empty_with_parent(family_id, parent_name)
        if not empty_result.is_valid:
            raise ValidationError(empty_result.errors[0])
        
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
        
        return cleaned_data
    
    def get_comment(self) -> str:
        """Get the comment for history/pending changes."""
        return self.cleaned_data.get("comment", "")


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
