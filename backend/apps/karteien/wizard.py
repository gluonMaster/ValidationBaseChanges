"""
New Family Wizard module for the karteien app.

This module implements a wizard for creating new families with multiple children,
discounts, and automatic billing calculation. Similar to VBA "Neue Eltern-Kinder".

Features:
- Generate globally unique FamilyID (format: "1. <number>")
- Create multiple KarteiRecord entries in one pass
- Assign FamilyDiscount and RecordDiscounts
- Auto-calculate month_* values after save

Access: Admin only.
"""

from __future__ import annotations

import re
from datetime import date
from decimal import Decimal
from typing import Any

from django import forms
from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db import transaction
from django.forms import formset_factory
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views import View

from apps.catalog.models import (
    Discount,
    DiscountKind,
    FamilyDiscount,
    PriceOption,
    RecordDiscount,
    Subject,
    Teacher,
    TeachingAssignment,
)

from .billing import (
    recalculate_record_months,
    is_per_hour_subject,
    ZERO,
)
from .models import KarteiRecord, RecordStatus, MonthsMode


# =============================================================================
# FamilyID Generation Helper
# =============================================================================

# Pattern for FamilyID format: "1. <number>"
FAMILY_ID_PATTERN = re.compile(r'^1\.\s*(\d+)$')


def generate_next_family_id() -> str:
    """
    Generate the next globally unique FamilyID.
    
    Scans all existing family_id values across all years that match
    the pattern "1. <number>" and returns "1. <max+1>".
    
    Returns:
        str: Next FamilyID in format "1. <number>".
    """
    # Get all family_ids that match the pattern
    all_family_ids = KarteiRecord.objects.values_list('family_id', flat=True).distinct()
    
    max_number = 0
    for fid in all_family_ids:
        if not fid:
            continue
        match = FAMILY_ID_PATTERN.match(str(fid).strip())
        if match:
            number = int(match.group(1))
            if number > max_number:
                max_number = number
    
    # Generate next FamilyID
    next_number = max_number + 1
    return f"1. {next_number}"


def validate_family_id_globally_unique(family_id: str) -> bool:
    """
    Check if a family_id is globally unique (across all years).
    
    Args:
        family_id: The FamilyID to check.
        
    Returns:
        True if unique, False if already exists.
    """
    return not KarteiRecord.objects.filter(family_id=family_id).exists()


# =============================================================================
# Permission Mixin
# =============================================================================

class AdminOnlyMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to Admin users only.
    """
    
    def test_func(self) -> bool:
        """Check if user is Admin."""
        user = self.request.user
        return user.is_authenticated and user.is_admin_role
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error for unauthorized users."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        messages.error(
            self.request,
            "Diese Funktion ist nur für Administratoren verfügbar."
        )
        return redirect("karteien:record_list")


# =============================================================================
# Form: Family Header
# =============================================================================

class FamilyHeaderForm(forms.Form):
    """
    Form for family header information (shared across all children).
    
    Includes contact info, address, and optional FamilyDiscount.
    """
    
    year = forms.IntegerField(
        widget=forms.Select(attrs={"class": "form-select"}),
        label="Jahr",
        help_text="Für welches Jahr sollen die Datensätze erstellt werden?",
    )
    
    family_id = forms.CharField(
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "readonly": "readonly",
        }),
        label="FamilyID",
        help_text="Wird automatisch generiert.",
    )
    
    parent_name = forms.CharField(
        max_length=255,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Eltern",
        help_text="Name der Eltern / Erziehungsberechtigten.",
    )
    
    address = forms.CharField(
        max_length=500,
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Adresse",
    )
    
    phone = forms.CharField(
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Telefon",
    )
    
    mobile = forms.CharField(
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Mobil",
    )
    
    email = forms.EmailField(
        max_length=255,
        required=False,
        widget=forms.EmailInput(attrs={"class": "form-control"}),
        label="E-Mail",
    )
    
    sepa_marker = forms.CharField(
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="SEPA-Marker",
        help_text="z.B. SEPA, Bar, Rechnung",
    )
    
    is_monthly_contract = forms.BooleanField(
        required=False,
        initial=False,
        widget=forms.CheckboxInput(attrs={"class": "form-check-input"}),
        label="Monatsvertrag",
        help_text="Ankreuzen für Monatsvertrag (O/V), sonst Jahresvertrag.",
    )

    is_contract_terminated = forms.BooleanField(
        required=False,
        initial=False,
        widget=forms.CheckboxInput(attrs={"class": "form-check-input"}),
        label="Vertrag gekündigt",
        help_text="Ankreuzen für gekündigten Vertrag (KN).",
    )
    
    # Family discount (optional)
    family_discount = forms.ModelChoiceField(
        queryset=Discount.objects.filter(is_active=True).order_by('kind', '-value'),
        required=False,
        widget=forms.Select(attrs={"class": "form-select"}),
        label="Familienrabatt",
        help_text="Optional: Rabatt für die gesamte Familie (alle Monate).",
    )
    
    def __init__(self, *args, **kwargs) -> None:
        """Initialize form with year choices."""
        super().__init__(*args, **kwargs)

        # Build year choices from existing data + current/next + selected year (if any)
        current_year = date.today().year
        selected_year = None

        if self.is_bound:
            try:
                selected_year = int(self.data.get("year"))
            except (TypeError, ValueError):
                selected_year = None
        else:
            try:
                selected_year = int(self.initial.get("year")) if self.initial.get("year") else None
            except (TypeError, ValueError):
                selected_year = None

        existing_years = list(
            KarteiRecord.objects.values_list("year", flat=True).distinct()
        )
        year_set = set(existing_years) | {current_year, current_year + 1}
        if selected_year is not None:
            year_set.add(selected_year)

        self.fields["year"].widget.choices = [(y, str(y)) for y in sorted(year_set)]
        if not self.is_bound and "year" not in (self.initial or {}):
            self.fields["year"].initial = current_year
        
        # Generate initial FamilyID
        self.fields["family_id"].initial = generate_next_family_id()


# =============================================================================
# Form: Child Record
# =============================================================================

class ChildRecordForm(forms.Form):
    """
    Form for a single child/record in the family wizard.
    
    Includes child info, subject/teacher/price selections for both semesters,
    and optional record discounts.
    """
    
    child_name = forms.CharField(
        max_length=255,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Kind",
    )
    
    birthdate = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={
            "class": "form-control",
            "type": "date",
        }),
        label="Geburtsdatum",
    )
    
    # Semester 1 (months 1-6)
    subject1_ref = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
        label="Fach 1 (1. HJ)",
    )
    
    teacher1_ref = forms.ModelChoiceField(
        queryset=Teacher.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
        label="Lehrer 1",
    )
    
    price1_ref = forms.ModelChoiceField(
        queryset=PriceOption.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "1"}),
        label="Preis 1",
    )
    
    start_month_1 = forms.ChoiceField(
        choices=[(i, f"Monat {i}") for i in range(1, 7)],
        initial=1,
        widget=forms.Select(attrs={"class": "form-select"}),
        label="Ab Monat (1. HJ)",
    )
    
    # Semester 2 (months 7-12)
    subject2_ref = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
        label="Fach 2 (2. HJ)",
    )
    
    teacher2_ref = forms.ModelChoiceField(
        queryset=Teacher.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
        label="Lehrer 2",
    )
    
    price2_ref = forms.ModelChoiceField(
        queryset=PriceOption.objects.none(),
        required=False,
        widget=forms.Select(attrs={"class": "form-select", "data-semester": "2"}),
        label="Preis 2",
    )
    
    start_month_2 = forms.ChoiceField(
        choices=[(i, f"Monat {i}") for i in range(7, 13)],
        initial=7,
        widget=forms.Select(attrs={"class": "form-select"}),
        label="Ab Monat (2. HJ)",
    )
    
    # Billing options
    discounts_disabled = forms.BooleanField(
        required=False,
        widget=forms.CheckboxInput(attrs={"class": "form-check-input"}),
        label="Rabatte deaktiviert",
        help_text="Keine Rabatte für dieses Kind anwenden.",
    )
    
    # Record discounts (multiple selection)
    record_discounts = forms.ModelMultipleChoiceField(
        queryset=Discount.objects.filter(is_active=True).order_by('kind', '-value'),
        required=False,
        widget=forms.CheckboxSelectMultiple(attrs={"class": "form-check-input"}),
        label="Eintragrabatte",
        help_text="Diese Rabatte gelten nur für dieses Kind.",
    )
    
    # Hours for Individual/Nachhilfe subjects (simplified: same for all months)
    hours_per_month_1 = forms.DecimalField(
        required=False,
        max_digits=6,
        decimal_places=2,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
            "step": "0.01",
            "min": "0",
            "placeholder": "0.00",
        }),
        label="Stunden/Monat (1. HJ)",
        help_text="Nur für Ind./NH Fächer: UE pro Monat.",
    )
    
    hours_per_month_2 = forms.DecimalField(
        required=False,
        max_digits=6,
        decimal_places=2,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
            "step": "0.01",
            "min": "0",
            "placeholder": "0.00",
        }),
        label="Stunden/Monat (2. HJ)",
        help_text="Nur für Ind./NH Fächer: UE pro Monat.",
    )
    
    def __init__(self, *args, year: int = None, **kwargs) -> None:
        """
        Initialize form with catalog data for the specified year.
        
        Args:
            year: The year for filtering catalog data.
        """
        self.year = year or date.today().year
        super().__init__(*args, **kwargs)
        
        self._configure_catalog_fields()
    
    def _configure_catalog_fields(self) -> None:
        """Configure catalog reference fields with filtered querysets."""
        # Active subjects
        active_subjects = Subject.objects.filter(is_active=True).order_by("name")
        
        # Teachers with assignments for this year
        teacher_ids = TeachingAssignment.objects.filter(
            year=self.year, is_active=True
        ).values_list("teacher_id", flat=True).distinct()
        active_teachers = Teacher.objects.filter(
            id__in=teacher_ids, is_active=True
        ).order_by("last_name", "first_name")
        
        # Prices for this year
        active_prices = PriceOption.objects.filter(
            year=self.year, is_active=True
        ).select_related("subject").order_by("subject__name", "amount")
        
        # Set querysets
        self.fields["subject1_ref"].queryset = active_subjects
        self.fields["subject2_ref"].queryset = active_subjects
        
        self.fields["teacher1_ref"].queryset = active_teachers
        self.fields["teacher2_ref"].queryset = active_teachers
        
        self.fields["price1_ref"].queryset = active_prices
        self.fields["price2_ref"].queryset = active_prices
    
    def clean(self) -> dict[str, Any]:
        """Validate child record data."""
        cleaned_data = super().clean()
        
        # Validate price matches subject and year
        subject1_ref = cleaned_data.get("subject1_ref")
        price1_ref = cleaned_data.get("price1_ref")
        teacher1_ref = cleaned_data.get("teacher1_ref")
        
        if price1_ref:
            if price1_ref.year != self.year:
                self.add_error(
                    "price1_ref",
                    f"Preis gehört zu Jahr {price1_ref.year}, nicht {self.year}."
                )
            if subject1_ref and price1_ref.subject_id != subject1_ref.id:
                self.add_error(
                    "price1_ref",
                    f"Preis gehört zu Fach '{price1_ref.subject.name}', nicht '{subject1_ref.name}'."
                )
        
        if teacher1_ref and subject1_ref:
            assignment_exists = TeachingAssignment.objects.filter(
                year=self.year,
                subject=subject1_ref,
                teacher=teacher1_ref,
                is_active=True,
            ).exists()
            if not assignment_exists:
                self.add_error(
                    "teacher1_ref",
                    f"Lehrer ist dem Fach im Jahr {self.year} nicht zugeordnet."
                )
        
        # Same for semester 2
        subject2_ref = cleaned_data.get("subject2_ref")
        price2_ref = cleaned_data.get("price2_ref")
        teacher2_ref = cleaned_data.get("teacher2_ref")
        
        if price2_ref:
            if price2_ref.year != self.year:
                self.add_error(
                    "price2_ref",
                    f"Preis gehört zu Jahr {price2_ref.year}, nicht {self.year}."
                )
            if subject2_ref and price2_ref.subject_id != subject2_ref.id:
                self.add_error(
                    "price2_ref",
                    f"Preis gehört zu Fach '{price2_ref.subject.name}', nicht '{subject2_ref.name}'."
                )
        
        if teacher2_ref and subject2_ref:
            assignment_exists = TeachingAssignment.objects.filter(
                year=self.year,
                subject=subject2_ref,
                teacher=teacher2_ref,
                is_active=True,
            ).exists()
            if not assignment_exists:
                self.add_error(
                    "teacher2_ref",
                    f"Lehrer ist dem Fach im Jahr {self.year} nicht zugeordnet."
                )
        
        return cleaned_data


# =============================================================================
# Create ChildRecordFormSet
# =============================================================================

def get_child_formset_class(extra: int = 1, min_num: int = 1, max_num: int = 10):
    """
    Create a formset class for child records.
    
    Args:
        extra: Number of extra empty forms.
        min_num: Minimum number of forms.
        max_num: Maximum number of forms.
    
    Returns:
        A formset class for ChildRecordForm.
    """
    return formset_factory(
        ChildRecordForm,
        extra=extra,
        min_num=min_num,
        max_num=max_num,
        validate_min=True,
        validate_max=True,
    )


# =============================================================================
# New Family Wizard View
# =============================================================================

class NewFamilyWizardView(AdminOnlyMixin, View):
    """
    Wizard view for creating a new family with multiple children.
    
    Flow:
    1. Display form with family header + child formset
    2. On submit: validate all forms
    3. Create FamilyDiscount if specified
    4. Create KarteiRecord for each child
    5. Create RecordDiscounts for each child
    6. Recalculate months for all records
    7. Redirect to family dashboard
    """
    
    template_name = "karteien/new_family_wizard.html"
    
    def get(self, request: HttpRequest) -> HttpResponse:
        """Display the wizard form."""
        # Get year from query param or default to current
        year = request.GET.get("year")
        try:
            year = int(year) if year else date.today().year
        except ValueError:
            year = date.today().year
        
        header_form = FamilyHeaderForm(initial={"year": year})
        ChildFormSet = get_child_formset_class(extra=1, min_num=1)
        child_formset = ChildFormSet(form_kwargs={"year": year})
        
        return render(request, self.template_name, {
            "header_form": header_form,
            "child_formset": child_formset,
            "year": year,
        })
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Process the wizard form submission."""
        header_form = FamilyHeaderForm(request.POST)
        
        # Get year from form data
        year = request.POST.get("year")
        try:
            year = int(year) if year else date.today().year
        except ValueError:
            year = date.today().year
        
        ChildFormSet = get_child_formset_class(extra=0, min_num=1)
        child_formset = ChildFormSet(request.POST, form_kwargs={"year": year})
        
        # Validate all forms
        if not header_form.is_valid() or not child_formset.is_valid():
            return render(request, self.template_name, {
                "header_form": header_form,
                "child_formset": child_formset,
                "year": year,
            })
        
        # Process in a transaction
        try:
            with transaction.atomic():
                created_records = self._create_family(
                    request,
                    header_form,
                    child_formset,
                )
                
                if not created_records:
                    messages.error(request, "Es wurden keine Datensätze erstellt.")
                    return render(request, self.template_name, {
                        "header_form": header_form,
                        "child_formset": child_formset,
                        "year": year,
                    })
                
                # Success message
                family_id = created_records[0].family_id
                year = created_records[0].year
                children_count = len(created_records)
                
                messages.success(
                    request,
                    f"Familie '{family_id}' mit {children_count} Kind(ern) wurde erstellt."
                )
                
                # Redirect to family dashboard
                return redirect(
                    f"{reverse('karteien:family_dashboard')}?year={year}&family_id={family_id}"
                )
        
        except Exception as e:
            messages.error(request, f"Fehler beim Erstellen der Familie: {str(e)}")
            return render(request, self.template_name, {
                "header_form": header_form,
                "child_formset": child_formset,
                "year": year,
            })
    
    def _create_family(
        self,
        request: HttpRequest,
        header_form: FamilyHeaderForm,
        child_formset,
    ) -> list[KarteiRecord]:
        """
        Create family with all records and discounts.
        
        Args:
            request: The HTTP request.
            header_form: Validated header form.
            child_formset: Validated child formset.
            
        Returns:
            List of created KarteiRecord instances.
        """
        header_data = header_form.cleaned_data
        year = header_data["year"]
        
        # Generate FamilyID (re-generate to ensure uniqueness)
        family_id = generate_next_family_id()
        
        # Ensure it's unique
        while not validate_family_id_globally_unique(family_id):
            # Extract number and increment
            match = FAMILY_ID_PATTERN.match(family_id)
            if match:
                next_num = int(match.group(1)) + 1
                family_id = f"1. {next_num}"
            else:
                break
        
        # Get max ID for this year
        max_id = KarteiRecord.objects.filter(year=year).order_by('-id').values_list('id', flat=True).first()
        next_id = (max_id or 0) + 1
        
        # Create FamilyDiscount if specified
        family_discount_obj = header_data.get("family_discount")
        if family_discount_obj:
            FamilyDiscount.objects.create(
                year=year,
                family_id=family_id,
                discount=family_discount_obj,
                start_month=1,
                end_month=12,
            )
        
        # Create records for each child
        created_records: list[KarteiRecord] = []
        user = request.user
        
        for child_form in child_formset:
            if not child_form.cleaned_data:
                continue
            
            child_data = child_form.cleaned_data
            child_name = child_data.get("child_name", "").strip()
            
            if not child_name:
                continue
            
            # Build hours_amounts
            hours_per_month_1 = child_data.get("hours_per_month_1") or Decimal("0.00")
            hours_per_month_2 = child_data.get("hours_per_month_2") or Decimal("0.00")
            
            hours_amounts = {}
            start_month_1 = int(child_data.get("start_month_1") or 1)
            start_month_2 = int(child_data.get("start_month_2") or 7)
            
            for i in range(1, 7):
                hours_amounts[f"month_{i}"] = str(hours_per_month_1) if i >= start_month_1 else "0.00"
            for i in range(7, 13):
                hours_amounts[f"month_{i}"] = str(hours_per_month_2) if i >= start_month_2 else "0.00"
            
            # Create record
            record = KarteiRecord(
                # Domain key
                year=year,
                id=next_id,
                
                # Family info
                family_id=family_id,
                parent_name=header_data["parent_name"],
                address=header_data.get("address", ""),
                phone=header_data.get("phone", ""),
                mobile=header_data.get("mobile", ""),
                email=header_data.get("email", ""),
                sepa_marker=header_data.get("sepa_marker", ""),
                is_monthly_contract=header_data.get("is_monthly_contract", False),
                contract_type_raw="O/V" if header_data.get("is_monthly_contract") else "",
                is_contract_terminated=header_data.get("is_contract_terminated", False),
                contract_status_raw="KN" if header_data.get("is_contract_terminated") else "",
                
                # Child info
                child_name=child_name,
                birthdate=child_data.get("birthdate"),
                
                # Subject/Teacher/Price references (semester 1)
                subject1_ref=child_data.get("subject1_ref"),
                teacher1_ref=child_data.get("teacher1_ref"),
                price1_ref=child_data.get("price1_ref"),
                start_month_1=start_month_1,
                
                # Subject/Teacher/Price references (semester 2)
                subject2_ref=child_data.get("subject2_ref"),
                teacher2_ref=child_data.get("teacher2_ref"),
                price2_ref=child_data.get("price2_ref"),
                start_month_2=start_month_2,
                
                # Billing
                months_mode=MonthsMode.AUTO,
                discounts_disabled=child_data.get("discounts_disabled", False),
                hours_amounts=hours_amounts,
                
                # Status
                status=RecordStatus.NORMAL,
                
                # Metadata
                last_change_role=user.role,
                last_change_date=date.today(),
                last_change_time=date.today().strftime("%H:%M"),
            )
            
            # Sync legacy fields from refs
            if record.subject1_ref:
                record.subject1 = record.subject1_ref.name
            if record.price1_ref:
                record.price1 = record.price1_ref.amount
            if record.subject2_ref:
                record.subject2 = record.subject2_ref.name
            if record.price2_ref:
                record.price2 = record.price2_ref.amount
            
            # Save record first (needed for RecordDiscount FK)
            record.save()
            
            # Create RecordDiscounts
            record_discounts = child_data.get("record_discounts", [])
            for discount in record_discounts:
                RecordDiscount.objects.create(
                    record=record,
                    discount=discount,
                    start_month=1,
                    end_month=12,
                )
            
            # Recalculate months with discounts
            recalculate_record_months(record, hours_amounts=hours_amounts)
            record.save()
            
            created_records.append(record)
            next_id += 1
        
        return created_records
