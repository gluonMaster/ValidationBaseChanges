"""
Forms for catalog app.

Provides forms for creating and editing Teachers, Subjects,
TeachingAssignments, PriceOptions, Discounts, and copying data between years.
"""

import json
from datetime import date

from django import forms

from .models import (
    Teacher, Subject, TeachingAssignment, PriceOption,
    Discount, DiscountKind, FamilyDiscount, RecordDiscount
)


class TeacherForm(forms.ModelForm):
    """Form for creating and editing Teachers."""
    
    class Meta:
        model = Teacher
        fields = ["last_name", "first_name", "is_active"]
        widgets = {
            "last_name": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Nachname",
            }),
            "first_name": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Vorname",
            }),
            "is_active": forms.CheckboxInput(attrs={
                "class": "form-check-input",
            }),
        }


class SubjectForm(forms.ModelForm):
    """Form for creating and editing Subjects."""
    
    class Meta:
        model = Subject
        fields = ["name", "is_active"]
        widgets = {
            "name": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Fachname",
            }),
            "is_active": forms.CheckboxInput(attrs={
                "class": "form-check-input",
            }),
        }


class TeachingAssignmentForm(forms.ModelForm):
    """Form for creating TeachingAssignments."""
    
    class Meta:
        model = TeachingAssignment
        fields = ["year", "subject", "teacher", "is_active"]
        widgets = {
            "year": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 2000,
                "max": 2100,
            }),
            "subject": forms.Select(attrs={
                "class": "form-select",
            }),
            "teacher": forms.Select(attrs={
                "class": "form-select",
            }),
            "is_active": forms.CheckboxInput(attrs={
                "class": "form-check-input",
            }),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Filter only active subjects and teachers by default
        self.fields["subject"].queryset = Subject.objects.filter(is_active=True)
        self.fields["teacher"].queryset = Teacher.objects.filter(is_active=True)
        
        # Set default year to current year
        if not self.instance.pk:
            self.fields["year"].initial = date.today().year


class CopyYearForm(forms.Form):
    """Form for copying TeachingAssignments from one year to another."""
    
    from_year = forms.IntegerField(
        label="Von Jahr",
        min_value=2000,
        max_value=2100,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
        }),
    )
    to_year = forms.IntegerField(
        label="Nach Jahr",
        min_value=2000,
        max_value=2100,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
        }),
    )
    only_active = forms.BooleanField(
        label="Nur aktive Zuweisungen kopieren",
        required=False,
        initial=True,
        widget=forms.CheckboxInput(attrs={
            "class": "form-check-input",
        }),
    )
    overwrite = forms.BooleanField(
        label="Bestehende Zuweisungen im Zieljahr ersetzen",
        required=False,
        initial=False,
        widget=forms.CheckboxInput(attrs={
            "class": "form-check-input",
        }),
        help_text="Wenn aktiviert, werden alle bestehenden Zuweisungen im Zieljahr zuerst deaktiviert.",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        current_year = date.today().year
        self.fields["from_year"].initial = current_year
        self.fields["to_year"].initial = current_year + 1

    def clean(self):
        cleaned_data = super().clean()
        from_year = cleaned_data.get("from_year")
        to_year = cleaned_data.get("to_year")
        
        if from_year and to_year and from_year == to_year:
            raise forms.ValidationError(
                "Quell- und Zieljahr dürfen nicht identisch sein."
            )
        
        return cleaned_data


class PriceOptionForm(forms.ModelForm):
    """Form for creating and editing PriceOptions."""
    
    class Meta:
        model = PriceOption
        fields = ["year", "subject", "amount", "comment", "is_active"]
        widgets = {
            "year": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 2000,
                "max": 2100,
            }),
            "subject": forms.Select(attrs={
                "class": "form-select",
            }),
            "amount": forms.NumberInput(attrs={
                "class": "form-control",
                "step": "0.01",
                "min": "0",
            }),
            "comment": forms.Textarea(attrs={
                "class": "form-control",
                "rows": 3,
                "placeholder": "Warum dieser Preis? (optional)",
            }),
            "is_active": forms.CheckboxInput(attrs={
                "class": "form-check-input",
            }),
        }

    def __init__(self, *args, require_comment: bool = False, **kwargs):
        """Initialize form with optional require_comment flag.
        
        Args:
            require_comment: If True, comment field becomes required.
                             Used in quick-add workflow from legacy records.
        """
        super().__init__(*args, **kwargs)
        # Filter only active subjects by default
        self.fields["subject"].queryset = Subject.objects.filter(is_active=True)
        
        # Set default year to current year
        if not self.instance.pk:
            self.fields["year"].initial = date.today().year
        
        # Make comment required in quick-add mode
        if require_comment:
            self.fields["comment"].required = True
            self.fields["comment"].widget.attrs["placeholder"] = (
                "Bitte Begründung angeben (Pflichtfeld im Schnell-Workflow)"
            )


class CopyPricesYearForm(forms.Form):
    """Form for copying PriceOptions from one year to another."""
    
    from_year = forms.IntegerField(
        label="Von Jahr",
        min_value=2000,
        max_value=2100,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
        }),
    )
    to_year = forms.IntegerField(
        label="Nach Jahr",
        min_value=2000,
        max_value=2100,
        widget=forms.NumberInput(attrs={
            "class": "form-control",
        }),
    )
    only_active = forms.BooleanField(
        label="Nur aktive Preise kopieren",
        required=False,
        initial=True,
        widget=forms.CheckboxInput(attrs={
            "class": "form-check-input",
        }),
    )
    overwrite = forms.BooleanField(
        label="Bestehende Preise im Zieljahr ersetzen",
        required=False,
        initial=False,
        widget=forms.CheckboxInput(attrs={
            "class": "form-check-input",
        }),
        help_text="Wenn aktiviert, werden alle bestehenden Preise im Zieljahr zuerst deaktiviert.",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        current_year = date.today().year
        self.fields["from_year"].initial = current_year
        self.fields["to_year"].initial = current_year + 1

    def clean(self):
        cleaned_data = super().clean()
        from_year = cleaned_data.get("from_year")
        to_year = cleaned_data.get("to_year")
        
        if from_year and to_year and from_year == to_year:
            raise forms.ValidationError(
                "Quell- und Zieljahr dürfen nicht identisch sein."
            )
        
        return cleaned_data


# =============================================================================
# Discount Forms
# =============================================================================

class DiscountForm(forms.ModelForm):
    """Form for creating and editing Discounts."""
    
    class Meta:
        model = Discount
        fields = ["kind", "value", "description", "is_active"]
        widgets = {
            "kind": forms.Select(attrs={
                "class": "form-select",
            }),
            "value": forms.NumberInput(attrs={
                "class": "form-control",
                "step": "0.01",
                "min": "0",
                "max": "0.99",
            }),
            "description": forms.Textarea(attrs={
                "class": "form-control",
                "rows": 2,
                "placeholder": "Beschreibung der Rabattart",
            }),
            "is_active": forms.CheckboxInput(attrs={
                "class": "form-check-input",
            }),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Update widget based on kind if editing
        if self.instance and self.instance.pk and self.instance.kind == DiscountKind.FIXED:
            self.fields["value"].widget.attrs.update({
                "max": "99999.99",
            })


class MonthsWidget(forms.TextInput):
    """Widget for entering months as comma-separated list."""
    
    def format_value(self, value):
        if value is None:
            return ''
        if isinstance(value, list):
            return ', '.join(str(m) for m in sorted(value))
        return value


class BaseDiscountAssignmentForm(forms.ModelForm):
    """Base form for FamilyDiscount and RecordDiscount with month handling."""
    
    months_input = forms.CharField(
        required=False,
        label='Einzelne Monate',
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "placeholder": "z.B. 1, 3, 5 (leer = Bereich verwenden)",
        }),
        help_text='Kommagetrennte Monatszahlen (1-12). Wenn ausgefüllt, wird der Bereich ignoriert.'
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Pre-fill months_input from instance
        if self.instance and self.instance.pk and self.instance.months:
            self.fields['months_input'].initial = ', '.join(str(m) for m in sorted(self.instance.months))
        
        # Filter only active discounts
        self.fields['discount'].queryset = Discount.objects.filter(is_active=True)

    def clean_months_input(self):
        """Parse comma-separated months into a list."""
        value = self.cleaned_data.get('months_input', '').strip()
        if not value:
            return None
        
        try:
            months = []
            for part in value.split(','):
                part = part.strip()
                if part:
                    m = int(part)
                    if m < 1 or m > 12:
                        raise forms.ValidationError(f"Ungültiger Monat: {m}. Muss 1-12 sein.")
                    months.append(m)
            
            if not months:
                return None
            
            # Check for duplicates
            if len(months) != len(set(months)):
                raise forms.ValidationError("Doppelte Monate sind nicht erlaubt.")
            
            return sorted(set(months))
        except ValueError:
            raise forms.ValidationError("Bitte nur Zahlen 1-12, durch Kommas getrennt, eingeben.")

    def clean(self):
        cleaned_data = super().clean()
        months_input = cleaned_data.get('months_input')
        start_month = cleaned_data.get('start_month')
        end_month = cleaned_data.get('end_month')
        
        # Save parsed months to the months field
        cleaned_data['months'] = months_input
        
        # Validate month range if not using months list
        if not months_input and start_month and end_month:
            if start_month > end_month:
                raise forms.ValidationError(
                    "Von-Monat darf nicht größer als Bis-Monat sein."
                )
        
        return cleaned_data

    def save(self, commit=True):
        instance = super().save(commit=False)
        instance.months = self.cleaned_data.get('months')
        if commit:
            instance.save()
        return instance


class FamilyDiscountForm(BaseDiscountAssignmentForm):
    """Form for creating and editing FamilyDiscounts."""
    
    class Meta:
        model = FamilyDiscount
        fields = ["year", "family_id", "discount", "start_month", "end_month"]
        widgets = {
            "year": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 2000,
                "max": 2100,
            }),
            "family_id": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "FamilyID (z.B. FAM001)",
            }),
            "discount": forms.Select(attrs={
                "class": "form-select",
            }),
            "start_month": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 1,
                "max": 12,
            }),
            "end_month": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 1,
                "max": 12,
            }),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set default year to current year
        if not self.instance.pk:
            self.fields["year"].initial = date.today().year


class RecordDiscountForm(BaseDiscountAssignmentForm):
    """Form for creating and editing RecordDiscounts."""
    
    record_pkid = forms.IntegerField(
        label='Record PKID',
        widget=forms.NumberInput(attrs={
            "class": "form-control",
            "placeholder": "PKID der Kartei-Eintrag",
        }),
        help_text='Die eindeutige PKID des Kartei-Eintrags'
    )

    class Meta:
        model = RecordDiscount
        fields = ["discount", "start_month", "end_month"]
        widgets = {
            "discount": forms.Select(attrs={
                "class": "form-select",
            }),
            "start_month": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 1,
                "max": 12,
            }),
            "end_month": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 1,
                "max": 12,
            }),
        }

    def __init__(self, *args, **kwargs):
        self.initial_record = kwargs.pop('record', None)
        super().__init__(*args, **kwargs)
        
        if self.instance and self.instance.pk:
            self.fields['record_pkid'].initial = self.instance.record.pkid
        elif self.initial_record:
            self.fields['record_pkid'].initial = self.initial_record.pkid

    def clean_record_pkid(self):
        """Validate and fetch the KarteiRecord."""
        from apps.karteien.models import KarteiRecord
        
        pkid = self.cleaned_data.get('record_pkid')
        try:
            record = KarteiRecord.objects.get(pkid=pkid)
            return record
        except KarteiRecord.DoesNotExist:
            raise forms.ValidationError(f"Kartei-Eintrag mit PKID {pkid} nicht gefunden.")

    def save(self, commit=True):
        instance = super().save(commit=False)
        instance.record = self.cleaned_data.get('record_pkid')
        instance.months = self.cleaned_data.get('months')
        if commit:
            instance.save()
        return instance
