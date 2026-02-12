"""
Forms for the approvals app.

This module contains:
- DeclinedChangeEditForm — form for editing DeclinedChange.snapshot fields
- PendingChangeEditForm — form for editing PendingChange.snapshot fields (ADMIN)

Used by:
- DeclinedChangeEditView: Edit declined snapshot before re-submitting to pending
- PendingChangeEditView: Edit pending snapshot before superadmin decision (ADMIN)
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any

from django import forms

from apps.karteien.models import KarteiRecord, TRACKED_FIELDS


class DeclinedChangeEditForm(forms.Form):
    """
    Form for editing DeclinedChange.snapshot.

    Dynamically creates fields based on TRACKED_FIELDS, using the field
    definitions from KarteiRecord._meta to get proper field types.

    Handles conversion between snapshot JSON format and Python types:
    - dates: ISO string <-> date
    - decimals: string <-> Decimal
    """

    def __init__(self, *args, snapshot: dict[str, Any] | None = None, **kwargs):
        """
        Initialize the form with snapshot values.

        Args:
            snapshot: The DeclinedChange.snapshot dict to prefill fields.
        """
        super().__init__(*args, **kwargs)
        self._original_snapshot = dict(snapshot) if snapshot else {}
        
        # Create fields from KarteiRecord field definitions
        for field_name in TRACKED_FIELDS:
            model_field = KarteiRecord._meta.get_field(field_name)
            form_field = model_field.formfield()
            
            if form_field is not None:
                # Copy the form field and adjust
                self.fields[field_name] = form_field
                self.fields[field_name].required = False
                
                # Add Bootstrap class
                widget = self.fields[field_name].widget
                if hasattr(widget, 'attrs'):
                    widget.attrs['class'] = widget.attrs.get('class', '') + ' form-control'
        
        # Prefill with snapshot values if provided
        if snapshot:
            self._prefill_from_snapshot(snapshot)

    def _prefill_from_snapshot(self, snapshot: dict[str, Any]) -> None:
        """
        Prefill form initial values from snapshot, converting types as needed.
        
        Args:
            snapshot: The snapshot dict with string representations of values.
        """
        initial_data = {}
        
        for field_name in TRACKED_FIELDS:
            if field_name not in snapshot:
                continue
                
            value = snapshot[field_name]
            
            if value is None or value == '':
                initial_data[field_name] = None
                continue
            
            # Get the form field to check its type
            form_field = self.fields.get(field_name)
            
            if form_field is None:
                continue
            
            # Convert based on field type
            if isinstance(form_field, forms.DateField):
                # ISO string -> date
                if isinstance(value, str):
                    try:
                        initial_data[field_name] = date.fromisoformat(value)
                    except ValueError:
                        initial_data[field_name] = None
                elif isinstance(value, date):
                    initial_data[field_name] = value
                else:
                    initial_data[field_name] = None
            elif isinstance(form_field, forms.DecimalField):
                # String -> Decimal
                if isinstance(value, (int, float)):
                    initial_data[field_name] = Decimal(str(value))
                elif isinstance(value, str):
                    try:
                        initial_data[field_name] = Decimal(value)
                    except InvalidOperation:
                        initial_data[field_name] = None
                elif isinstance(value, Decimal):
                    initial_data[field_name] = value
                else:
                    initial_data[field_name] = None
            else:
                # Other fields (CharField, etc.) - use as-is
                initial_data[field_name] = value
        
        # Set initial values
        for field_name, value in initial_data.items():
            if field_name in self.fields:
                self.fields[field_name].initial = value

    def to_snapshot(self) -> dict[str, Any]:
        """
        Convert cleaned form data back to snapshot format (JSON-serializable).
        
        Returns:
            Dict with all tracked fields in snapshot format:
            - dates as ISO strings
            - decimals as strings
            - other values as-is
            Plus preserved metadata keys from the original snapshot
            (keys starting with "_" that are not tracked fields).
        """
        snapshot: dict[str, Any] = {}
        
        for field_name in TRACKED_FIELDS:
            if field_name not in self.cleaned_data:
                snapshot[field_name] = None
                continue
            
            value = self.cleaned_data[field_name]
            
            if value is None:
                snapshot[field_name] = None
            elif isinstance(value, date):
                snapshot[field_name] = value.isoformat()
            elif isinstance(value, Decimal):
                snapshot[field_name] = str(value)
            else:
                snapshot[field_name] = value

        # Preserve metadata keys from original snapshot (e.g. pending-contract payloads).
        for key, value in self._original_snapshot.items():
            if key not in TRACKED_FIELDS and key.startswith("_"):
                snapshot[key] = value
        
        return snapshot


class PendingChangeEditForm(DeclinedChangeEditForm):
    """
    Form for editing PendingChange.snapshot (ADMIN editing pending changes).

    Extends DeclinedChangeEditForm with an additional required admin_comment
    field so Superadmin can see why Admin edited the pending change.

    UI labels are in German.
    """

    admin_comment = forms.CharField(
        label="Änderungsgrund (erforderlich)",
        widget=forms.Textarea(attrs={
            "class": "form-control",
            "rows": 3,
            "placeholder": "Bitte beschreiben Sie, warum Sie diese Änderung vornehmen...",
        }),
        required=True,
        help_text="Dieser Kommentar wird dem Superadmin angezeigt.",
    )

    def __init__(
        self,
        *args,
        snapshot: dict[str, Any] | None = None,
        admin_comment: str | None = None,
        **kwargs
    ):
        """
        Initialize the form with snapshot and admin comment.

        Args:
            snapshot: The PendingChange.snapshot dict to prefill fields.
            admin_comment: Existing admin_comment from PendingChange.
        """
        super().__init__(*args, snapshot=snapshot, **kwargs)
        
        # Pre-fill admin_comment if provided
        if admin_comment:
            self.fields["admin_comment"].initial = admin_comment
