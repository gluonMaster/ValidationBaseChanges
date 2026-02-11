"""
Views for the karteien app.

This module contains views for:
- Listing KarteiRecord entries (with filters)
- Creating new KarteiRecord entries
- Editing existing KarteiRecord entries
- Deleting KarteiRecord entries
- Emergency months override (admin-only)

Access is restricted to Admin and Operator roles, with appropriate
restrictions applied (SEPA, past-months) for Operators.
"""

from __future__ import annotations

from datetime import date
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db.models import Count, Exists, OuterRef, Q
from django.http import HttpRequest, HttpResponse, HttpResponseRedirect
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse, reverse_lazy
from django.views import View
from django.views.generic import (
    CreateView,
    DeleteView,
    DetailView,
    ListView,
    UpdateView,
)

from decimal import Decimal

from apps.approvals.services import (
    build_snapshot,
    classify_change,
    create_or_update_pending_change,
    get_changed_tracked_fields,
    write_history_entry,
)
from apps.approvals.models import DeclinedChange, PendingChange
from apps.catalog.models import FamilyDiscount, RecordDiscount

from .billing import recalculate_record_months, build_base_amounts, get_month_mismatches, detect_meaningful_changes
from .forms import KarteiRecordForm, KarteiRecordFilterForm, MonthsOverrideForm
from .models import KarteiRecord, RecordStatus, MonthsMode
from .validators import validate_kartei_record, apply_operator_filters
from apps.familyid_reservations.services import get_next_family_id


# =============================================================================
# Permission Mixins
# =============================================================================

class KarteiViewerMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin for read-only access to Kartei.
    
    Allows Admin, Operator, Superadmin, and User to view Kartei records.
    - Admin/Operator: full edit access
    - Superadmin/User: read-only access (list/detail), no edit/delete
    """
    
    def test_func(self) -> bool:
        """Check if user can view Kartei records (Admin/Operator/Superadmin/User)."""
        user = self.request.user
        return user.is_authenticated and (
            user.can_edit_kartei or user.is_superadmin or user.is_user_role
        )
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error for unauthorized users."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        user = self.request.user
        messages.error(
            self.request,
            "Sie haben keine Berechtigung, Kartei-Einträge anzuzeigen."
        )
        
        # Redirect based on role to avoid infinite loops
        if user.is_user_role:
            return redirect("accounts:user_dashboard")
        else:
            return redirect("login")


class KarteiEditorMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to users who can edit Kartei.
    
    Only Admin and Operator roles can access views using this mixin.
    Superadmin is explicitly excluded from editing.
    """
    
    def test_func(self) -> bool:
        """Check if user can edit Kartei records."""
        user = self.request.user
        return user.is_authenticated and user.can_edit_kartei
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error for unauthorized users."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        user = self.request.user
        messages.error(
            self.request,
            "Sie haben keine Berechtigung, Kartei-Einträge zu bearbeiten."
        )
        
        # Redirect based on role to avoid infinite loops
        if user.is_superadmin:
            return redirect("karteien:record_list")
        elif user.is_user_role:
            return redirect("accounts:user_dashboard")
        else:
            return redirect("login")


# =============================================================================
# List View
# =============================================================================

class KarteiRecordListView(KarteiViewerMixin, ListView):
    """
    List view for KarteiRecord entries.
    
    Supports filtering by:
    - Year
    - FamilyID (partial match)
    - Parent name (partial match)
    - Child name (partial match)
    - Status (PENDING, DECLINED, normal)
    
    Access: Admin, Operator (full), Superadmin (read-only)
    """
    
    model = KarteiRecord
    template_name = "karteien/record_list.html"
    context_object_name = "records"
    paginate_by = 50
    
    def get_queryset(self):
        """Filter records based on query parameters."""
        from apps.catalog.models import FamilyDiscount
        
        qs = KarteiRecord.objects.annotate(
            record_discounts_count=Count('record_discounts', distinct=True),
            has_family_discounts=Exists(
                FamilyDiscount.objects.filter(
                    year=OuterRef('year'),
                    family_id=OuterRef('family_id')
                )
            ),
        )
        
        # Determine selected year with session persistence
        # Priority: 1) GET param, 2) Session, 3) Current year
        year_param = self.request.GET.get("year")
        system_year = date.today().year
        
        if year_param:
            try:
                selected_year = int(year_param)
                # Save to session for persistence
                self.request.session["karteien_selected_year"] = selected_year
            except ValueError:
                # Invalid year - fall back to session or system year
                selected_year = self.request.session.get("karteien_selected_year", system_year)
        else:
            # No year in GET - use session or default to system year
            selected_year = self.request.session.get("karteien_selected_year", system_year)
            # Save system year as default if nothing in session yet
            if "karteien_selected_year" not in self.request.session:
                self.request.session["karteien_selected_year"] = selected_year
        
        # Store selected year for use in get_context_data
        self._selected_year = selected_year
        
        # Filter by selected year
        qs = qs.filter(year=selected_year)
        
        # Filter by FamilyID
        family_id = self.request.GET.get("family_id")
        if family_id:
            qs = qs.filter(family_id__icontains=family_id)
        
        # Filter by Parent
        parent = self.request.GET.get("parent")
        if parent:
            qs = qs.filter(parent_name__icontains=parent)
        
        # Filter by Child
        child = self.request.GET.get("child")
        if child:
            qs = qs.filter(child_name__icontains=child)
        
        # Filter by Status
        status = self.request.GET.get("status")
        if status:
            if status == "PENDING":
                qs = qs.filter(status=RecordStatus.PENDING)
            elif status == "DECLINED":
                qs = qs.filter(status=RecordStatus.DECLINED)
            elif status == "NORMAL":
                qs = qs.filter(status=RecordStatus.NORMAL)
        
        # Filter by contract type (monthly/yearly)
        contract_type = self.request.GET.get("contract_type")
        if contract_type:
            if contract_type == "monthly":
                qs = qs.filter(is_monthly_contract=True)
            elif contract_type == "yearly":
                qs = qs.filter(is_monthly_contract=False)
        
        # Filter by contract status (active/terminated/active_sepa/terminated_sepa)
        contract_status = self.request.GET.get("contract_status")
        if contract_status:
            if contract_status == "active":
                qs = qs.filter(is_contract_terminated=False)
            elif contract_status == "terminated":
                qs = qs.filter(is_contract_terminated=True)
            elif contract_status == "active_sepa":
                qs = qs.filter(is_contract_terminated=False, sepa_marker__iexact="SEPA")
            elif contract_status == "terminated_sepa":
                qs = qs.filter(is_contract_terminated=True, sepa_marker__iexact="SEPA")
        
        # Filter by subject (Unterricht)
        subject_query = self.request.GET.get("subject_query", "").strip()
        subject_semester = self.request.GET.get("subject_semester", "")
        if subject_query:
            if subject_semester == "1":
                qs = qs.filter(
                    Q(subject1__icontains=subject_query) |
                    Q(subject1_ref__name__icontains=subject_query)
                )
            elif subject_semester == "2":
                qs = qs.filter(
                    Q(subject2__icontains=subject_query) |
                    Q(subject2_ref__name__icontains=subject_query)
                )
            else:
                # Both semesters
                qs = qs.filter(
                    Q(subject1__icontains=subject_query) |
                    Q(subject1_ref__name__icontains=subject_query) |
                    Q(subject2__icontains=subject_query) |
                    Q(subject2_ref__name__icontains=subject_query)
                )
        
        # Filter by teacher (Lehrer)
        teacher_query = self.request.GET.get("teacher_query", "").strip()
        teacher_semester = self.request.GET.get("teacher_semester", "")
        if teacher_query:
            if teacher_semester == "1":
                qs = qs.filter(
                    Q(teacher1_legacy_name__icontains=teacher_query) |
                    Q(teacher1_ref__first_name__icontains=teacher_query) |
                    Q(teacher1_ref__last_name__icontains=teacher_query)
                )
            elif teacher_semester == "2":
                qs = qs.filter(
                    Q(teacher2_legacy_name__icontains=teacher_query) |
                    Q(teacher2_ref__first_name__icontains=teacher_query) |
                    Q(teacher2_ref__last_name__icontains=teacher_query)
                )
            else:
                # Both semesters
                qs = qs.filter(
                    Q(teacher1_legacy_name__icontains=teacher_query) |
                    Q(teacher1_ref__first_name__icontains=teacher_query) |
                    Q(teacher1_ref__last_name__icontains=teacher_query) |
                    Q(teacher2_legacy_name__icontains=teacher_query) |
                    Q(teacher2_ref__first_name__icontains=teacher_query) |
                    Q(teacher2_ref__last_name__icontains=teacher_query)
                )
        
        return qs.order_by("family_id", "parent_name", "child_name")
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add filter form and year list to context."""
        context = super().get_context_data(**kwargs)
        
        # Filter form with current values
        context["filter_form"] = KarteiRecordFilterForm(self.request.GET)
        
        # Available years for dropdown
        years = (
            KarteiRecord.objects.values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        
        # Current filter values - use selected year from get_queryset
        context["current_year"] = getattr(self, "_selected_year", date.today().year)
        context["system_year"] = date.today().year  # For badge highlighting
        context["current_filters"] = {
            "family_id": self.request.GET.get("family_id", ""),
            "parent": self.request.GET.get("parent", ""),
            "child": self.request.GET.get("child", ""),
            "status": self.request.GET.get("status", ""),
            "contract_type": self.request.GET.get("contract_type", ""),
            "contract_status": self.request.GET.get("contract_status", ""),
            "subject_semester": self.request.GET.get("subject_semester", ""),
            "subject_query": self.request.GET.get("subject_query", ""),
            "teacher_semester": self.request.GET.get("teacher_semester", ""),
            "teacher_query": self.request.GET.get("teacher_query", ""),
        }
        
        return context


# =============================================================================
# Detail View
# =============================================================================

class KarteiRecordDetailView(KarteiViewerMixin, DetailView):
    """
    Detail view for a single KarteiRecord.
    
    Access: Admin, Operator (full), Superadmin (read-only)
    
    For ADMIN users viewing PENDING records:
    - Shows pending values by default (from PendingChange snapshot)
    - Displays badges on changed fields
    - Provides toggle to switch between pending and current values
    """
    
    model = KarteiRecord
    template_name = "karteien/record_detail.html"
    context_object_name = "record"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add related information to context."""
        context = super().get_context_data(**kwargs)
        record = self.object
        user = self.request.user
        
        # Check for pending/declined changes
        has_pending = hasattr(record, "pending_change") and record.pending_change
        context["has_pending"] = has_pending
        context["has_declined"] = record.declined_changes.exists()
        
        # User restrictions
        context["user_has_sepa_restrictions"] = user.has_sepa_restrictions
        context["user_has_past_months_restrictions"] = user.has_past_months_restrictions
        context["is_sepa_record"] = record.is_sepa
        
        # Month mismatch detection for highlighting suspicious values
        context["mismatch_months"] = list(get_month_mismatches(record))
        context["is_override_mode"] = record.months_mode == MonthsMode.OVERRIDE
        
        # Pending preview for ADMIN users with PENDING records
        context["pending_record_preview"] = None
        context["pending_changed_fields"] = set()
        context["pending_view_mode"] = None
        context["show_pending_toggle"] = False
        
        if (
            record.status == RecordStatus.PENDING
            and has_pending
            and record.pending_change
            and not record.pending_change.is_processed
            and user.is_admin_role
        ):
            # Admin viewing pending record - prepare preview data
            context["show_pending_toggle"] = True
            
            # Determine view mode from query param (default: pending for admin)
            view_param = self.request.GET.get("view", "pending")
            context["pending_view_mode"] = view_param
            
            if view_param == "pending":
                # Build preview record from pending snapshot (unsaved copy)
                pending = record.pending_change
                snapshot = pending.snapshot
                
                # Create a copy of the record for display
                preview_record = KarteiRecord()
                preview_record.pk = record.pk
                preview_record.pkid = record.pkid
                preview_record.id = record.id
                preview_record.year = record.year
                preview_record.status = record.status
                preview_record.months_mode = record.months_mode
                # Note: is_sepa is a computed property (from sepa_marker),
                # so it doesn't need explicit copying - sepa_marker is copied below
                
                # Copy non-tracked fields from original record
                for field in record._meta.fields:
                    field_name = field.name
                    if field_name not in ('pk',):
                        try:
                            setattr(preview_record, field_name, getattr(record, field_name))
                        except Exception:
                            pass
                
                # Apply snapshot values to preview record
                from decimal import Decimal
                from apps.karteien.models import TRACKED_FIELDS
                
                changed_fields = set()
                for field_name in TRACKED_FIELDS:
                    if field_name in snapshot:
                        snap_value = snapshot[field_name]
                        orig_value = getattr(record, field_name, None)
                        
                        # Convert snapshot value to proper type
                        if snap_value is not None:
                            if field_name.startswith("month_") or field_name in ("price1", "price2"):
                                if isinstance(snap_value, str) and snap_value:
                                    snap_value = Decimal(snap_value)
                                elif not snap_value:
                                    snap_value = None
                            elif field_name == "birthdate":
                                if isinstance(snap_value, str) and snap_value:
                                    from datetime import datetime as dt
                                    snap_value = dt.strptime(snap_value, "%Y-%m-%d").date()
                        
                        setattr(preview_record, field_name, snap_value)
                        
                        # Check if value changed
                        if not self._values_equal(orig_value, snap_value):
                            changed_fields.add(field_name)
                
                context["pending_record_preview"] = preview_record
                context["pending_changed_fields"] = changed_fields
                # Use preview record as display_record
                context["display_record"] = preview_record
            else:
                # Current view - use original record
                context["display_record"] = record
        else:
            # Default: use the record as-is
            context["display_record"] = record
        
        # -------------------------------------------------------------------------
        # Discounts Summary
        # -------------------------------------------------------------------------
        # Family discounts for this record's year + family_id
        context["family_discounts"] = FamilyDiscount.objects.filter(
            year=record.year,
            family_id=record.family_id,
        ).select_related("discount").order_by("start_month")
        
        # Record-level discounts
        context["record_discounts"] = record.record_discounts.select_related(
            "discount"
        ).order_by("start_month")
        
        # Helper for discounts_disabled_months display
        if record.discounts_disabled and record.discounts_disabled_months:
            months = sorted(record.discounts_disabled_months)
            if len(months) == 1:
                context["disabled_months_display"] = f"Monat {months[0]}"
            elif months == list(range(months[0], months[-1] + 1)):
                context["disabled_months_display"] = f"Monate {months[0]}-{months[-1]}"
            else:
                context["disabled_months_display"] = f"Monate {', '.join(map(str, months))}"
        else:
            context["disabled_months_display"] = None
        
        # -------------------------------------------------------------------------
        # History Parsing for Timeline UI
        # -------------------------------------------------------------------------
        history_entries = []
        raw_history_fallback = None
        
        if record.history_raw:
            from apps.history.services import parse_raw_history
            try:
                history_entries = parse_raw_history(record.history_raw)
                # If parsing returned empty but we have raw data, show raw as fallback
                if not history_entries:
                    raw_history_fallback = record.history_raw
            except Exception:
                # Parsing failed - show raw history as fallback
                raw_history_fallback = record.history_raw
        
        context["history_entries"] = history_entries
        context["raw_history"] = raw_history_fallback
        context["has_history"] = bool(history_entries or record.history_raw)
        
        return context
    
    def _values_equal(self, val1: Any, val2: Any) -> bool:
        """Compare two values for equality, handling None, Decimal, etc."""
        from decimal import Decimal
        
        # Both None
        if val1 is None and val2 is None:
            return True
        # One None, one not
        if val1 is None or val2 is None:
            return False
        # Decimal comparison (avoid floating point issues)
        if isinstance(val1, Decimal) or isinstance(val2, Decimal):
            try:
                return Decimal(str(val1)) == Decimal(str(val2))
            except Exception:
                return val1 == val2
        # String comparison (strip whitespace for safety)
        if isinstance(val1, str) and isinstance(val2, str):
            return val1.strip() == val2.strip()
        return val1 == val2


# =============================================================================
# Create View
# =============================================================================

class KarteiRecordCreateView(KarteiEditorMixin, CreateView):
    """
    Create view for new KarteiRecord entries.
    
    New records are always considered "safe" and don't require approval.
    
    Supports prefill_family_id query param to prefill form fields from
    an existing family record (used by Family Dashboard).
    """
    
    model = KarteiRecord
    form_class = KarteiRecordForm
    template_name = "karteien/record_form.html"
    
    # Store prefill info for context
    _prefill_from_family = False
    _prefill_source_year = None
    _prefill_family_id = None
    
    def get_form_kwargs(self) -> dict[str, Any]:
        """Pass user and year to form."""
        kwargs = super().get_form_kwargs()
        kwargs["user"] = self.request.user
        kwargs["year"] = int(self.request.GET.get("year", date.today().year))
        return kwargs
    
    def get_initial(self) -> dict[str, Any]:
        """
        Get initial form values.
        
        If prefill_family_id is provided in query params, find a reference record
        for this family and prefill contact fields from it.
        """
        initial = super().get_initial()
        
        prefill_family_id = self.request.GET.get("prefill_family_id", "").strip()
        if not prefill_family_id:
            return initial
        
        year = int(self.request.GET.get("year", date.today().year))
        
        # Find a reference record for this family
        # Priority: same year, then fallback to most recent year
        reference_record = (
            KarteiRecord.objects
            .filter(family_id=prefill_family_id, year=year)
            .order_by("-id")
            .first()
        )
        
        if not reference_record:
            # Fallback: find most recent record from any year
            reference_record = (
                KarteiRecord.objects
                .filter(family_id=prefill_family_id)
                .order_by("-year", "-id")
                .first()
            )
        
        if reference_record:
            # Prefill family and contact fields
            initial["family_id"] = reference_record.family_id
            initial["parent_name"] = reference_record.parent_name
            initial["child_name"] = reference_record.child_name
            initial["birthdate"] = reference_record.birthdate
            initial["address"] = reference_record.address
            initial["phone"] = reference_record.phone
            initial["mobile"] = reference_record.mobile
            initial["email"] = reference_record.email
            
            # Store prefill info for context
            self._prefill_from_family = True
            self._prefill_source_year = reference_record.year
            self._prefill_family_id = prefill_family_id
        
        return initial
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add year info to context."""
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["year"] = int(self.request.GET.get("year", date.today().year))
        
        # Prefill from family info
        context["prefill_from_family"] = self._prefill_from_family
        context["prefill_source_year"] = self._prefill_source_year
        context["prefill_family_id"] = self._prefill_family_id
        
        # Collect family prefill options if family_id is provided
        context.update(self._collect_family_prefill_options())
        
        # New records are AUTO mode by default
        context["is_auto_mode"] = True
        context["months_mode"] = MonthsMode.AUTO
        
        # Compute max FamilyID and suggestion for next
        context.update(self._compute_family_id_hints())
        
        # Get hourly months for template
        form = context.get('form')
        if form:
            context["hourly_months"] = form.get_hourly_months()
            
            # Zero clamp confirmation context
            if hasattr(form, '_calculation_flags') and form._calculation_flags:
                flags = form._calculation_flags
                context["needs_zero_clamp_confirmation"] = flags.requires_confirmation
                context["clamped_zero_months"] = list(flags.clamped_to_zero_months) if flags.clamped_to_zero_months else []
            else:
                context["needs_zero_clamp_confirmation"] = False
                context["clamped_zero_months"] = []
        
        return context
    
    def _collect_family_prefill_options(self) -> dict[str, Any]:
        """
        Collect unique non-empty values for family fields from existing records.
        
        Used when prefill_family_id is provided to show dropdown options for fields
        where the family has multiple different values.
        
        Returns:
            dict with family_prefill_options and family_prefill_active
        """
        if not self._prefill_family_id:
            return {
                "family_prefill_options": {},
                "family_prefill_active": False,
            }
        
        # Get all records for this family (across all years for comprehensive options)
        family_records = KarteiRecord.objects.filter(
            family_id=self._prefill_family_id
        ).values(
            "parent_name", "child_name", "birthdate",
            "address", "phone", "mobile", "email"
        )
        
        # Collect unique non-empty values for each field
        options = {
            "parent_name": set(),
            "child_name": set(),
            "birthdate": [],  # Will be list of (value, label) tuples
            "address": set(),
            "phone": set(),
            "mobile": set(),
            "email": set(),
        }
        
        birthdate_set = set()  # For deduplication
        
        for record in family_records:
            for field_name in ["parent_name", "child_name", "address", "phone", "mobile", "email"]:
                value = record.get(field_name)
                if value and str(value).strip():
                    options[field_name].add(str(value).strip())
            
            # Handle birthdate specially (serialize for value, format for label)
            bd = record.get("birthdate")
            if bd:
                bd_value = bd.strftime("%Y-%m-%d") if hasattr(bd, "strftime") else str(bd)
                if bd_value not in birthdate_set:
                    birthdate_set.add(bd_value)
                    bd_label = bd.strftime("%d.%m.%Y") if hasattr(bd, "strftime") else str(bd)
                    options["birthdate"].append({"value": bd_value, "label": bd_label})
        
        # Convert sets to sorted lists
        for field_name in ["parent_name", "child_name", "address", "phone", "mobile", "email"]:
            options[field_name] = sorted(options[field_name])
        
        # Sort birthdates by value
        options["birthdate"] = sorted(options["birthdate"], key=lambda x: x["value"])
        
        return {
            "family_prefill_options": options,
            "family_prefill_active": True,
        }
    
    def _compute_family_id_hints(self) -> dict[str, str]:
        """
        Compute max FamilyID in database and suggest next one.
        
        Uses the centralized FamilyID service that considers both
        existing KarteiRecords and active reservations.
        
        FamilyID format: "N. NNNN" (prefix.number)
        Returns dict with family_id_max_display and family_id_next_suggestion.
        """
        import re
        
        family_ids = (
            KarteiRecord.objects
            .exclude(family_id__isnull=True)
            .exclude(family_id="")
            .values_list("family_id", flat=True)
            .distinct()
        )
        
        # Pattern: optional spaces, digit(s), dot, optional spaces, digit(s), optional spaces
        pattern = re.compile(r'^\s*(\d+)\.\s*(\d+)\s*$')
        
        max_entry = None
        max_prefix = 0
        max_number = 0
        
        for fid in family_ids:
            if not fid:
                continue
            match = pattern.match(fid)
            if match:
                prefix = int(match.group(1))
                number = int(match.group(2))
                # Compare by (prefix, number) tuple
                if (prefix, number) > (max_prefix, max_number):
                    max_prefix = prefix
                    max_number = number
                    max_entry = fid.strip()
        
        # Use centralized service for next suggestion (considers reservations)
        next_suggestion = get_next_family_id(prefix="1")
        
        if max_entry:
            return {
                "family_id_max_display": max_entry,
                "family_id_next_suggestion": next_suggestion,
            }
        else:
            return {
                "family_id_max_display": "(unbekannt)",
                "family_id_next_suggestion": next_suggestion or "",
            }
    
    def form_valid(self, form) -> HttpResponse:
        """Save new record with metadata and AUTO billing."""
        record = form.save(commit=False)
        
        # Set year
        record.year = int(self.request.GET.get("year", date.today().year))
        
        # Generate next ID for this year (max ID + 1)
        max_id = KarteiRecord.objects.filter(year=record.year).order_by('-id').values_list('id', flat=True).first()
        record.id = (max_id or 0) + 1
        
        # Set last change metadata
        user = self.request.user
        record.last_change_role = user.role
        record.last_change_date = date.today()
        record.last_change_time = date.today().strftime("%H:%M")
        
        # Status is NORMAL for new records
        record.status = RecordStatus.NORMAL
        
        # Set months_mode to AUTO for new records
        record.months_mode = MonthsMode.AUTO
        
        # Get billing data from form
        billing_data = form.get_billing_data()
        record.hours_amounts = billing_data['hours_amounts']
        
        # Calculate billing
        flags = recalculate_record_months(
            record,
            hours_amounts=billing_data['hours_amounts'],
        )
        
        # Show warning if percent discount was clamped
        if flags.percent_discount_exceeded:
            messages.warning(
                self.request,
                f"Warnung: Die Summe der Prozentrabatte ({flags.original_percent_sum * 100:.0f}%) "
                f"wurde auf 99% begrenzt."
            )
        
        record.save()
        
        messages.success(self.request, f"Datensatz {record.id} wurde erstellt.")
        
        # TODO: Add to history_raw (requires history module)
        # See Export_HistoryBuilder for format
        
        # If Admin clicked "Erstellen & Rabatt hinzufügen", redirect to record discount create
        if (
            self.request.POST.get("save_add_record_discount")
            and self.request.user.is_admin_role
        ):
            return redirect(f"/catalog/record-discounts/create/?record_pk={record.pkid}")
        
        return redirect("karteien:record_detail", pk=record.pk)
    
    def get_success_url(self) -> str:
        """Return URL to record detail page."""
        return reverse("karteien:record_detail", kwargs={"pk": self.object.pk})


# =============================================================================
# Update View
# =============================================================================

class KarteiRecordUpdateView(KarteiEditorMixin, UpdateView):
    """
    Update view for existing KarteiRecord entries.
    
    Changes are classified as SAFE or RISKY:
    - SAFE: Applied directly to the record (only for NORMAL status)
    - RISKY: Create/update PendingChange, mark record as PENDING
    
    Important: This view only allows editing records with status NORMAL.
    For PENDING/DECLINED records, users must use the approvals workflow:
    - PENDING: Wait for Superadmin decision (approvals UI)
    - DECLINED: Use DeclinedOverview to fix and resubmit
    
    Operator restrictions (SEPA, past-months) are enforced here.
    """
    
    model = KarteiRecord
    form_class = KarteiRecordForm
    template_name = "karteien/record_form.html"
    
    def dispatch(self, request: HttpRequest, *args, **kwargs) -> HttpResponse:
        """Check record status before allowing access to the edit form.
        
        For PENDING/DECLINED records:
        - Admin can edit via standard form (changes go through approvals)
        - Operator is blocked with a message
        """
        self.object = self.get_object()
        user = request.user
        
        # Store original status for form_valid logic
        self._original_status = self.object.status
        
        # For PENDING/DECLINED: allow ADMIN, block Operator
        if self.object.status == RecordStatus.PENDING:
            if not user.is_admin_role:
                messages.error(
                    request,
                    "Datensatz ist PENDING. Änderungen müssen über den "
                    "Freigabe-Workflow erfolgen. Bitte warten Sie auf die "
                    "Entscheidung des Superadmin."
                )
                return redirect("karteien:record_detail", pk=self.object.pk)
            # Admin can proceed to edit PENDING record
        
        if self.object.status == RecordStatus.DECLINED:
            if not user.is_admin_role:
                messages.error(
                    request,
                    "Datensatz ist DECLINED. Bitte nutzen Sie die Ansicht "
                    "\"Abgelehnte Änderungen\", um Korrekturen erneut einzureichen."
                )
                return redirect("karteien:record_detail", pk=self.object.pk)
            # Admin can proceed to edit DECLINED record
            # Store declined_change_id from query param if provided
            self._declined_change_id = request.GET.get("declined_change_id")
        
        return super().dispatch(request, *args, **kwargs)
    
    def _apply_snapshot_to_instance(
        self,
        instance: KarteiRecord,
        snapshot: dict[str, Any],
    ) -> KarteiRecord:
        """
        Apply a snapshot dictionary to a KarteiRecord instance (in-memory only).
        
        Converts string representations back to proper Django field types.
        Does NOT copy FK ref fields (like teacher1_ref_id etc.) — those stay from original.
        
        Args:
            instance: The record to update (in-memory, not saved).
            snapshot: Dictionary of field_name -> value from PendingChange/DeclinedChange.
        
        Returns:
            The updated instance (not saved).
        """
        from apps.karteien.models import TRACKED_FIELDS
        
        for field_name, value in snapshot.items():
            if field_name not in TRACKED_FIELDS:
                continue
            
            if value is not None:
                # DecimalField
                if field_name in ("price1", "price2") or field_name.startswith("month_"):
                    if isinstance(value, str) and value:
                        value = Decimal(value)
                    elif not value:
                        value = None
                # DateField
                elif field_name == "birthdate":
                    if isinstance(value, str) and value:
                        from datetime import datetime as dt
                        value = dt.strptime(value, "%Y-%m-%d").date()
            
            setattr(instance, field_name, value)
        
        return instance
    
    def _get_snapshot_instance(self) -> tuple[KarteiRecord | None, DeclinedChange | None]:
        """
        Get instance with snapshot applied for PENDING/DECLINED records.
        
        For PENDING: applies pending_change.snapshot to a copy of the record.
        For DECLINED: applies declined_change.snapshot (from query param or latest).
        
        Returns:
            Tuple of (snapshot_applied_instance, declined_change_if_applicable).
            Returns (None, None) if not PENDING/DECLINED or no snapshot available.
        """
        record = self.object
        
        if record.status == RecordStatus.PENDING:
            pending = getattr(record, "pending_change", None)
            if pending and not pending.is_processed and pending.snapshot:
                # Create a copy of the record for form initialization
                instance_copy = KarteiRecord()
                # Copy all fields from original
                for field in record._meta.fields:
                    field_name = field.name
                    if field_name != 'pk':
                        try:
                            setattr(instance_copy, field_name, getattr(record, field_name))
                        except Exception:
                            pass
                # Apply snapshot to the copy
                self._apply_snapshot_to_instance(instance_copy, pending.snapshot)
                return (instance_copy, None)
        
        elif record.status == RecordStatus.DECLINED:
            declined = None
            declined_change_id = getattr(self, "_declined_change_id", None)
            
            if declined_change_id:
                try:
                    declined = DeclinedChange.objects.get(
                        pk=int(declined_change_id),
                        record=record,
                    )
                except (DeclinedChange.DoesNotExist, ValueError):
                    declined = None
            
            # Fallback: use the most recent declined change for this record
            if not declined:
                declined = record.declined_changes.order_by("-created_at").first()
            
            if declined and declined.snapshot:
                # Create a copy of the record for form initialization
                instance_copy = KarteiRecord()
                # Copy all fields from original
                for field in record._meta.fields:
                    field_name = field.name
                    if field_name != 'pk':
                        try:
                            setattr(instance_copy, field_name, getattr(record, field_name))
                        except Exception:
                            pass
                # Apply snapshot to the copy
                self._apply_snapshot_to_instance(instance_copy, declined.snapshot)
                return (instance_copy, declined)
        
        return (None, None)
    
    def get_form_kwargs(self) -> dict[str, Any]:
        """Pass user, year, and requires_comment flag to form.
        
        For PENDING/DECLINED records: prefill form with snapshot values.
        """
        kwargs = super().get_form_kwargs()
        kwargs["user"] = self.request.user
        kwargs["year"] = self.object.year
        
        # If refresh_billing=1 is in URL, discounts were changed externally
        # and we require a comment for saving
        if self.request.GET.get("refresh_billing") == "1":
            kwargs["requires_comment"] = True
        
        # For PENDING/DECLINED: require comment and prefill from snapshot
        original_status = getattr(self, "_original_status", self.object.status)
        if original_status in (RecordStatus.PENDING, RecordStatus.DECLINED):
            kwargs["requires_comment"] = True
            
            # Get snapshot-applied instance for prefill
            snapshot_instance, declined_change = self._get_snapshot_instance()
            if snapshot_instance:
                kwargs["instance"] = snapshot_instance
                # Store declined_change for form_valid to delete later
                self._active_declined_change = declined_change
        
        return kwargs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add record info to context.
        
        For PENDING/DECLINED records: use snapshot-applied instance for display.
        """
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["year"] = self.object.year
        
        # For PENDING/DECLINED: use snapshot-applied instance as context["record"]
        original_status = getattr(self, "_original_status", self.object.status)
        if original_status in (RecordStatus.PENDING, RecordStatus.DECLINED):
            snapshot_instance, _ = self._get_snapshot_instance()
            if snapshot_instance:
                context["record"] = snapshot_instance
                # Also update mode info based on snapshot instance
                context["is_auto_mode"] = snapshot_instance.months_mode == MonthsMode.AUTO
                context["months_mode"] = snapshot_instance.months_mode
                context["mismatch_months"] = list(get_month_mismatches(snapshot_instance))
                context["is_override_mode"] = snapshot_instance.months_mode == MonthsMode.OVERRIDE
                context["legacy_base_amounts_enabled"] = getattr(snapshot_instance, 'legacy_base_amounts_enabled', False)
            else:
                context["record"] = self.object
                context["is_auto_mode"] = self.object.months_mode == MonthsMode.AUTO
                context["months_mode"] = self.object.months_mode
                context["mismatch_months"] = list(get_month_mismatches(self.object))
                context["is_override_mode"] = self.object.months_mode == MonthsMode.OVERRIDE
                context["legacy_base_amounts_enabled"] = getattr(self.object, 'legacy_base_amounts_enabled', False)
            
            # Info alert about editing mode
            context["is_pending_edit_mode"] = (original_status == RecordStatus.PENDING)
            context["is_declined_edit_mode"] = (original_status == RecordStatus.DECLINED)
        else:
            context["record"] = self.object
            # AUTO mode info
            context["is_auto_mode"] = self.object.months_mode == MonthsMode.AUTO
            context["months_mode"] = self.object.months_mode
            context["mismatch_months"] = list(get_month_mismatches(self.object))
            context["is_override_mode"] = self.object.months_mode == MonthsMode.OVERRIDE
            context["legacy_base_amounts_enabled"] = getattr(self.object, 'legacy_base_amounts_enabled', False)
        
        # Get hourly months for template
        form = context.get('form')
        if form:
            context["hourly_months"] = form.get_hourly_months()
            
            # Zero clamp confirmation context
            if hasattr(form, '_calculation_flags') and form._calculation_flags:
                flags = form._calculation_flags
                context["needs_zero_clamp_confirmation"] = flags.requires_confirmation
                context["clamped_zero_months"] = list(flags.clamped_to_zero_months) if flags.clamped_to_zero_months else []
            else:
                context["needs_zero_clamp_confirmation"] = False
                context["clamped_zero_months"] = []
        
        # Show restrictions info for Operators
        user = self.request.user
        if user.has_sepa_restrictions:
            context["sepa_warning"] = self.object.is_sepa
        if user.has_past_months_restrictions:
            from .validators import get_allowed_months
            allowed, reason = get_allowed_months(self.object.year)
            context["allowed_months"] = allowed
            context["past_months_reason"] = reason
        
        # Legacy badges context: next URL for catalog quick-links
        context["edit_next_url"] = self.request.get_full_path()
        
        # Find teacher candidates by parsing legacy names
        context.update(self._get_legacy_teacher_candidates())
        
        # -------------------------------------------------------------------------
        # Discounts Summary
        # -------------------------------------------------------------------------
        record = self.object
        # Family discounts for this record's year + family_id
        context["family_discounts"] = FamilyDiscount.objects.filter(
            year=record.year,
            family_id=record.family_id,
        ).select_related("discount").order_by("start_month")
        
        # Record-level discounts
        context["record_discounts"] = record.record_discounts.select_related(
            "discount"
        ).order_by("start_month")
        
        # Helper for discounts_disabled_months display
        if record.discounts_disabled and record.discounts_disabled_months:
            months = sorted(record.discounts_disabled_months)
            if len(months) == 1:
                context["disabled_months_display"] = f"Monat {months[0]}"
            elif months == list(range(months[0], months[-1] + 1)):
                context["disabled_months_display"] = f"Monate {months[0]}-{months[-1]}"
            else:
                context["disabled_months_display"] = f"Monate {', '.join(map(str, months))}"
        else:
            context["disabled_months_display"] = None
        
        return context
    
    def _get_legacy_teacher_candidates(self) -> dict[str, Any]:
        """
        Find teacher IDs from legacy teacher names for catalog quick-links.
        
        Parses legacy name as 'Nachname Vorname' and looks up Teacher model.
        Returns dict with teacher*_candidate_id for use in assignment create links.
        """
        from apps.catalog.models import Teacher
        
        result = {
            "teacher1_candidate_id": None,
            "teacher2_candidate_id": None,
        }
        
        def parse_and_find_teacher(legacy_name: str) -> int | None:
            """Parse legacy name and find teacher ID."""
            if not legacy_name:
                return None
            # Normalize whitespace
            normalized = " ".join(legacy_name.split())
            if not normalized:
                return None
            
            parts = normalized.split()
            if len(parts) < 2:
                # Single word: cannot reliably find teacher
                return None
            
            # Format: "Nachname Vorname" -> last_name = all but last, first_name = last
            first_name = parts[-1]
            last_name = " ".join(parts[:-1])
            
            try:
                teacher = Teacher.objects.get(
                    last_name=last_name,
                    first_name=first_name,
                    is_active=True,
                )
                return teacher.id
            except Teacher.DoesNotExist:
                return None
            except Teacher.MultipleObjectsReturned:
                return None
        
        record = self.object
        if record.teacher1_legacy_name and not record.teacher1_ref_id:
            result["teacher1_candidate_id"] = parse_and_find_teacher(record.teacher1_legacy_name)
        
        if record.teacher2_legacy_name and not record.teacher2_ref_id:
            result["teacher2_candidate_id"] = parse_and_find_teacher(record.teacher2_legacy_name)
        
        return result
    
    def form_valid(self, form) -> HttpResponse:
        """Handle save with safe/risky classification and AUTO billing.
        
        For NORMAL status: SAFE changes apply directly, RISKY go to pending.
        For PENDING status (Admin edit): update PendingChange, keep status PENDING.
        For DECLINED status (Admin edit): create PendingChange, set status PENDING,
            delete the DeclinedChange being edited.
        
        LEGACY mode handling:
        - If no meaningful changes, months_mode stays LEGACY, no recalculation
        - If meaningful changes detected, convert to AUTO and recalculate touched months
        """
        record = form.save(commit=False)
        original = KarteiRecord.objects.get(pk=self.object.pk)
        user = self.request.user
        original_status = getattr(self, "_original_status", original.status)
        
        # Get billing data from form
        billing_data = form.get_billing_data()
        record.hours_amounts = billing_data['hours_amounts']
        
        # Handle billing calculations based on mode
        if original.months_mode == MonthsMode.LEGACY:
            # LEGACY mode: check for meaningful changes
            should_convert = billing_data.get('should_convert_to_auto', False)
            touched_months = billing_data.get('touched_months', set())
            
            if should_convert and touched_months:
                # Convert LEGACY → AUTO with partial update
                from .billing import recalculate_legacy_to_auto
                
                flags = recalculate_legacy_to_auto(
                    record,
                    touched_months=touched_months,
                    hours_amounts=billing_data['hours_amounts'],
                )
                
                # Show warning if percent discount was clamped
                if flags.percent_discount_exceeded:
                    messages.warning(
                        self.request,
                        f"Warnung: Die Summe der Prozentrabatte ({flags.original_percent_sum * 100:.0f}%) "
                        f"wurde auf 99% begrenzt."
                    )
                
                messages.info(
                    self.request,
                    f"Datensatz wurde von LEGACY auf AUTO umgestellt. "
                    f"Aktualisierte Monate: {', '.join(str(m) for m in sorted(touched_months))}."
                )
            # else: no meaningful changes, keep LEGACY mode and original month values
            
        elif record.months_mode == MonthsMode.AUTO:
            # AUTO mode: check for meaningful billing changes before recalculating
            has_billing_changes, _touched = detect_meaningful_changes(
                original,
                form.cleaned_data,
                billing_data['hours_amounts'],
            )
            
            if has_billing_changes:
                # Meaningful changes detected: recalculate months
                flags = recalculate_record_months(
                    record,
                    apply_from_month_1=billing_data['apply_from_month_1'],
                    apply_from_month_2=billing_data['apply_from_month_2'],
                    apply_to_month_1=billing_data.get('apply_to_month_1'),
                    apply_to_month_2=billing_data.get('apply_to_month_2'),
                    hours_amounts=billing_data['hours_amounts'],
                )
                
                # Show warning if percent discount was clamped
                if flags.percent_discount_exceeded:
                    messages.warning(
                        self.request,
                        f"Warnung: Die Summe der Prozentrabatte ({flags.original_percent_sum * 100:.0f}%) "
                        f"wurde auf 99% begrenzt."
                    )
            else:
                # No billing changes: preserve existing price history (base_amounts and month values)
                record.base_amounts = original.base_amounts
                for month_num in range(1, 13):
                    field_name = f"month_{month_num}"
                    setattr(record, field_name, getattr(original, field_name))
        
        # Apply termination zeroing if contract status becomes terminated
        if billing_data.get('needs_termination_zeroing'):
            termination_month = billing_data.get('termination_effective_month')
            if termination_month:
                zeroed_months = []
                for m in range(termination_month, 13):
                    field_name = f"month_{m}"
                    setattr(record, field_name, Decimal('0.00'))
                    zeroed_months.append(m)
                
                if zeroed_months:
                    messages.info(
                        self.request,
                        f"Vertrag gekündigt ab Monat {termination_month}: "
                        f"Monate {', '.join(str(m) for m in zeroed_months)} wurden auf 0 € gesetzt."
                    )
        
        # Get proposed changes (comparing form data with original)
        proposed_data = form.cleaned_data
        
        # Apply Operator filters if needed
        if user.has_sepa_restrictions or user.has_past_months_restrictions:
            filtered_changes, warnings = apply_operator_filters(
                original, proposed_data, user
            )
            for warning in warnings:
                messages.warning(self.request, warning)
            
            # Apply only allowed changes
            for field, value in filtered_changes.items():
                setattr(record, field, value)
        
        # ==================================================================
        # Handle PENDING status (Admin editing a pending record)
        # ==================================================================
        if original_status == RecordStatus.PENDING:
            return self._handle_pending_save(record, original, form, user)
        
        # ==================================================================
        # Handle DECLINED status (Admin editing a declined record)
        # ==================================================================
        if original_status == RecordStatus.DECLINED:
            return self._handle_declined_save(record, original, form, user)
        
        # ==================================================================
        # Handle NORMAL status (standard workflow)
        # ==================================================================
        # Classify the change
        classification = classify_change(record, original)
        
        if classification == "SAFE":
            # Safe change: apply directly, including safe non-tracked fields
            record.last_change_role = user.role
            record.last_change_date = date.today()
            record.last_change_time = date.today().strftime("%H:%M")
            record.save()
            
            # Write admin comment to history_raw if provided
            admin_comment = form.get_comment().strip()
            if admin_comment:
                write_history_entry(record, "ADM", user, admin_comment)
            
            messages.success(
                self.request, 
                f"Datensatz {record.id} wurde aktualisiert."
            )
            
        else:
            # Risky change: create pending with admin comment
            admin_comment = form.get_comment()
            pending = create_or_update_pending_change(
                record,
                admin_comment=admin_comment,
            )
            
            # Save safe (non-tracked) fields directly to DB
            # so Lehrer/subject refs, discounts settings etc. are not lost
            safe_updates = self._build_safe_fields_update(record)
            safe_updates["status"] = RecordStatus.PENDING
            KarteiRecord.objects.filter(pk=record.pk).update(**safe_updates)
            
            messages.info(
                self.request,
                f"Änderungen für Datensatz {record.id} wurden zur Genehmigung eingereicht. "
                "Ein Superadmin muss die Änderungen prüfen."
            )
        
        return redirect("karteien:record_detail", pk=record.pk)
    
    def _build_safe_fields_update(self, record: KarteiRecord) -> dict[str, Any]:
        """
        Build a dict of safe (non-tracked) fields to update in DB.
        
        These fields are NOT tracked by approvals and should be persisted
        immediately so the user's Lehrer/Fach selection, discount settings,
        and contract info are not lost when a risky change goes to pending.
        """
        return {
            # Subject/Teacher/Price references (catalog FKs)
            "subject1_ref_id": record.subject1_ref_id,
            "teacher1_ref_id": record.teacher1_ref_id,
            "price1_ref_id": record.price1_ref_id,
            "start_month_1": record.start_month_1,
            "end_month_1": record.end_month_1,
            "months_csv_1": record.months_csv_1,
            "subject2_ref_id": record.subject2_ref_id,
            "teacher2_ref_id": record.teacher2_ref_id,
            "price2_ref_id": record.price2_ref_id,
            "start_month_2": record.start_month_2,
            "end_month_2": record.end_month_2,
            "months_csv_2": record.months_csv_2,
            # SEPA marker
            "sepa_marker": record.sepa_marker,
            # Months mode and billing helpers
            "months_mode": record.months_mode,
            "base_amounts": record.base_amounts,
            "hours_amounts": record.hours_amounts,
            # LEGACY recalc marker (must be persisted to track LEGACY→AUTO transitions)
            "legacy_base_amounts_enabled": record.legacy_base_amounts_enabled,
            # Discount settings
            "discounts_disabled": record.discounts_disabled,
            "discounts_disabled_months": record.discounts_disabled_months,
            # Contract fields
            "is_monthly_contract": record.is_monthly_contract,
            "contract_type_raw": record.contract_type_raw,
            "is_contract_terminated": record.is_contract_terminated,
            "contract_status_raw": record.contract_status_raw,
            "contract_terminated_from_month": record.contract_terminated_from_month,
        }
    
    def _handle_pending_save(
        self,
        record: KarteiRecord,
        original: KarteiRecord,
        form,
        user,
    ) -> HttpResponse:
        """
        Handle saving when editing a PENDING record (Admin only).
        
        Updates the existing PendingChange with new snapshot.
        Safe (non-tracked) fields are saved directly to DB.
        Status remains PENDING.
        """
        admin_comment = form.get_comment()
        
        # Create/update pending change with new proposed values
        pending = create_or_update_pending_change(
            record,
            admin_comment=admin_comment,
        )
        
        # Save safe (non-tracked) fields directly to DB
        safe_updates = self._build_safe_fields_update(record)
        # Keep status as PENDING
        safe_updates["status"] = RecordStatus.PENDING
        KarteiRecord.objects.filter(pk=original.pk).update(**safe_updates)
        
        messages.info(
            self.request,
            f"Wartende Änderungen für Datensatz {record.id} wurden aktualisiert. "
            "Die Änderungen warten weiterhin auf Genehmigung durch einen Superadmin."
        )
        
        return redirect("karteien:record_detail", pk=original.pk)
    
    def _handle_declined_save(
        self,
        record: KarteiRecord,
        original: KarteiRecord,
        form,
        user,
    ) -> HttpResponse:
        """
        Handle saving when editing a DECLINED record (Admin only).
        
        Creates a new PendingChange (or updates existing if any).
        Sets status to PENDING.
        Deletes the DeclinedChange that was being edited (if identified).
        """
        admin_comment = form.get_comment()
        
        # Create/update pending change with corrected values
        pending = create_or_update_pending_change(
            record,
            admin_comment=admin_comment,
        )
        
        # Save safe (non-tracked) fields and set status to PENDING
        safe_updates = self._build_safe_fields_update(record)
        safe_updates["status"] = RecordStatus.PENDING
        KarteiRecord.objects.filter(pk=original.pk).update(**safe_updates)
        
        # Delete the DeclinedChange that was being edited
        declined_to_delete = getattr(self, "_active_declined_change", None)
        if declined_to_delete:
            declined_to_delete.delete()
            messages.success(
                self.request,
                f"Korrektur für Datensatz {record.id} wurde eingereicht. "
                "Der abgelehnte Eintrag wurde entfernt und die Änderungen warten "
                "auf Genehmigung durch einen Superadmin."
            )
        else:
            messages.info(
                self.request,
                f"Korrektur für Datensatz {record.id} wurde eingereicht. "
                "Die Änderungen warten auf Genehmigung durch einen Superadmin."
            )
        
        return redirect("karteien:record_detail", pk=original.pk)
    
    def get_success_url(self) -> str:
        """Return URL to record detail page."""
        return reverse("karteien:record_detail", kwargs={"pk": self.object.pk})


# =============================================================================
# Delete View
# =============================================================================

class KarteiRecordDeleteView(KarteiEditorMixin, DeleteView):
    """
    Delete view for KarteiRecord entries.
    
    Only Admin can delete records. Deletion is permanent.
    """
    
    model = KarteiRecord
    template_name = "karteien/record_confirm_delete.html"
    success_url = reverse_lazy("karteien:record_list")
    
    def test_func(self) -> bool:
        """Only Admin can delete records."""
        user = self.request.user
        return user.is_authenticated and user.is_admin_role
    
    def delete(self, request, *args, **kwargs) -> HttpResponse:
        """Perform deletion with logging."""
        record = self.get_object()
        record_id = record.id
        
        response = super().delete(request, *args, **kwargs)
        
        messages.success(request, f"Datensatz {record_id} wurde gelöscht.")
        
        return response


# =============================================================================
# Months Override View (Emergency Admin-only tool)
# =============================================================================

class MonthsOverrideView(KarteiEditorMixin, View):
    """
    Emergency override view for manually setting month values.
    
    Admin-only tool for correcting month values when automatic
    calculation doesn't apply or needs manual correction.
    
    Changes go through the approvals workflow as RISKY changes.
    """
    
    template_name = "karteien/months_override.html"
    
    def test_func(self) -> bool:
        """Only Admin can use override tool."""
        user = self.request.user
        return user.is_authenticated and user.is_admin_role
    
    def get(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Show override form."""
        record = get_object_or_404(KarteiRecord, pk=pk)
        form = MonthsOverrideForm(record=record)
        
        return render(request, self.template_name, {
            "record": record,
            "form": form,
        })
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Process override form."""
        record = get_object_or_404(KarteiRecord, pk=pk)
        form = MonthsOverrideForm(request.POST, record=record)
        
        if not form.is_valid():
            return render(request, self.template_name, {
                "record": record,
                "form": form,
            })
        
        # Get month changes
        month_changes = form.get_month_changes()
        reason = form.cleaned_data['reason']
        
        # Apply changes to record
        for field_name, value in month_changes.items():
            setattr(record, field_name, value)
        
        # Set mode to OVERRIDE
        record.months_mode = MonthsMode.OVERRIDE
        
        # Create pending change (always risky for override)
        # Pass admin_comment with OVERRIDE prefix so Superadmin can see the reason
        admin_comment = f"OVERRIDE: {reason}"
        pending = create_or_update_pending_change(record, admin_comment=admin_comment)
        
        # Mark record as PENDING
        KarteiRecord.objects.filter(pk=record.pk).update(
            status=RecordStatus.PENDING
        )
        
        messages.info(
            request,
            f"Override für Datensatz {record.id} wurde zur Genehmigung eingereicht. "
            "Ein Superadmin muss die Änderungen prüfen."
        )
        
        return redirect("karteien:record_detail", pk=record.pk)
