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
    classify_change,
    create_or_update_pending_change,
    create_or_update_pending_change_from_snapshot,
    get_changed_tracked_fields,
    write_history_entry,
)
from apps.approvals.models import DeclinedChange, PendingChange
from apps.catalog.models import FamilyDiscount, RecordDiscount

from .billing import recalculate_record_months, get_month_mismatches, detect_meaningful_changes, get_semester_month_ranges
from apps.catalog.pricing import determine_pricing_source, get_suggested_prices_for_record
from apps.catalog.warnings import (
    get_record_price_mismatches,
    get_record_pricing_warnings,
)
from .forms import (
    KarteiRecordForm,
    KarteiRecordFilterForm,
    MonthsOverrideForm,
)
from .models import (
    KarteiRecord, RecordStatus, MonthsMode, ContractStatusKind, ContractStatusEntry,
    ContractTypeEntry, TRACKED_FIELDS, get_contract_type_for_month,
    get_contract_status_for_month,
)
from .services.billing_pipeline import (
    build_apply_category_proposal,
    build_contract_status_proposal,
    build_contract_type_proposal,
    build_months_override_proposal,
    build_quick_set_subject_ref_proposal,
)
from .services.pending_snapshot import (
    build_projected_record_from_snapshot,
    changed_fields_between_records,
)
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
                pending = record.pending_change
                projection = build_projected_record_from_snapshot(
                    record,
                    pending.snapshot,
                )
                preview_record = projection.record
                changed_fields = changed_fields_between_records(
                    record,
                    preview_record,
                    TRACKED_FIELDS,
                )

                context["pending_record_preview"] = preview_record
                context["pending_changed_fields"] = changed_fields
                context["display_record"] = preview_record
                context["pending_timeline_context"] = projection.timeline
                context["pending_meta"] = projection.pending_meta
                context["is_sepa_record"] = preview_record.is_sepa
                context["mismatch_months"] = list(get_month_mismatches(preview_record))
                context["is_override_mode"] = preview_record.months_mode == MonthsMode.OVERRIDE
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
        discount_display_record = context.get("display_record", record)
        if (
            discount_display_record.discounts_disabled
            and discount_display_record.discounts_disabled_months
        ):
            months = sorted(discount_display_record.discounts_disabled_months)
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

        # -----------------------------------------------------------------
        # Pricing info & price mismatches  (PROMPT 147.1)
        # -----------------------------------------------------------------
        display_rec = context.get("display_record", record)
        sem1_months, sem2_months = get_semester_month_ranges(record.year)
        all_suggested = get_suggested_prices_for_record(display_rec)

        pricing_info = {}
        for _sem, _sem_months in ((1, sem1_months), (2, sem2_months)):
            _source = determine_pricing_source(display_rec, _sem)
            _suggested = {m: all_suggested[m] for m in _sem_months}
            _cat_name = None
            _formula = None
            for _m in sorted(_suggested):
                if _suggested[_m] is not None:
                    _cat_name = _suggested[_m].get("category_name")
                    _formula = _suggested[_m]
                    break
            pricing_info[_sem] = {
                "source_value": _source.value,
                "suggested_prices": _suggested,
                "category_name": _cat_name,
                "formula_example": _formula,
            }
        context["pricing_info"] = pricing_info

        raw_mismatches = get_record_price_mismatches(display_rec)
        price_mismatches: dict[int, dict] = {}
        for month, mismatch in raw_mismatches.items():
            current_base = mismatch["stored"]
            if current_base is None:
                current_base = Decimal("0")
            price_mismatches[month] = {
                "current": current_base,
                "suggested": mismatch["suggested"],
            }
        context["price_mismatches"] = price_mismatches
        context["price_mismatch_count"] = len(price_mismatches)

        # -----------------------------------------------------------------
        # Contract Type / Status Timeline  (PROMPT 148.1)
        # -----------------------------------------------------------------
        contract_type_entries = list(
            record.contract_type_entries.all().select_related("changed_by")
        )
        contract_status_entries = list(
            record.contract_status_entries.all().select_related("changed_by")
        )
        pending_timeline_context = context.get("pending_timeline_context")
        if pending_timeline_context:
            pending_type = pending_timeline_context.contract_type_entry
            if pending_type:
                contract_type_entries.append(
                    ContractTypeEntry(
                        record=record,
                        effective_from_month=int(pending_type["effective_from_month"]),
                        is_monthly=bool(pending_type.get("is_monthly", False)),
                        comment=pending_type.get("comment", ""),
                    )
                )
            pending_status = pending_timeline_context.contract_status_entry
            if pending_status:
                contract_status_entries.append(
                    ContractStatusEntry(
                        record=record,
                        effective_from_month=int(pending_status["effective_from_month"]),
                        kind=pending_status["kind"],
                        comment=pending_status.get("comment", ""),
                    )
                )

        STATUS_DISPLAY = {
            ContractStatusKind.ACTIVE: ("✓", "success", "Aktiv"),
            ContractStatusKind.PAUSED: ("⏸", "warning", "Pausiert"),
            ContractStatusKind.TERMINATED: ("✗", "danger", "Gekündigt"),
        }

        contract_timeline = []
        for m in range(1, 13):
            is_monthly = get_contract_type_for_month(
                display_rec, m, entries=contract_type_entries,
            )
            status_kind = get_contract_status_for_month(
                display_rec, m, entries=contract_status_entries,
            )
            symbol, color, label = STATUS_DISPLAY.get(
                status_kind, ("?", "secondary", status_kind),
            )
            contract_timeline.append({
                "month": m,
                "type_label": "M" if is_monthly else "J",
                "type_long": "Monatsvertrag" if is_monthly else "Jahresvertrag",
                "status_kind": status_kind,
                "status_symbol": symbol,
                "status_color": color,
                "status_label": label,
            })

        context["contract_timeline"] = contract_timeline
        context["contract_type_entries"] = contract_type_entries
        context["contract_status_entries"] = contract_status_entries

        # -----------------------------------------------------------------
        # Pricing warnings  (PROMPT 149.2)
        # -----------------------------------------------------------------
        context["pricing_warnings"] = get_record_pricing_warnings(display_rec)

        # -----------------------------------------------------------------
        # Quick-set subject_ref candidates  (PROMPT 149.3)
        # -----------------------------------------------------------------
        context["subject_ref_candidates"] = _find_subject_ref_candidates(record)

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
        sem1, sem2 = get_semester_month_ranges(context["year"])
        context["sem1_first"] = min(sem1)
        context["sem1_last"] = max(sem1)
        context["sem2_first"] = min(sem2)
        context["sem2_last"] = max(sem2)
        
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
    
    Current status behavior:
    - NORMAL: Admin/Operator can use the standard editor, with role filters.
    - PENDING: Admin can edit through this standard editor; the existing
      PendingChange snapshot is updated. Operator is blocked.
    - DECLINED: Admin can edit through this standard editor; saving creates
      a new PendingChange and moves the record back to PENDING. Operator is blocked.
    
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
        
        # Current behavior: Admin may use the standard editor for
        # PENDING/DECLINED; Operator is blocked.
        if self.object.status == RecordStatus.PENDING:
            if not user.is_admin_role:
                messages.error(
                    request,
                    "Datensatz ist PENDING. Änderungen müssen über den "
                    "Freigabe-Workflow erfolgen. Bitte warten Sie auf die "
                    "Entscheidung des Superadmin."
                )
                return redirect("karteien:record_detail", pk=self.object.pk)
            # Admin can proceed to edit PENDING through the standard editor.
        
        if self.object.status == RecordStatus.DECLINED:
            if not user.is_admin_role:
                messages.error(
                    request,
                    "Datensatz ist DECLINED. Bitte nutzen Sie die Ansicht "
                    "\"Abgelehnte Änderungen\", um Korrekturen erneut einzureichen."
                )
                return redirect("karteien:record_detail", pk=self.object.pk)
            # Admin can proceed to edit DECLINED through the standard editor.
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
        return build_projected_record_from_snapshot(instance, snapshot).record
    
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
                instance_copy = self._apply_snapshot_to_instance(
                    instance_copy,
                    pending.snapshot,
                )
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
                instance_copy = self._apply_snapshot_to_instance(
                    instance_copy,
                    declined.snapshot,
                )
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
        sem1, sem2 = get_semester_month_ranges(context["year"])
        context["sem1_first"] = min(sem1)
        context["sem1_last"] = max(sem1)
        context["sem2_first"] = min(sem2)
        context["sem2_last"] = max(sem2)
        
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

        # -----------------------------------------------------------------
        # Pricing info & price mismatches  (PROMPT 147.1)
        # -----------------------------------------------------------------
        ctx_record = context.get("record", record)
        sem1_months, sem2_months = get_semester_month_ranges(record.year)
        all_suggested = get_suggested_prices_for_record(ctx_record)

        pricing_info = {}
        for _sem, _sem_months in ((1, sem1_months), (2, sem2_months)):
            _source = determine_pricing_source(ctx_record, _sem)
            _suggested = {m: all_suggested[m] for m in _sem_months}
            _cat_name = None
            _formula = None
            for _m in sorted(_suggested):
                if _suggested[_m] is not None:
                    _cat_name = _suggested[_m].get("category_name")
                    _formula = _suggested[_m]
                    break
            pricing_info[_sem] = {
                "source_value": _source.value,
                "suggested_prices": _suggested,
                "category_name": _cat_name,
                "formula_example": _formula,
            }
        context["pricing_info"] = pricing_info

        raw_mismatches = get_record_price_mismatches(ctx_record)
        price_mismatches: dict[int, dict] = {}
        for month, mismatch in raw_mismatches.items():
            current_base = mismatch["stored"]
            if current_base is None:
                current_base = Decimal("0")
            price_mismatches[month] = {
                "current": current_base,
                "suggested": mismatch["suggested"],
            }
        context["price_mismatches"] = price_mismatches
        context["price_mismatch_count"] = len(price_mismatches)

        # -----------------------------------------------------------------
        # Pricing warnings  (PROMPT 149.2)
        # -----------------------------------------------------------------
        context["pricing_warnings"] = get_record_pricing_warnings(ctx_record)

        # -----------------------------------------------------------------
        # Quick-set subject_ref candidates  (PROMPT 149.3)
        # -----------------------------------------------------------------
        context["subject_ref_candidates"] = _find_subject_ref_candidates(record)

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
    
    def _get_blocked_months(self, record: KarteiRecord) -> dict[int, str]:
        """Return {month_num: status_kind} for PAUSED/TERMINATED months."""
        from .models import get_contract_status_for_month

        if record.pk and hasattr(record, 'contract_status_entries'):
            entries: list | None = list(record.contract_status_entries.all())
        else:
            entries = None

        blocked: dict[int, str] = {}
        for m in range(1, 13):
            kind = get_contract_status_for_month(record, m, entries=entries)
            if kind in ('PAUSED', 'TERMINATED'):
                blocked[m] = kind
        return blocked

    def _build_month_infos(
        self, blocked_months: dict[int, str],
    ) -> list[dict[str, object]]:
        """Build a list of per-month metadata dicts for the template."""
        infos = []
        STATUS_LABELS = {
            'PAUSED': 'pausiert',
            'TERMINATED': 'gekündigt',
        }
        for m in range(1, 13):
            status = blocked_months.get(m)
            infos.append({
                'num': m,
                'field_name': f'month_{m}',
                'blocked': status is not None,
                'status': status,
                'status_label': STATUS_LABELS.get(status, ''),
            })
        return infos

    def get(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Show override form."""
        record = get_object_or_404(KarteiRecord, pk=pk)
        form = MonthsOverrideForm(record=record)
        blocked_months = self._get_blocked_months(record)
        
        return render(request, self.template_name, {
            "record": record,
            "form": form,
            "blocked_months": blocked_months,
            "month_infos": self._build_month_infos(blocked_months),
        })
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Process override form."""
        record = get_object_or_404(KarteiRecord, pk=pk)
        form = MonthsOverrideForm(request.POST, record=record)
        blocked_months = self._get_blocked_months(record)
        
        if not form.is_valid():
            return render(request, self.template_name, {
                "record": record,
                "form": form,
                "blocked_months": blocked_months,
                "month_infos": self._build_month_infos(blocked_months),
            })
        
        # Get month changes
        month_changes = form.get_month_changes()
        reason = form.cleaned_data['reason']

        # Build pending proposal only. The live row keeps its approved month
        # values and mode until Superadmin approval.
        admin_comment = f"OVERRIDE: {reason}"
        proposal = build_months_override_proposal(
            record,
            month_changes=month_changes,
            blocked_months=blocked_months,
        )
        create_or_update_pending_change_from_snapshot(
            record,
            snapshot=proposal.snapshot,
            admin_comment=admin_comment,
        )
        
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


# =============================================================================
# Apply Category Price View
# =============================================================================

class ApplyCategoryPriceView(KarteiEditorMixin, View):
    """
    POST-only endpoint that applies category-based suggested prices to a
    KarteiRecord and creates a PendingChange for approval.

    The proposal is fully frozen in snapshot v2. The live row keeps its
    approved ``base_amounts``, ``months_mode``, and ``month_*`` values until
    Superadmin approval.
    """

    http_method_names = ["post"]

    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        record = get_object_or_404(KarteiRecord, pk=pk)
        detail_url = reverse("karteien:record_detail", kwargs={"pk": pk})

        # ------------------------------------------------------------------
        # Guard: status must be NORMAL
        # ------------------------------------------------------------------
        if record.status != RecordStatus.NORMAL:
            messages.error(
                request,
                "Kategoriepreis kann nur auf Datensätze im Status NORMAL "
                "angewendet werden.",
            )
            return redirect(detail_url)

        # ------------------------------------------------------------------
        # Guard: months_mode must not be LEGACY
        # ------------------------------------------------------------------
        if record.months_mode == MonthsMode.LEGACY:
            messages.error(
                request,
                "Bitte zuerst auf AUTO umstellen (Monate neu berechnen).",
            )
            return redirect(detail_url)

        # ------------------------------------------------------------------
        # Parse payload
        # ------------------------------------------------------------------
        try:
            semester = int(request.POST.get("semester", 0))
            from_month = int(request.POST.get("from_month", 0))
        except (ValueError, TypeError):
            messages.error(request, "Ungültige Parameter.")
            return redirect(detail_url)

        if semester not in (1, 2) or not 1 <= from_month <= 12:
            messages.error(request, "Ungültige Semester-/Monatsangabe.")
            return redirect(detail_url)

        comment = (request.POST.get("comment") or "").strip()
        if not comment:
            messages.error(
                request,
                "Ein Kommentar ist erforderlich.",
            )
            return redirect(detail_url)

        try:
            proposal = build_apply_category_proposal(
                record, semester=semester, from_month=from_month,
            )
        except ValueError as exc:
            messages.error(request, str(exc))
            return redirect(detail_url)

        months_updated = list(proposal.touched_months)

        # ------------------------------------------------------------------
        # Operator past-month restrictions
        # ------------------------------------------------------------------
        user = request.user
        if user.has_past_months_restrictions:
            from .validators import get_allowed_months

            allowed_fields, _reason = get_allowed_months(record.year)
            blocked = [
                m for m in months_updated if f"month_{m}" not in allowed_fields
            ]
            if blocked:
                months_str = ", ".join(str(m) for m in blocked)
                messages.error(
                    request,
                    f"Sie dürfen vergangene Monate nicht ändern "
                    f"(Monat(e) {months_str}).",
                )
                return redirect(detail_url)

        create_or_update_pending_change_from_snapshot(
            record,
            snapshot=proposal.snapshot,
            admin_comment=comment,
        )

        # Persist only approval status. Billing payload lives in the snapshot.
        KarteiRecord.objects.filter(pk=record.pk).update(status=RecordStatus.PENDING)

        messages.info(
            request,
            "Änderungen wurden zur Genehmigung eingereicht. "
            "Ein Superadmin muss die Änderungen prüfen.",
        )
        return redirect(detail_url)


# =============================================================================
# Apply Category Price — Preview (JSON, no side-effects)
# =============================================================================

from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.contrib.auth.decorators import login_required


@login_required
@require_GET
def apply_price_preview(request: HttpRequest, pk: int) -> JsonResponse:
    """
    Return a JSON preview of the APPLY_CATEGORY pipeline proposal without
    persisting anything.

    Query params:
        semester   – 1 or 2
        from_month – 1-12
    """
    user = request.user
    if not user.can_edit_kartei:
        return JsonResponse({"error": "Keine Berechtigung."}, status=403)

    record = get_object_or_404(KarteiRecord, pk=pk)

    # --- parse params ---
    try:
        semester = int(request.GET.get("semester", 0))
        from_month = int(request.GET.get("from_month", 0))
    except (ValueError, TypeError):
        return JsonResponse({"error": "Ungültige Parameter."}, status=400)

    if semester not in (1, 2) or not 1 <= from_month <= 12:
        return JsonResponse({"error": "Ungültige Semester-/Monatsangabe."}, status=400)

    # --- status / mode guards ---
    if record.status != RecordStatus.NORMAL:
        return JsonResponse(
            {"error": "Aktion nicht verfügbar (Status: " + record.get_status_display() + ")."},
            status=409,
        )
    if record.months_mode == MonthsMode.LEGACY:
        return JsonResponse(
            {"error": "Bitte zuerst auf AUTO umstellen (Monate neu berechnen)."},
            status=409,
        )

    try:
        proposal = build_apply_category_proposal(
            record, semester=semester, from_month=from_month,
        )
        diff = proposal.diff
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)

    # --- serialise Decimals → strings ---
    def _dec_dict(d: dict) -> dict:
        return {k: str(v) if v is not None else None for k, v in d.items()}

    return JsonResponse({
        "months_updated": diff["months_updated"],
        "old_bases": _dec_dict(diff["old_bases"]),
        "new_bases": _dec_dict(diff["new_bases"]),
        "old_months": _dec_dict(diff["old_months"]),
        "new_months": _dec_dict(diff["new_months"]),
    })


# =============================================================================
# Contract Type / Status Change Views  (PROMPT 148.2)
# =============================================================================

class ContractTypeChangeView(KarteiEditorMixin, View):
    """
    POST-only view to propose a ContractType change via PendingChange.

    The actual ContractTypeEntry is NOT created here — it will be created
    only upon APPROVE in PROMPT_148.3.
    """

    http_method_names = ["post"]

    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        record = get_object_or_404(KarteiRecord, pk=pk)

        # --- status guard ---
        if record.status != RecordStatus.NORMAL:
            messages.error(
                request,
                "Aktion nicht verfügbar \u2013 der Datensatz hat den Status "
                f'\u201e{record.get_status_display()}\u201c.',
            )
            return redirect("karteien:record_detail", pk=pk)

        # --- parse & validate form data ---
        try:
            effective_from_month = int(request.POST.get("effective_from_month", 0))
        except (ValueError, TypeError):
            effective_from_month = 0
        if not 1 <= effective_from_month <= 12:
            messages.error(request, "Ung\u00fcltiger Monat (1\u201312).")
            return redirect("karteien:record_detail", pk=pk)

        is_monthly_raw = request.POST.get("is_monthly", "")
        if is_monthly_raw not in ("0", "1"):
            messages.error(request, "Bitte Vertragstyp auswählen.")
            return redirect("karteien:record_detail", pk=pk)
        is_monthly = is_monthly_raw == "1"

        comment = (request.POST.get("comment") or "").strip()
        if not comment:
            messages.error(request, "Bitte einen Kommentar eingeben.")
            return redirect("karteien:record_detail", pk=pk)

        # --- operator past-month restriction ---
        user = request.user
        if user.has_past_months_restrictions:
            from .validators import get_allowed_months
            allowed, _reason = get_allowed_months(record.year)
            if f"month_{effective_from_month}" not in allowed:
                messages.error(
                    request,
                    f"Sie dürfen Monat {effective_from_month} nicht mehr ändern "
                    "(vergangener Monat).",
                )
                return redirect("karteien:record_detail", pk=pk)

        type_label = "Monatsvertrag" if is_monthly else "Jahresvertrag"
        proposal = build_contract_type_proposal(
            record,
            effective_from_month=effective_from_month,
            is_monthly=is_monthly,
            comment=comment,
            user=user,
        )

        create_or_update_pending_change_from_snapshot(
            record, snapshot=proposal.snapshot, admin_comment=comment,
        )

        # Set status = PENDING (safe DB update)
        KarteiRecord.objects.filter(pk=pk).update(status=RecordStatus.PENDING)

        messages.info(
            request,
            f'Vertragstyp-\u00c4nderung \u201e{type_label} ab Monat {effective_from_month}\u201c '
            "wurde zur Genehmigung eingereicht.",
        )
        return redirect("karteien:record_detail", pk=pk)


class ContractStatusChangeView(KarteiEditorMixin, View):
    """
    POST-only view to propose a ContractStatus change via PendingChange.

    If the new status is PAUSED or TERMINATED the affected months are zeroed
    in the snapshot. If ACTIVE and months_mode == AUTO the pipeline restores
    affected months from frozen base_amounts.

    The actual ContractStatusEntry is NOT created here — PROMPT_148.3.
    """

    http_method_names = ["post"]

    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        record = get_object_or_404(KarteiRecord, pk=pk)

        # --- status guard ---
        if record.status != RecordStatus.NORMAL:
            status_display = record.get_status_display()
            messages.error(
                request,
                f"Aktion nicht verfügbar \u2013 der Datensatz hat den Status {status_display}.",
            )
            return redirect("karteien:record_detail", pk=pk)

        # --- parse & validate ---
        try:
            effective_from_month = int(request.POST.get("effective_from_month", 0))
        except (ValueError, TypeError):
            effective_from_month = 0
        if not 1 <= effective_from_month <= 12:
            messages.error(request, "Ungültiger Monat (1–12).")
            return redirect("karteien:record_detail", pk=pk)

        kind = (request.POST.get("kind") or "").strip()
        if kind not in (
            ContractStatusKind.ACTIVE,
            ContractStatusKind.PAUSED,
            ContractStatusKind.TERMINATED,
        ):
            messages.error(request, "Bitte einen gültigen Status auswählen.")
            return redirect("karteien:record_detail", pk=pk)

        comment = (request.POST.get("comment") or "").strip()
        if not comment:
            messages.error(request, "Bitte einen Kommentar eingeben.")
            return redirect("karteien:record_detail", pk=pk)

        # --- operator past-month restriction ---
        user = request.user
        if user.has_past_months_restrictions:
            from .validators import get_allowed_months
            allowed, _reason = get_allowed_months(record.year)
            if f"month_{effective_from_month}" not in allowed:
                messages.error(
                    request,
                    f"Sie dürfen Monat {effective_from_month} nicht mehr ändern "
                    "(vergangener Monat).",
                )
                return redirect("karteien:record_detail", pk=pk)

        STATUS_LABELS = {
            ContractStatusKind.ACTIVE: "Aktiv",
            ContractStatusKind.PAUSED: "Pausiert",
            ContractStatusKind.TERMINATED: "Gekuendigt",
        }
        status_label = STATUS_LABELS.get(kind, kind)
        proposal = build_contract_status_proposal(
            record,
            effective_from_month=effective_from_month,
            kind=kind,
            comment=comment,
            user=user,
        )

        create_or_update_pending_change_from_snapshot(
            record, snapshot=proposal.snapshot, admin_comment=comment,
        )

        # Set status = PENDING
        KarteiRecord.objects.filter(pk=pk).update(status=RecordStatus.PENDING)

        messages.info(
            request,
            f"Vertragsstatus-Änderung ({status_label} ab Monat "
            f"{effective_from_month}) wurde zur Genehmigung eingereicht.",
        )
        return redirect("karteien:record_detail", pk=pk)


# =====================================================================
# Quick-set subject*_ref  (PROMPT 149.3)
# =====================================================================

def _find_subject_ref_candidates(record) -> dict:
    """Return unambiguous Subject matches for legacy text fields.

    Returns ``{"sem1": {"subject": Subject, "legacy_text": str} | None,
               "sem2": ... }``.
    A match is included **only** when:
    * ``subject*_ref`` is NULL *and* legacy text is non-empty
    * Exactly one active Subject has a normalised name equal to the legacy text
    """
    from apps.catalog.models import Subject
    from .billing import _normalize_subject_name

    result: dict = {"sem1": None, "sem2": None}
    for sem, ref_field, text_field in (
        ("sem1", "subject1_ref", "subject1"),
        ("sem2", "subject2_ref", "subject2"),
    ):
        if getattr(record, ref_field) is not None:
            continue
        legacy_text = (getattr(record, text_field, "") or "").strip()
        if not legacy_text:
            continue
        norm = _normalize_subject_name(legacy_text)
        if not norm:
            continue
        # Query active subjects; filter in Python for casefold match
        candidates = [
            s for s in Subject.objects.filter(is_active=True).iterator()
            if _normalize_subject_name(s.name) == norm
        ]
        if len(candidates) == 1:
            result[sem] = {"subject": candidates[0], "legacy_text": legacy_text}
    return result


class QuickSetSubjectRefView(KarteiEditorMixin, View):
    """POST-only: set subject*_ref from unambiguous legacy text match.

    Creates a PendingChange so the billing impact is reviewed.
    """

    http_method_names = ["post"]

    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        record = get_object_or_404(KarteiRecord, pk=pk)
        detail_url = reverse("karteien:record_detail", kwargs={"pk": pk})

        # Guard: status must be NORMAL
        if record.status != RecordStatus.NORMAL:
            messages.error(
                request,
                "subject_ref kann nur auf Datensätze im Status NORMAL "
                "angewendet werden.",
            )
            return redirect(detail_url)

        semester_key = request.POST.get("semester")  # "sem1" or "sem2"
        if semester_key not in ("sem1", "sem2"):
            messages.error(request, "Ungültiger Semester-Parameter.")
            return redirect(detail_url)

        candidates = _find_subject_ref_candidates(record)
        match = candidates.get(semester_key)
        if match is None:
            messages.error(
                request,
                "Keine eindeutige Übereinstimmung (mehr) gefunden.",
            )
            return redirect(detail_url)

        subject = match["subject"]
        ref_field = "subject1_ref" if semester_key == "sem1" else "subject2_ref"
        sem_label = "1. HJ" if semester_key == "sem1" else "2. HJ"
        proposal = build_quick_set_subject_ref_proposal(
            record,
            semester_key=semester_key,
            subject=subject,
        )
        create_or_update_pending_change_from_snapshot(
            record,
            snapshot=proposal.snapshot,
            admin_comment=(
                f"Quick-set {ref_field} \u2192 {subject.name} "
                f"(\u00fcbereinstimmend mit Legacy-Text)"
            ),
        )

        # Persist only approval status; FK changes live in the frozen payload.
        KarteiRecord.objects.filter(pk=pk).update(status=RecordStatus.PENDING)

        messages.info(
            request,
            f"{sem_label}: {ref_field} wurde auf \u201e{subject.name}\u201c gesetzt. "
            f"\u00c4nderung wurde zur Genehmigung eingereicht.",
        )
        return redirect(detail_url)
