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
from django.db.models import Q
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

from apps.approvals.services import (
    build_snapshot,
    classify_change,
    create_or_update_pending_change,
    get_changed_tracked_fields,
)

from .billing import recalculate_record_months, build_base_amounts
from .forms import KarteiRecordForm, KarteiRecordFilterForm, MonthsOverrideForm
from .models import KarteiRecord, RecordStatus, MonthsMode
from .validators import validate_kartei_record, apply_operator_filters


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
        qs = KarteiRecord.objects.all()
        
        # Default to current year if not specified
        year = self.request.GET.get("year")
        if year:
            try:
                qs = qs.filter(year=int(year))
            except ValueError:
                pass
        else:
            qs = qs.filter(year=date.today().year)
        
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
        
        # Filter by contract status (active/terminated)
        contract_status = self.request.GET.get("contract_status")
        if contract_status:
            if contract_status == "active":
                qs = qs.filter(is_contract_terminated=False)
            elif contract_status == "terminated":
                qs = qs.filter(is_contract_terminated=True)
        
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
        
        # Current filter values
        context["current_year"] = self.request.GET.get("year", date.today().year)
        context["current_filters"] = {
            "family_id": self.request.GET.get("family_id", ""),
            "parent": self.request.GET.get("parent", ""),
            "child": self.request.GET.get("child", ""),
            "status": self.request.GET.get("status", ""),
            "contract_type": self.request.GET.get("contract_type", ""),
            "contract_status": self.request.GET.get("contract_status", ""),
        }
        
        return context


# =============================================================================
# Detail View
# =============================================================================

class KarteiRecordDetailView(KarteiViewerMixin, DetailView):
    """
    Detail view for a single KarteiRecord.
    
    Access: Admin, Operator (full), Superadmin (read-only)
    """
    
    model = KarteiRecord
    template_name = "karteien/record_detail.html"
    context_object_name = "record"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add related information to context."""
        context = super().get_context_data(**kwargs)
        record = self.object
        
        # Check for pending/declined changes
        context["has_pending"] = hasattr(record, "pending_change") and record.pending_change
        context["has_declined"] = record.declined_changes.exists()
        
        # User restrictions
        context["user_has_sepa_restrictions"] = self.request.user.has_sepa_restrictions
        context["user_has_past_months_restrictions"] = self.request.user.has_past_months_restrictions
        context["is_sepa_record"] = record.is_sepa
        
        return context


# =============================================================================
# Create View
# =============================================================================

class KarteiRecordCreateView(KarteiEditorMixin, CreateView):
    """
    Create view for new KarteiRecord entries.
    
    New records are always considered "safe" and don't require approval.
    """
    
    model = KarteiRecord
    form_class = KarteiRecordForm
    template_name = "karteien/record_form.html"
    
    def get_form_kwargs(self) -> dict[str, Any]:
        """Pass user and year to form."""
        kwargs = super().get_form_kwargs()
        kwargs["user"] = self.request.user
        kwargs["year"] = int(self.request.GET.get("year", date.today().year))
        return kwargs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add year info to context."""
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["year"] = int(self.request.GET.get("year", date.today().year))
        
        # New records are AUTO mode by default
        context["is_auto_mode"] = True
        context["months_mode"] = MonthsMode.AUTO
        
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
        """Check record status before allowing access to the edit form."""
        self.object = self.get_object()
        
        # Block editing for PENDING records
        if self.object.status == RecordStatus.PENDING:
            messages.error(
                request,
                "Datensatz ist PENDING. Änderungen müssen über den "
                "Freigabe-Workflow erfolgen. Bitte warten Sie auf die "
                "Entscheidung des Superadmin."
            )
            return redirect("karteien:record_detail", pk=self.object.pk)
        
        # Block editing for DECLINED records
        if self.object.status == RecordStatus.DECLINED:
            messages.error(
                request,
                "Datensatz ist DECLINED. Bitte nutzen Sie die Ansicht "
                "\"Abgelehnte Änderungen\", um Korrekturen erneut einzureichen."
            )
            return redirect("karteien:record_detail", pk=self.object.pk)
        
        return super().dispatch(request, *args, **kwargs)
    
    def get_form_kwargs(self) -> dict[str, Any]:
        """Pass user and year to form."""
        kwargs = super().get_form_kwargs()
        kwargs["user"] = self.request.user
        kwargs["year"] = self.object.year
        return kwargs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add record info to context."""
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["year"] = self.object.year
        context["record"] = self.object
        
        # AUTO mode info
        context["is_auto_mode"] = self.object.months_mode == MonthsMode.AUTO
        context["months_mode"] = self.object.months_mode
        
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
        
        return context
    
    def form_valid(self, form) -> HttpResponse:
        """Handle save with safe/risky classification and AUTO billing.
        
        SAFE path (direct update) is only allowed for records with status NORMAL.
        For any other status, changes must go through the approvals workflow.
        """
        record = form.save(commit=False)
        original = KarteiRecord.objects.get(pk=self.object.pk)
        user = self.request.user
        
        # Double-check status (defense in depth, dispatch should have caught this)
        if original.status != RecordStatus.NORMAL:
            if original.status == RecordStatus.PENDING:
                messages.error(
                    self.request,
                    "Datensatz ist PENDING. Änderungen müssen über den "
                    "Freigabe-Workflow erfolgen."
                )
            else:  # DECLINED
                messages.error(
                    self.request,
                    "Datensatz ist DECLINED. Bitte nutzen Sie die Ansicht "
                    "\"Abgelehnte Änderungen\", um Korrekturen erneut einzureichen."
                )
            return redirect("karteien:record_detail", pk=original.pk)
        
        # Handle AUTO mode billing calculations
        if record.months_mode == MonthsMode.AUTO:
            billing_data = form.get_billing_data()
            record.hours_amounts = billing_data['hours_amounts']
            
            # Recalculate with partial updates if price changed
            flags = recalculate_record_months(
                record,
                apply_from_month_1=billing_data['apply_from_month_1'],
                apply_from_month_2=billing_data['apply_from_month_2'],
                hours_amounts=billing_data['hours_amounts'],
            )
            
            # Show warning if percent discount was clamped
            if flags.percent_discount_exceeded:
                messages.warning(
                    self.request,
                    f"Warnung: Die Summe der Prozentrabatte ({flags.original_percent_sum * 100:.0f}%) "
                    f"wurde auf 99% begrenzt."
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
        
        # Classify the change
        classification = classify_change(record, original)
        
        if classification == "SAFE":
            # Safe change: apply directly
            record.last_change_role = user.role
            record.last_change_date = date.today()
            record.last_change_time = date.today().strftime("%H:%M")
            record.save()
            
            messages.success(
                self.request, 
                f"Datensatz {record.id} wurde aktualisiert."
            )
            
            # TODO: Update history_raw (requires history module)
            
        else:
            # Risky change: create pending
            pending = create_or_update_pending_change(record)
            
            # Mark record as PENDING
            KarteiRecord.objects.filter(pk=record.pk).update(
                status=RecordStatus.PENDING
            )
            
            messages.info(
                self.request,
                f"Änderungen für Datensatz {record.id} wurden zur Genehmigung eingereicht. "
                "Ein Superadmin muss die Änderungen prüfen."
            )
            
            # TODO: Create notification for Superadmin
        
        return redirect("karteien:record_detail", pk=record.pk)
    
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
        pending = create_or_update_pending_change(record)
        
        # Add reason to pending change comment
        if pending and hasattr(pending, 'comment'):
            pending.comment = f"OVERRIDE: {reason}"
            pending.save()
        
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
