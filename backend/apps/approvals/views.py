"""
Views for the approvals app.

This module contains views for:
- DeclinedOverview: List of declined records for Admin review
- Applying fixes to declined records (move back to pending)
- Viewing pending changes (for information)

Superadmin views (added in PROMPT_09):
- SuperadminPendingOverviewView: List all pending changes for Superadmin
- SuperadminWarIstView: War/Ist comparison for single pending change
- Bulk actions: ApproveAll, DeclineAll, ClearAll
- NeuListView: New records view

Mirrors functionality from:
- Export_DeclinedTools.bas: ShowDeclinedOverview, ApplyDeclinedFixes
- Export_DeclinedHelpers.bas: Record comparison and movement logic
- valid_ApproveFlow.bas: LoadPendingChanges, SyncDecisions
- valid_GrossGeschichteDecision.bas: War/Ist comparison
- valid_NeuList.bas: New records tracking
"""

from __future__ import annotations

from datetime import date
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db import transaction
from django.http import HttpRequest, HttpResponse, HttpResponseRedirect
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse, reverse_lazy
from django.views import View
from django.views.generic import DetailView, ListView, TemplateView

from apps.karteien.models import KarteiRecord, RecordStatus, TRACKED_FIELDS

from .forms import DeclinedChangeEditForm, PendingChangeEditForm
from .models import DeclinedChange, PendingChange
from .services import (
    apply_decision,
    approve_all_pending,
    build_snapshot,
    create_or_update_pending_change,
    create_or_update_pending_change_from_snapshot,
    decline_all_pending,
    get_changed_tracked_fields,
    get_new_records,
    get_new_records_count,
    update_last_seen_id,
)


# =============================================================================
# Permission Mixins
# =============================================================================

class AdminEditorMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to Admin role users.
    
    Only Admin can access declined overview and apply fixes.
    """
    
    def test_func(self) -> bool:
        """Check if user is Admin."""
        user = self.request.user
        return user.is_authenticated and user.can_edit_kartei
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        messages.error(
            self.request,
            "Sie haben keine Berechtigung für diese Aktion."
        )
        return redirect("karteien:record_list")


class SuperadminMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to Superadmin role users.
    
    Only Superadmin can approve/decline pending changes and view War/Ist.
    """
    
    def test_func(self) -> bool:
        """Check if user is Superadmin."""
        user = self.request.user
        return user.is_authenticated and user.is_superadmin
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        messages.error(
            self.request,
            "Diese Funktion ist nur für Superadmin verfügbar."
        )
        return redirect("karteien:record_list")


# =============================================================================
# Declined Overview
# =============================================================================

class DeclinedOverviewView(AdminEditorMixin, ListView):
    """
    List view showing all declined records.
    
    Mirrors VBA: Export_DeclinedTools.ShowDeclinedOverview
    
    Shows:
    - Record ID, FamilyID, Parent, Child
    - Decline reason
    - Decline date and by whom
    - Number of previous declines (from history)
    - Actions: View/Edit record, Apply fix
    """
    
    model = DeclinedChange
    template_name = "approvals/declined_overview.html"
    context_object_name = "declined_changes"
    paginate_by = 50
    
    def get_queryset(self):
        """Get declined changes, optionally filtered by year, family_id, parent, child."""
        qs = DeclinedChange.objects.select_related(
            "record", "declined_by"
        ).order_by("-created_at")
        
        # Filter by year if specified
        year = self.request.GET.get("year")
        if year:
            try:
                qs = qs.filter(record__year=int(year))
            except ValueError:
                pass
        else:
            # Default to current year
            qs = qs.filter(record__year=date.today().year)
        
        # Filter by family_id if specified
        family_id = self.request.GET.get("family_id")
        if family_id:
            qs = qs.filter(record__family_id__icontains=family_id)
        
        # Filter by parent name if specified
        parent = self.request.GET.get("parent")
        if parent:
            qs = qs.filter(record__parent_name__icontains=parent)
        
        # Filter by child name if specified
        child = self.request.GET.get("child")
        if child:
            qs = qs.filter(record__child_name__icontains=child)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add context for template."""
        context = super().get_context_data(**kwargs)
        
        # Available years
        years = (
            KarteiRecord.objects.filter(status=RecordStatus.DECLINED)
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        context["current_year"] = self.request.GET.get("year", date.today().year)
        
        # Count by status
        context["declined_count"] = (
            KarteiRecord.objects.filter(status=RecordStatus.DECLINED).count()
        )
        context["pending_count"] = (
            KarteiRecord.objects.filter(status=RecordStatus.PENDING).count()
        )
        
        return context


class DeclinedDetailView(AdminEditorMixin, DetailView):
    """
    Detail view for a single declined change.
    
    Shows the full snapshot of declined values along with the
    current record values for comparison (War/Ist).
    """
    
    model = DeclinedChange
    template_name = "approvals/declined_detail.html"
    context_object_name = "declined_change"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add comparison data to context."""
        context = super().get_context_data(**kwargs)
        declined = self.object
        record = declined.record
        
        # Build current snapshot for comparison
        current_snapshot = build_snapshot(record)
        declined_snapshot = declined.snapshot
        
        # Build comparison data
        from apps.karteien.models import TRACKED_FIELDS
        comparisons = []
        for field in TRACKED_FIELDS:
            current_val = current_snapshot.get(field)
            declined_val = declined_snapshot.get(field)
            
            comparisons.append({
                "field": field,
                "current": current_val,
                "declined": declined_val,
                "changed": current_val != declined_val,
            })
        
        context["comparisons"] = comparisons
        context["record"] = record
        
        return context


class DeclinedChangeEditView(AdminEditorMixin, View):
    """
    Edit view for DeclinedChange.snapshot.
    
    Allows Admin to edit the proposed values that were declined,
    without modifying the actual KarteiRecord.
    
    GET: Display form prefilled with snapshot values.
    POST: Validate and update snapshot, redirect to detail view.
    """
    
    template_name = "approvals/declined_edit.html"
    
    def get(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Show edit form prefilled with snapshot values."""
        declined = get_object_or_404(DeclinedChange, pk=pk)
        record = declined.record
        
        form = DeclinedChangeEditForm(snapshot=declined.snapshot)
        
        return render(request, self.template_name, {
            "form": form,
            "declined_change": declined,
            "record": record,
        })
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Validate and update snapshot."""
        declined = get_object_or_404(DeclinedChange, pk=pk)
        record = declined.record
        
        form = DeclinedChangeEditForm(request.POST, snapshot=declined.snapshot)
        
        if form.is_valid():
            # Update snapshot with new values
            declined.snapshot = form.to_snapshot()
            declined.save(update_fields=["snapshot"])
            
            messages.success(
                request,
                f"Snapshot für Datensatz {record.id} wurde aktualisiert."
            )
            
            return redirect("approvals:declined_detail", pk=pk)
        
        # Form invalid - re-render with errors
        return render(request, self.template_name, {
            "form": form,
            "declined_change": declined,
            "record": record,
        })


# =============================================================================
# Apply Fixes
# =============================================================================

class ApplyDeclinedFixView(AdminEditorMixin, View):
    """
    Apply fix to a single declined record.
    
    Mirrors VBA: Export_DeclinedTools.ApplyDeclinedFixes
    
    Process:
    1. Take snapshot from DeclinedChange (with admin's corrections)
    2. Create new PendingChange with the corrected snapshot
    3. Delete the DeclinedChange
    4. Change record status from DECLINED to PENDING
    """
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Apply fix for a single declined change."""
        declined = get_object_or_404(DeclinedChange, pk=pk)
        record = declined.record
        
        with transaction.atomic():
            # Create new pending change with the declined snapshot (which may have been edited)
            pending = create_or_update_pending_change_from_snapshot(record, declined.snapshot)
            
            # Update record status to PENDING
            record.status = RecordStatus.PENDING
            record.save(update_fields=["status"])
            
            # Delete the declined change (or mark as processed)
            # For audit trail, we could keep it with a processed flag
            # For now, delete to match VBA behavior
            declined.delete()
        
        messages.success(
            request,
            f"Datensatz {record.id} wurde zur erneuten Genehmigung eingereicht. "
            "Status geändert von DECLINED zu PENDING."
        )
        
        return redirect("approvals:declined_overview")


class ApplyAllDeclinedFixesView(AdminEditorMixin, View):
    """
    Apply fixes to all declined records.
    
    Bulk operation to move all DECLINED records to PENDING.
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Apply fixes to all declined records."""
        year = request.POST.get("year", date.today().year)
        
        try:
            year = int(year)
        except ValueError:
            year = date.today().year
        
        # Get all declined changes for the year
        declined_changes = DeclinedChange.objects.filter(
            record__year=year
        ).select_related("record")
        
        moved_count = 0
        error_count = 0
        
        with transaction.atomic():
            for declined in declined_changes:
                try:
                    record = declined.record
                    
                    # Create pending change with the declined snapshot
                    create_or_update_pending_change_from_snapshot(record, declined.snapshot)
                    
                    # Update status
                    record.status = RecordStatus.PENDING
                    record.save(update_fields=["status"])
                    
                    # Delete declined change
                    declined.delete()
                    
                    moved_count += 1
                except Exception as e:
                    error_count += 1
        
        if moved_count > 0:
            messages.success(
                request,
                f"{moved_count} Datensatz(e) wurden zur erneuten Genehmigung eingereicht."
            )
        
        if error_count > 0:
            messages.warning(
                request,
                f"{error_count} Datensatz(e) konnten nicht verarbeitet werden."
            )
        
        if moved_count == 0 and error_count == 0:
            messages.info(request, "Keine abgelehnten Datensätze gefunden.")
        
        return redirect("approvals:declined_overview")


# =============================================================================
# Pending Changes View (Info)
# =============================================================================

class PendingChangesListView(AdminEditorMixin, ListView):
    """
    List view showing pending changes (for Admin information).
    
    Admin can see what changes are waiting for Superadmin approval.
    """
    
    model = PendingChange
    template_name = "approvals/pending_list.html"
    context_object_name = "pending_changes"
    paginate_by = 50
    
    def get_queryset(self):
        """Get unprocessed pending changes."""
        qs = PendingChange.objects.filter(
            is_processed=False
        ).select_related("record").order_by("-created_at")
        
        # Filter by year
        year = self.request.GET.get("year")
        if year:
            try:
                qs = qs.filter(record__year=int(year))
            except ValueError:
                pass
        else:
            qs = qs.filter(record__year=date.today().year)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add context."""
        context = super().get_context_data(**kwargs)
        
        years = (
            KarteiRecord.objects.filter(status=RecordStatus.PENDING)
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        context["current_year"] = self.request.GET.get("year", date.today().year)
        
        return context


class PendingDetailView(AdminEditorMixin, DetailView):
    """
    Detail view for a pending change.
    
    Shows what changes are proposed.
    """
    
    model = PendingChange
    template_name = "approvals/pending_detail.html"
    context_object_name = "pending_change"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add comparison data."""
        context = super().get_context_data(**kwargs)
        pending = self.object
        record = pending.record
        
        # Build comparison
        from apps.karteien.models import TRACKED_FIELDS
        pending_snapshot = pending.snapshot
        comparisons = []
        
        for field in TRACKED_FIELDS:
            current_val = getattr(record, field, None)
            pending_val = pending_snapshot.get(field)
            
            # Normalize for comparison
            if current_val is None:
                current_val = ""
            if pending_val is None:
                pending_val = ""
            
            comparisons.append({
                "field": field,
                "current": current_val,
                "pending": pending_val,
                "changed": str(current_val) != str(pending_val),
            })
        
        context["comparisons"] = comparisons
        context["record"] = record
        
        return context


class PendingChangeEditView(AdminEditorMixin, View):
    """
    Edit view for PendingChange.snapshot (ADMIN only).
    
    Allows Admin to edit a pending change before Superadmin decides.
    This updates PendingChange.snapshot and admin_comment without
    affecting the actual KarteiRecord.
    
    Includes optimistic conflict guard to prevent edit if the pending
    change was already processed or modified by another user.
    
    GET: Display form prefilled with snapshot values and admin_comment.
    POST: Validate and update snapshot + admin_comment, notify Superadmin.
    """
    
    template_name = "approvals/pending_edit.html"
    
    def get(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Show edit form prefilled with snapshot values."""
        pending = get_object_or_404(PendingChange, pk=pk, is_processed=False)
        record = pending.record
        
        # Check record is still PENDING
        if record.status != RecordStatus.PENDING:
            messages.warning(
                request,
                "Dieser Datensatz ist nicht mehr im Status 'Wartend'."
            )
            return redirect("approvals:pending_detail", pk=pk)
        
        form = PendingChangeEditForm(
            snapshot=pending.snapshot,
            admin_comment=pending.admin_comment or "",
        )
        
        return render(request, self.template_name, {
            "form": form,
            "pending_change": pending,
            "record": record,
            "pending_updated_at": pending.updated_at.isoformat(),
        })
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Validate and update snapshot + admin_comment."""
        # Reload pending from DB to check current state
        pending = PendingChange.objects.filter(pk=pk).first()
        
        if not pending:
            messages.error(request, "Diese wartende Änderung wurde nicht gefunden.")
            return redirect("approvals:pending_list")
        
        # Check if already processed
        if pending.is_processed:
            messages.warning(
                request,
                "Diese Änderung wurde bereits bearbeitet (genehmigt oder abgelehnt). "
                "Bearbeitung nicht möglich."
            )
            return redirect("approvals:pending_detail", pk=pk)
        
        record = pending.record
        
        # Check record is still PENDING
        if record.status != RecordStatus.PENDING:
            messages.warning(
                request,
                "Dieser Datensatz ist nicht mehr im Status 'Wartend'."
            )
            return redirect("approvals:pending_detail", pk=pk)
        
        # Optimistic conflict guard: check updated_at
        submitted_updated_at = request.POST.get("pending_updated_at", "")
        if submitted_updated_at:
            try:
                from datetime import datetime
                # Parse ISO format
                submitted_dt = datetime.fromisoformat(submitted_updated_at.replace("Z", "+00:00"))
                # Compare with current pending.updated_at
                # Allow small delta for timezone rounding
                current_dt = pending.updated_at
                if hasattr(current_dt, "isoformat"):
                    if submitted_dt.isoformat() != current_dt.isoformat():
                        messages.warning(
                            request,
                            "Diese Änderung wurde zwischenzeitlich aktualisiert. "
                            "Bitte laden Sie die Seite neu und versuchen Sie es erneut."
                        )
                        return redirect("approvals:pending_edit", pk=pk)
            except (ValueError, TypeError):
                # If we can't parse, proceed anyway (fail-open for usability)
                pass
        
        form = PendingChangeEditForm(
            request.POST,
            snapshot=pending.snapshot,
            admin_comment=pending.admin_comment or "",
        )
        
        if form.is_valid():
            # Update snapshot with new values
            pending.snapshot = form.to_snapshot()
            pending.admin_comment = form.cleaned_data["admin_comment"]
            pending.save(update_fields=["snapshot", "admin_comment", "updated_at"])
            
            # Notify Superadmin about the update (raises pending in their list)
            from apps.notifications.services import notify_pending_created
            notify_pending_created(record, pending)
            
            messages.success(
                request,
                f"Wartende Änderung für Datensatz {record.id} wurde aktualisiert."
            )
            
            return redirect("karteien:record_detail", pk=record.pk)
        
        # Form invalid - re-render with errors
        return render(request, self.template_name, {
            "form": form,
            "pending_change": pending,
            "record": record,
            "pending_updated_at": pending.updated_at.isoformat(),
        })


# =============================================================================
# Superadmin Pending Overview
# =============================================================================

class SuperadminPendingOverviewView(SuperadminMixin, ListView):
    """
    Superadmin view: List of all pending changes awaiting decision.
    
    Mirrors VBA: valid_ImportPending.LoadPendingChangesFromPre
    
    Shows:
    - All unprocessed pending changes
    - Key record fields (ID, FamilyID, Parent, Child)
    - Change type summary
    - Filters by year, FamilyID, Parent, Child
    - Links to War/Ist detail for each pending
    """
    
    model = PendingChange
    template_name = "approvals/superadmin/pending_overview.html"
    context_object_name = "pending_changes"
    paginate_by = 50
    
    def get_queryset(self):
        """Get unprocessed pending changes with filters."""
        qs = PendingChange.objects.filter(
            is_processed=False
        ).select_related("record").order_by("-created_at")
        
        # Year filter - "all" means no year filter
        year = self.request.GET.get("year")
        if year == "all":
            # No year filter - show all years
            pass
        elif year:
            try:
                qs = qs.filter(record__year=int(year))
            except ValueError:
                # Default to current year on invalid value
                qs = qs.filter(record__year=date.today().year)
        else:
            # Default to current year
            qs = qs.filter(record__year=date.today().year)
        
        # FamilyID filter
        family_id = self.request.GET.get("family_id")
        if family_id:
            qs = qs.filter(record__family_id__icontains=family_id)
        
        # Parent filter
        parent = self.request.GET.get("parent")
        if parent:
            qs = qs.filter(record__parent_name__icontains=parent)
        
        # Child filter
        child = self.request.GET.get("child")
        if child:
            qs = qs.filter(record__child_name__icontains=child)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add context for template."""
        context = super().get_context_data(**kwargs)
        
        # Available years
        years = (
            KarteiRecord.objects.filter(status=RecordStatus.PENDING)
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        
        # Current year can be "all" or an integer
        year_param = self.request.GET.get("year", "")
        if year_param == "all":
            context["current_year"] = "all"
        elif year_param:
            try:
                context["current_year"] = int(year_param)
            except ValueError:
                context["current_year"] = date.today().year
        else:
            context["current_year"] = date.today().year
        
        # Current filters
        context["filter_family_id"] = self.request.GET.get("family_id", "")
        context["filter_parent"] = self.request.GET.get("parent", "")
        context["filter_child"] = self.request.GET.get("child", "")
        
        # Total counts - for "all" count without year filter
        if context["current_year"] == "all":
            context["pending_count"] = PendingChange.objects.filter(
                is_processed=False
            ).count()
            context["declined_count"] = KarteiRecord.objects.filter(
                status=RecordStatus.DECLINED
            ).count()
        else:
            context["pending_count"] = PendingChange.objects.filter(
                is_processed=False,
                record__year=context["current_year"]
            ).count()
            context["declined_count"] = KarteiRecord.objects.filter(
                status=RecordStatus.DECLINED,
                year=context["current_year"]
            ).count()
        
        # New records count - only for specific year
        try:
            if context["current_year"] != "all":
                context["new_records_count"] = get_new_records_count(
                    self.request.user, context["current_year"]
                )
            else:
                context["new_records_count"] = 0
        except Exception:
            context["new_records_count"] = 0
        
        return context


# =============================================================================
# Superadmin War/Ist (Detail) View
# =============================================================================

class SuperadminWarIstView(SuperadminMixin, DetailView):
    """
    Superadmin view: War/Ist comparison for a single pending change.
    
    Mirrors VBA: valid_GrossGeschichteDecision.BuildPendingDecisionSheet
    
    Shows:
    - War (original values from KarteiRecord)
    - Ist (pending values from snapshot)
    - Differences highlighted
    - Decision form (Approved/Declined + comment)
    - Link to record history
    """
    
    model = PendingChange
    template_name = "approvals/superadmin/war_ist.html"
    context_object_name = "pending_change"
    
    def get_queryset(self):
        """Only show unprocessed pending changes."""
        return PendingChange.objects.filter(
            is_processed=False
        ).select_related("record")
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Build War/Ist comparison data."""
        context = super().get_context_data(**kwargs)
        pending = self.object
        record = pending.record
        
        # Get pending snapshot
        pending_snapshot = pending.snapshot
        
        # Build comparison for all tracked fields
        comparisons = []
        for field_name in TRACKED_FIELDS:
            # War = current value in record (original)
            war_value = getattr(record, field_name, None)
            # Ist = value from pending snapshot (proposed)
            ist_value = pending_snapshot.get(field_name)
            
            # Normalize for display
            if war_value is None:
                war_value = ""
            if ist_value is None:
                ist_value = ""
            
            # Check if changed
            changed = str(war_value) != str(ist_value)
            
            comparisons.append({
                "field": field_name,
                "field_label": _get_field_label(field_name),
                "war": war_value,
                "ist": ist_value,
                "changed": changed,
            })
        
        context["comparisons"] = comparisons
        context["record"] = record
        context["changed_count"] = sum(1 for c in comparisons if c["changed"])
        
        # Check if record has history
        context["has_history"] = bool(record.history_raw)
        
        # Optimistic conflict guard: store exact timestamp for comparison
        # Using isoformat() preserves microseconds (unlike Django's date:'c' filter)
        context["pending_updated_at"] = pending.updated_at.isoformat()
        
        return context
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Handle decision submission with optimistic conflict guard."""
        pending = get_object_or_404(PendingChange, pk=pk, is_processed=False)
        
        # Optimistic conflict guard: check pending_updated_at from form
        # against current pending.updated_at to detect if Admin updated
        # the pending change while Superadmin had the page open.
        submitted_updated_at = request.POST.get("pending_updated_at", "")
        if submitted_updated_at:
            try:
                from datetime import datetime
                # Parse ISO format (handles both with and without microseconds)
                submitted_dt = datetime.fromisoformat(
                    submitted_updated_at.replace("Z", "+00:00")
                )
                current_dt = pending.updated_at
                
                # Compare as datetime objects, not strings.
                # This avoids issues with microseconds or timezone formatting.
                # Both must be aware datetimes for proper comparison.
                if submitted_dt != current_dt:
                    messages.warning(
                        request,
                        "Diese Änderung wurde inzwischen aktualisiert. "
                        "Bitte laden Sie die Seite neu und prüfen Sie erneut."
                    )
                    return redirect("approvals:superadmin_war_ist", pk=pk)
            except (ValueError, TypeError):
                # If we can't parse, proceed anyway (fail-open for usability)
                pass
        
        decision = request.POST.get("decision", "").upper()
        comment = request.POST.get("comment", "").strip()
        
        if decision not in ("APPROVED", "DECLINED"):
            messages.error(request, "Ungültige Entscheidung. Bitte wählen Sie 'Approved' oder 'Declined'.")
            return redirect("approvals:superadmin_war_ist", pk=pk)
        
        try:
            apply_decision(pending, decision, comment, request.user)
            
            if decision == "APPROVED":
                messages.success(
                    request,
                    f"Änderung für Datensatz {pending.record_id} wurde genehmigt."
                )
            else:
                messages.success(
                    request,
                    f"Änderung für Datensatz {pending.record_id} wurde abgelehnt."
                )
        except Exception as e:
            messages.error(request, f"Fehler bei der Verarbeitung: {str(e)}")
            return redirect("approvals:superadmin_war_ist", pk=pk)
        
        return redirect("approvals:superadmin_pending_overview")


def _get_field_label(field_name: str) -> str:
    """Get human-readable label for a field name."""
    labels = {
        "family_id": "FamilyID",
        "parent_name": "Eltern",
        "child_name": "Kind",
        "birthdate": "Geburtsdatum",
        "address": "Adresse",
        "phone": "Telefon",
        "mobile": "Mobil",
        "email": "E-Mail",
        "subject1": "Fach 1",
        "price1": "Preis 1",
        "subject2": "Fach 2",
        "price2": "Preis 2",
        "extra1": "Extra 1",
        "extra2": "Extra 2",
        "extra3": "Extra 3",
    }
    # Month fields
    if field_name.startswith("month_"):
        month_num = field_name.replace("month_", "")
        return f"Monat {month_num}"
    
    return labels.get(field_name, field_name)


# =============================================================================
# Superadmin Bulk Actions
# =============================================================================

class SuperadminApproveAllView(SuperadminMixin, View):
    """
    Approve all pending changes.
    
    Mirrors VBA: valid_ApproveFlow.ApproveAllPending
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Process bulk approve."""
        year_param = request.POST.get("year")
        if year_param == "all":
            year = None  # No year filter
        else:
            try:
                year = int(year_param) if year_param else date.today().year
            except ValueError:
                year = date.today().year
        
        # Confirmation check
        confirm = request.POST.get("confirm")
        if confirm != "yes":
            messages.warning(
                request,
                "Bitte bestätigen Sie die Aktion durch Aktivieren der Checkbox."
            )
            return redirect("approvals:superadmin_pending_overview")
        
        approved_count, errors = approve_all_pending(request.user, year)
        
        if approved_count > 0:
            messages.success(
                request,
                f"{approved_count} Änderung(en) wurden genehmigt."
            )
        
        if errors:
            messages.warning(
                request,
                f"{len(errors)} Fehler aufgetreten: " + "; ".join(errors[:3])
            )
        
        if approved_count == 0 and not errors:
            messages.info(request, "Keine wartenden Änderungen zum Genehmigen gefunden.")
        
        return redirect("approvals:superadmin_pending_overview")


class SuperadminDeclineAllView(SuperadminMixin, View):
    """
    Decline all pending changes with a common comment.
    
    Mirrors VBA: valid_ApproveFlow.DeclineAllPending
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Process bulk decline."""
        year_param = request.POST.get("year")
        if year_param == "all":
            year = None  # No year filter
        else:
            try:
                year = int(year_param) if year_param else date.today().year
            except ValueError:
                year = date.today().year
        
        comment = request.POST.get("comment", "").strip()
        if not comment:
            messages.error(
                request,
                "Bitte geben Sie eine Begründung für die Ablehnung ein."
            )
            return redirect("approvals:superadmin_pending_overview")
        
        declined_count, errors = decline_all_pending(comment, request.user, year)
        
        if declined_count > 0:
            messages.success(
                request,
                f"{declined_count} Änderung(en) wurden abgelehnt."
            )
        
        if errors:
            messages.warning(
                request,
                f"{len(errors)} Fehler aufgetreten: " + "; ".join(errors[:3])
            )
        
        if declined_count == 0 and not errors:
            messages.info(request, "Keine wartenden Änderungen zum Ablehnen gefunden.")
        
        return redirect("approvals:superadmin_pending_overview")


# =============================================================================
# Superadmin NeuList (New Records)
# =============================================================================

class SuperadminNeuListView(SuperadminMixin, ListView):
    """
    View new records that have not been seen by the Superadmin.
    
    Mirrors VBA: valid_NeuList.RefreshNeuList
    
    Shows records with ID > last_seen_id for the current user.
    """
    
    model = KarteiRecord
    template_name = "approvals/superadmin/neu_list.html"
    context_object_name = "new_records"
    paginate_by = 50
    
    def get_queryset(self):
        """Get new records based on last_seen_id."""
        year = self.request.GET.get("year")
        try:
            year = int(year) if year else date.today().year
        except ValueError:
            year = date.today().year
        
        return get_new_records(self.request.user, year)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add context for template."""
        context = super().get_context_data(**kwargs)
        
        # Available years
        years = (
            KarteiRecord.objects
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        
        year = self.request.GET.get("year")
        try:
            context["current_year"] = int(year) if year else date.today().year
        except ValueError:
            context["current_year"] = date.today().year
        
        # Get per-year last_seen_id for display
        from .services import get_or_create_superadmin_state, _get_last_seen_id_for_year
        state = get_or_create_superadmin_state(self.request.user)
        context["last_seen_id"] = _get_last_seen_id_for_year(state, context["current_year"])
        context["new_count"] = get_new_records_count(
            self.request.user, context["current_year"]
        )
        
        return context


class SuperadminMarkSeenView(SuperadminMixin, View):
    """
    Mark all current records as seen (update last_seen_id) for a specific year.
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Update last_seen_id to current max for the given year."""
        # Get year from POST data or default to current year
        year_str = request.POST.get("year")
        try:
            year = int(year_str) if year_str else date.today().year
        except ValueError:
            year = date.today().year

        update_last_seen_id(request.user, year=year)
        messages.success(
            request,
            f"Alle Datensätze für Jahr {year} wurden als gesehen markiert."
        )
        # Redirect back with year parameter to preserve selection
        return redirect(f"{reverse('approvals:superadmin_neu_list')}?year={year}")


# =============================================================================
# Superadmin Record History View
# =============================================================================

class SuperadminRecordHistoryView(SuperadminMixin, DetailView):
    """
    View history for a specific record.
    
    Mirrors VBA: Geschichte.bas and valid_ParseHistory.bas
    """
    
    model = KarteiRecord
    template_name = "approvals/superadmin/record_history.html"
    context_object_name = "record"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Parse and display history."""
        context = super().get_context_data(**kwargs)
        record = self.object
        
        # Parse raw history
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
                # Fallback: show raw history if parsing failed
                raw_history_fallback = record.history_raw
        
        context["history_entries"] = history_entries
        context["raw_history"] = raw_history_fallback
        context["has_history"] = bool(history_entries or record.history_raw)
        
        return context
