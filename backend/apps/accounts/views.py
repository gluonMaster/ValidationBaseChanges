"""
Views for the accounts app.

This module contains:
- Role-based redirect function for home page
- User dashboard/home views
- User record search view (read-only)
- User record detail view (read-only)
- User history view (read-only)

All User views are strictly read-only. No write operations are allowed.

See ARCHITECTURE.md Section 2.1 for accounts app overview.
See DOMAIN_MODEL.md Section 3 for role definitions.
"""

from __future__ import annotations

from datetime import date
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.http import HttpRequest, HttpResponse, HttpResponseRedirect
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse
from django.views.generic import TemplateView, ListView, DetailView

from apps.approvals.models import PendingChange, DeclinedChange
from apps.history.models import HistoryEvent
from apps.karteien.models import KarteiRecord, RecordStatus


# =============================================================================
# Role-Based Redirect
# =============================================================================

@login_required
def role_based_redirect(request: HttpRequest) -> HttpResponseRedirect:
    """
    Redirect user to appropriate dashboard based on role.
    
    - USER: /user/ (User cabinet)
    - ADMIN/OPERATOR: /karteien/ (Kartei list)
    - SUPERADMIN: /approvals/superadmin/pending/ (Pending approvals)
    """
    user = request.user
    
    if user.is_user_role:
        return redirect("accounts:user_dashboard")
    elif user.is_superadmin:
        return redirect("approvals:superadmin_pending_overview")
    elif user.can_edit_kartei:
        return redirect("karteien:record_list")
    else:
        # Fallback to login
        return redirect("login")


# =============================================================================
# Permission Mixins
# =============================================================================

class UserRoleMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to users with USER role.
    
    USER role has read-only access to records and history.
    No write operations are allowed.
    """
    
    def test_func(self) -> bool:
        """Check if user has USER role."""
        user = self.request.user
        return user.is_authenticated and user.is_user_role
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect based on user role or to login."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        # Redirect to appropriate dashboard based on role
        user = self.request.user
        if user.is_superadmin:
            return redirect("approvals:superadmin_pending_overview")
        elif user.can_edit_kartei:
            return redirect("karteien:record_list")
        else:
            messages.error(
                self.request,
                "Sie haben keine Berechtigung für diese Seite."
            )
            return redirect("login")


class AnyAuthenticatedUserMixin(LoginRequiredMixin):
    """
    Mixin that allows any authenticated user.
    
    Used for views that should be accessible to all logged-in users
    regardless of role.
    """
    pass


# =============================================================================
# User Dashboard
# =============================================================================

class UserDashboardView(UserRoleMixin, TemplateView):
    """
    User dashboard/home page.
    
    Provides:
    - Search form for records
    - Quick stats (if any)
    - Navigation to search results
    
    GET /user/
    """
    
    template_name = "accounts/user_dashboard.html"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add dashboard context."""
        context = super().get_context_data(**kwargs)
        
        # Current year as default
        context["current_year"] = date.today().year
        
        # Available years for dropdown
        years = (
            KarteiRecord.objects.values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        
        return context


# =============================================================================
# User Search View
# =============================================================================

class UserKarteiSearchView(UserRoleMixin, ListView):
    """
    Search view for User role (read-only).
    
    Allows searching records by:
    - FamilyID (partial match)
    - Parent name (partial match)
    - Child name (partial match)
    - Year
    
    Results are limited to 50 records.
    
    GET /user/search/
    """
    
    model = KarteiRecord
    template_name = "accounts/user_search.html"
    context_object_name = "records"
    paginate_by = 50
    
    def get_queryset(self):
        """Filter records based on search parameters."""
        qs = KarteiRecord.objects.all()
        
        # Check if any search parameters are provided
        has_search = any(
            self.request.GET.get(param)
            for param in ["family_id", "parent", "child"]
        )
        
        # If no search parameters, return empty queryset
        if not has_search:
            return KarteiRecord.objects.none()
        
        # Filter by year
        year = self.request.GET.get("year")
        if year:
            try:
                qs = qs.filter(year=int(year))
            except ValueError:
                pass
        else:
            # Default to current year
            qs = qs.filter(year=date.today().year)
        
        # Filter by FamilyID
        family_id = self.request.GET.get("family_id", "").strip()
        if family_id:
            qs = qs.filter(family_id__icontains=family_id)
        
        # Filter by Parent
        parent = self.request.GET.get("parent", "").strip()
        if parent:
            qs = qs.filter(parent_name__icontains=parent)
        
        # Filter by Child
        child = self.request.GET.get("child", "").strip()
        if child:
            qs = qs.filter(child_name__icontains=child)
        
        return qs.order_by("family_id", "parent_name", "child_name")[:50]
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add search context."""
        context = super().get_context_data(**kwargs)
        
        # Preserve search parameters
        context["search_params"] = {
            "family_id": self.request.GET.get("family_id", ""),
            "parent": self.request.GET.get("parent", ""),
            "child": self.request.GET.get("child", ""),
            "year": self.request.GET.get("year", str(date.today().year)),
        }
        
        # Check if search was performed
        context["search_performed"] = any(
            self.request.GET.get(param)
            for param in ["family_id", "parent", "child"]
        )
        
        # Available years for dropdown
        years = (
            KarteiRecord.objects.values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = list(years) or [date.today().year]
        
        return context


# =============================================================================
# User Record Detail View (Read-Only)
# =============================================================================

class UserKarteiDetailView(UserRoleMixin, DetailView):
    """
    Read-only detail view of a KarteiRecord for User role.
    
    Displays:
    - Basic record information (family, child, contacts, subjects)
    - Record status (normal/pending/declined)
    - Indicator if pending/declined changes exist
    
    No edit buttons are shown. User cannot modify data.
    
    GET /user/record/<pk>/
    """
    
    model = KarteiRecord
    template_name = "accounts/user_record_detail.html"
    context_object_name = "record"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add record context with status info."""
        context = super().get_context_data(**kwargs)
        record = self.object
        
        # Check for pending changes
        context["has_pending"] = PendingChange.objects.filter(
            record=record,
            is_processed=False
        ).exists()
        
        # Check for declined changes
        context["declined_count"] = DeclinedChange.objects.filter(
            record=record
        ).count()
        
        # Status display
        context["status_display"] = self._get_status_display(record.status)
        context["status_class"] = self._get_status_class(record.status)
        
        # Month names for display
        context["months"] = self._get_month_data(record)
        
        return context
    
    def _get_status_display(self, status: str) -> str:
        """Get human-readable status display."""
        status_map = {
            RecordStatus.NORMAL: "Normal",
            RecordStatus.PENDING: "Wartend auf Genehmigung",
            RecordStatus.DECLINED: "Abgelehnt",
        }
        return status_map.get(status, status)
    
    def _get_status_class(self, status: str) -> str:
        """Get CSS class for status badge."""
        class_map = {
            RecordStatus.NORMAL: "bg-success",
            RecordStatus.PENDING: "bg-primary",
            RecordStatus.DECLINED: "bg-danger",
        }
        return class_map.get(status, "bg-secondary")
    
    def _get_month_data(self, record: KarteiRecord) -> list[dict]:
        """Get month data for display."""
        month_names = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ]
        
        months = []
        for i in range(1, 13):
            field_name = f"month_{i}"
            value = getattr(record, field_name, None)
            months.append({
                "number": i,
                "name": month_names[i - 1],
                "value": value,
            })
        return months


# =============================================================================
# User History View
# =============================================================================

class UserRecordHistoryView(UserRoleMixin, DetailView):
    """
    Read-only history view for a KarteiRecord.
    
    Displays:
    - Record basic info
    - Chronological list of history events
    - For each event: date/time, user, changed fields (old → new), comment
    
    GET /user/record/<pk>/history/
    """
    
    model = KarteiRecord
    template_name = "accounts/user_record_history.html"
    context_object_name = "record"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        """Add history events to context."""
        context = super().get_context_data(**kwargs)
        record = self.object
        
        # Get history events ordered by time (newest first)
        events = HistoryEvent.objects.filter(
            record=record
        ).select_related("user").order_by("-event_time", "-id")
        
        # Parse and format events for display
        context["history_events"] = [
            self._format_event(event) for event in events
        ]
        
        # Total events count
        context["total_events"] = events.count()
        
        # Check if there's raw history but no parsed events
        context["has_raw_history"] = bool(record.history_raw)
        
        return context
    
    def _format_event(self, event: HistoryEvent) -> dict:
        """Format a history event for display."""
        # Format changes as list of field changes
        formatted_changes = []
        if event.changes:
            for field_name, change in event.changes.items():
                formatted_changes.append({
                    "field": self._get_field_display_name(field_name),
                    "old": change.get("old", ""),
                    "new": change.get("new", ""),
                })
        
        return {
            "id": event.id,
            "time": event.event_time,
            "type": event.event_type,
            "type_display": self._get_event_type_display(event.event_type),
            "type_class": self._get_event_type_class(event.event_type),
            "user": event.user.username if event.user else "System",
            "changes": formatted_changes,
            "comment": event.comment or "",
            "raw_fragment": event.raw_history_fragment or "",
        }
    
    def _get_field_display_name(self, field_name: str) -> str:
        """Get human-readable field name."""
        field_names = {
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
            "extra1": "Zusatz 1",
            "extra2": "Zusatz 2",
            "extra3": "Zusatz 3",
        }
        
        # Month fields
        for i in range(1, 13):
            field_names[f"month_{i}"] = f"Monat {i}"
        
        return field_names.get(field_name, field_name)
    
    def _get_event_type_display(self, event_type: str) -> str:
        """Get display name for event type."""
        type_names = {
            "CHANGE": "Änderung",
            "CREATE": "Erstellt",
            "APPROVE": "Genehmigt",
            "DECLINE": "Abgelehnt",
            "IMPORT": "Importiert",
            "RUCK": "Rückwirkend",
        }
        return type_names.get(event_type, event_type)
    
    def _get_event_type_class(self, event_type: str) -> str:
        """Get CSS class for event type."""
        type_classes = {
            "CHANGE": "bg-info",
            "CREATE": "bg-success",
            "APPROVE": "bg-success",
            "DECLINE": "bg-danger",
            "IMPORT": "bg-secondary",
            "RUCK": "bg-warning",
        }
        return type_classes.get(event_type, "bg-secondary")
