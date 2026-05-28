"""
Views for catalog app.

This module contains views for managing:
- Teachers (Lehrer)
- Subjects (Fächer)
- Teaching Assignments (Zuweisungen)
- Price Options (Preise)
- Copying assignments and prices between years

Most catalog management is Admin-only. Some bulk pricing views are available
to Admin and Operator.
"""

from __future__ import annotations

from datetime import date
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db import IntegrityError
from django.db.models import Count, OuterRef, Q, QuerySet, Subquery
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse, reverse_lazy
from django.views.generic import (
    CreateView,
    DeleteView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
    View,
)

from .forms import (
    CopyCategoriesYearForm,
    CopyYearForm,
    CopyPricesYearForm,
    DurationEntryForm,
    GroupSizeEntryForm,
    PriceOptionForm,
    SubjectCategoryForm,
    SubjectForm,
    SyncFromLegacyForm,
    TeacherForm,
    TeachingAssignmentForm,
)
from .models import (
    CategoryKind,
    Discount,
    DisciplineGroup,
    DurationEntry,
    FamilyDiscount,
    GroupSizeEntry,
    PriceOption,
    RecordDiscount,
    Subject,
    SubjectCategory,
    SubjectCategoryLink,
    Teacher,
    TeachingAssignment,
    get_duration_for_month,
    get_manual_size_for_month,
)
from .group_size_service import get_group_size_for_month
from .pricing import calculate_suggested_price_group
from .services import copy_categories_between_years, ensure_default_categories
from .warnings import get_group_warnings



# =============================================================================
# Permission Mixin
# =============================================================================

class CatalogAdminMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to users with Admin role.
    
    Only Admin users can access catalog management views.
    """
    
    def test_func(self) -> bool:
        """Check if user has Admin role."""
        user = self.request.user
        return user.is_authenticated and user.is_admin_role
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to karteien list with error message."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        messages.error(
            self.request,
            "Sie haben keine Berechtigung, den Katalog zu verwalten. "
            "Nur Administratoren haben Zugriff."
        )
        return redirect("karteien:record_list")


class CatalogEditorMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that allows Admin and Operator roles.

    Used for pricing workflows where Operator is allowed to create PENDING
    changes but still has month restrictions.
    """

    def test_func(self) -> bool:
        user = self.request.user
        return user.is_authenticated and (user.is_admin_role or user.is_operator)

    def handle_no_permission(self) -> HttpResponse:
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()

        messages.error(
            self.request,
            "Sie haben keine Berechtigung für diese Aktion. "
            "Nur Admin und Operator haben Zugriff."
        )
        return redirect("karteien:record_list")


# =============================================================================
# Group Record Lookup Helpers
# =============================================================================

def _get_group_records_for_semester(
    group: DisciplineGroup,
    *,
    semester: int,
):
    """
    Return records for a group/semester with legacy subject-name fallback.

    Matches:
    - direct FK (subject*_ref_id == group.subject_id), or
    - legacy text (subject*), when subject*_ref is NULL and normalised names match.
    """
    from apps.karteien.billing import _normalize_subject_name
    from apps.karteien.models import KarteiRecord

    if semester not in (1, 2):
        return []

    if semester == 1:
        ref_field = "subject1_ref_id"
        legacy_field = "subject1"
        ref_null_field = "subject1_ref__isnull"
    else:
        ref_field = "subject2_ref_id"
        legacy_field = "subject2"
        ref_null_field = "subject2_ref__isnull"

    candidates = (
        KarteiRecord.objects
        .filter(year=group.year)
        .filter(
            Q(**{ref_field: group.subject_id})
            | (Q(**{ref_null_field: True}) & ~Q(**{legacy_field: ""}))
        )
    )

    norm_subject = _normalize_subject_name(group.subject.name)
    matched = []
    for rec in candidates.iterator():
        ref_value = getattr(rec, ref_field)
        if ref_value == group.subject_id:
            matched.append(rec)
            continue

        legacy_value = getattr(rec, legacy_field) or ""
        if _normalize_subject_name(legacy_value) == norm_subject:
            matched.append(rec)

    return matched


def _get_group_records(
    group: DisciplineGroup,
    *,
    semester: int | None = None,
):
    """Return group records for one semester or both semesters (deduplicated)."""
    if semester in (1, 2):
        return _get_group_records_for_semester(group, semester=semester)

    by_pk = {}
    for sem in (1, 2):
        for rec in _get_group_records_for_semester(group, semester=sem):
            by_pk[rec.pk] = rec
    return list(by_pk.values())


def _get_group_semester_stats(group: DisciplineGroup, *, semester: int) -> dict[str, int]:
    """Return readiness stats for group records in a semester."""
    from apps.karteien.models import MonthsMode, RecordStatus

    records = _get_group_records(group, semester=semester)
    ref_id_attr = "subject1_ref_id" if semester == 1 else "subject2_ref_id"

    total = len(records)
    normal = 0
    legacy = 0
    pending = 0
    declined = 0
    normal_legacy = 0
    normal_missing_ref = 0
    eligible_apply = 0

    for rec in records:
        status = rec.status
        if status == RecordStatus.NORMAL:
            normal += 1
            if rec.months_mode == MonthsMode.LEGACY:
                normal_legacy += 1
            if getattr(rec, ref_id_attr) is None:
                normal_missing_ref += 1
            if rec.months_mode != MonthsMode.LEGACY:
                eligible_apply += 1
        elif status == RecordStatus.PENDING:
            pending += 1
        elif status == RecordStatus.DECLINED:
            declined += 1

        if rec.months_mode == MonthsMode.LEGACY:
            legacy += 1

    return {
        "total": total,
        "normal": normal,
        "legacy": legacy,
        "pending": pending,
        "declined": declined,
        "normal_legacy": normal_legacy,
        "normal_missing_ref": normal_missing_ref,
        "eligible_apply": eligible_apply,
    }


# =============================================================================
# Index View
# =============================================================================

class CatalogIndexView(CatalogAdminMixin, TemplateView):
    """
    Catalog index page with quick links to all catalog sections.
    """
    
    template_name = "catalog/index.html"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["teacher_count"] = Teacher.objects.filter(is_active=True).count()
        context["subject_count"] = Subject.objects.filter(is_active=True).count()
        context["assignment_count"] = TeachingAssignment.objects.filter(is_active=True).count()
        context["price_count"] = PriceOption.objects.filter(is_active=True).count()
        context["discount_count"] = Discount.objects.filter(is_active=True).count()
        context["current_year"] = date.today().year
        return context


# =============================================================================
# Teacher Views
# =============================================================================

class TeacherListView(CatalogAdminMixin, ListView):
    """List all teachers with active/inactive filter."""
    
    model = Teacher
    template_name = "catalog/teachers_list.html"
    context_object_name = "teachers"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet[Teacher]:
        qs = Teacher.objects.all()
        
        # Filter by active status
        show_inactive = self.request.GET.get("show_inactive")
        if not show_inactive:
            qs = qs.filter(is_active=True)
        
        # Search by name
        search = self.request.GET.get("search", "").strip()
        if search:
            from django.db.models import Q
            qs = qs.filter(
                Q(last_name__icontains=search) |
                Q(first_name__icontains=search)
            )
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["show_inactive"] = self.request.GET.get("show_inactive", "")
        context["search"] = self.request.GET.get("search", "")
        return context


class TeacherCreateView(CatalogAdminMixin, CreateView):
    """Create a new teacher."""
    
    model = Teacher
    form_class = TeacherForm
    template_name = "catalog/teacher_form.html"
    success_url = reverse_lazy("catalog:teacher_list")

    def get_initial(self) -> dict[str, Any]:
        """Prefill form fields from query parameters.
        
        Parses 'full_name' query param as 'Nachname Vorname':
        - If >= 2 words: first_name = last token, last_name = rest
        - If 1 word: last_name = full_name, first_name = empty
        """
        initial = super().get_initial()
        full_name = self.request.GET.get("full_name", "")
        # Normalize whitespace: trim and collapse multiple spaces
        full_name = " ".join(full_name.split())
        if full_name:
            parts = full_name.split()
            if len(parts) >= 2:
                # Last token is first_name, rest is last_name
                initial["first_name"] = parts[-1]
                initial["last_name"] = " ".join(parts[:-1])
            else:
                # Single word: treat as last_name
                initial["last_name"] = full_name
        return initial

    def get_success_url(self) -> str:
        """Redirect to 'next' if provided, otherwise to teacher list."""
        next_url = self.request.POST.get("next") or self.request.GET.get("next", "")
        if next_url:
            return next_url
        return super().get_success_url()
    
    def form_valid(self, form):
        messages.success(self.request, "Lehrer wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Lehrer"
        context["next"] = self.request.GET.get("next", "")
        return context


class TeacherUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing teacher."""
    
    model = Teacher
    form_class = TeacherForm
    template_name = "catalog/teacher_form.html"
    success_url = reverse_lazy("catalog:teacher_list")
    
    def form_valid(self, form):
        messages.success(self.request, "Lehrer wurde erfolgreich aktualisiert.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Lehrer bearbeiten: {self.object}"
        return context


# =============================================================================
# Subject Views
# =============================================================================

class SubjectListView(CatalogAdminMixin, ListView):
    """List all subjects with active/inactive filter."""
    
    model = Subject
    template_name = "catalog/subjects_list.html"
    context_object_name = "subjects"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet[Subject]:
        qs = Subject.objects.all()
        
        # Filter by active status
        show_inactive = self.request.GET.get("show_inactive")
        if not show_inactive:
            qs = qs.filter(is_active=True)
        
        # Search by name
        search = self.request.GET.get("search", "").strip()
        if search:
            qs = qs.filter(name__icontains=search)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["show_inactive"] = self.request.GET.get("show_inactive", "")
        context["search"] = self.request.GET.get("search", "")
        return context


class SubjectCreateView(CatalogAdminMixin, CreateView):
    """Create a new subject."""
    
    model = Subject
    form_class = SubjectForm
    template_name = "catalog/subject_form.html"
    success_url = reverse_lazy("catalog:subject_list")

    def get_initial(self) -> dict[str, Any]:
        """Prefill form fields from query parameters."""
        initial = super().get_initial()
        name = self.request.GET.get("name", "").strip()
        if name:
            initial["name"] = name
        return initial

    def get_success_url(self) -> str:
        """Redirect to 'next' if provided, otherwise to subject list."""
        next_url = self.request.POST.get("next") or self.request.GET.get("next", "")
        if next_url:
            return next_url
        return super().get_success_url()
    
    def form_valid(self, form):
        messages.success(self.request, "Fach wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neues Fach"
        context["next"] = self.request.GET.get("next", "")
        return context


class SubjectUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing subject."""
    
    model = Subject
    form_class = SubjectForm
    template_name = "catalog/subject_form.html"
    success_url = reverse_lazy("catalog:subject_list")
    
    def form_valid(self, form):
        messages.success(self.request, "Fach wurde erfolgreich aktualisiert.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Fach bearbeiten: {self.object}"
        return context


# =============================================================================
# Teaching Assignment Views
# =============================================================================

class AssignmentListView(CatalogAdminMixin, ListView):
    """List teaching assignments with year/subject/teacher filters."""
    
    model = TeachingAssignment
    template_name = "catalog/assignments_list.html"
    context_object_name = "assignments"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet[TeachingAssignment]:
        qs = TeachingAssignment.objects.select_related("subject", "teacher")
        
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
        
        # Filter by subject
        subject_id = self.request.GET.get("subject")
        if subject_id:
            try:
                qs = qs.filter(subject_id=int(subject_id))
            except ValueError:
                pass
        
        # Filter by teacher
        teacher_id = self.request.GET.get("teacher")
        if teacher_id:
            try:
                qs = qs.filter(teacher_id=int(teacher_id))
            except ValueError:
                pass
        
        # Filter by active status
        show_inactive = self.request.GET.get("show_inactive")
        if not show_inactive:
            qs = qs.filter(is_active=True)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        
        # Current filters
        context["current_year"] = self.request.GET.get("year", str(date.today().year))
        context["current_subject"] = self.request.GET.get("subject", "")
        context["current_teacher"] = self.request.GET.get("teacher", "")
        context["show_inactive"] = self.request.GET.get("show_inactive", "")
        
        # Available years for dropdown
        years = TeachingAssignment.objects.values_list("year", flat=True).distinct()
        current_year = date.today().year
        year_list = sorted(set(years) | {current_year, current_year + 1}, reverse=True)
        context["available_years"] = year_list
        
        # Available subjects and teachers for filtering
        context["subjects"] = Subject.objects.filter(is_active=True)
        context["teachers"] = Teacher.objects.filter(is_active=True)
        
        return context


class AssignmentCreateView(CatalogAdminMixin, CreateView):
    """Create a new teaching assignment."""
    
    model = TeachingAssignment
    form_class = TeachingAssignmentForm
    template_name = "catalog/assignment_form.html"
    success_url = reverse_lazy("catalog:assignment_list")
    
    def form_valid(self, form):
        try:
            response = super().form_valid(form)
            messages.success(self.request, "Zuweisung wurde erfolgreich erstellt.")
            return response
        except IntegrityError:
            messages.error(
                self.request,
                "Diese Zuweisung existiert bereits (gleicher Lehrer, Fach und Jahr)."
            )
            return self.form_invalid(form)

    def get_success_url(self) -> str:
        """Redirect to 'next' if provided, otherwise to assignment list."""
        next_url = self.request.POST.get("next") or self.request.GET.get("next", "")
        if next_url:
            return next_url
        return super().get_success_url()
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neue Zuweisung"
        context["next"] = self.request.GET.get("next", "")
        return context
    
    def get_initial(self):
        """Prefill form fields from query parameters."""
        initial = super().get_initial()
        # Pre-fill year from query parameter
        year = self.request.GET.get("year")
        if year:
            try:
                initial["year"] = int(year)
            except ValueError:
                pass
        # Pre-fill subject from query parameter
        subject_id = self.request.GET.get("subject")
        if subject_id:
            try:
                initial["subject"] = int(subject_id)
            except ValueError:
                pass
        # Pre-fill teacher from query parameter
        teacher_id = self.request.GET.get("teacher")
        if teacher_id:
            try:
                initial["teacher"] = int(teacher_id)
            except ValueError:
                pass
        return initial


class AssignmentDeleteView(CatalogAdminMixin, DeleteView):
    """Delete a teaching assignment."""
    
    model = TeachingAssignment
    success_url = reverse_lazy("catalog:assignment_list")
    
    def post(self, request, *args, **kwargs):
        assignment = self.get_object()
        assignment_str = str(assignment)
        response = super().post(request, *args, **kwargs)
        messages.success(request, f"Zuweisung '{assignment_str}' wurde gelöscht.")
        return response


# =============================================================================
# Copy Year View
# =============================================================================

class CopyYearView(CatalogAdminMixin, FormView):
    """
    Copy teaching assignments from one year to another.
    
    Features:
    - Copy only active assignments (optional)
    - Overwrite existing assignments in target year (optional)
    - Idempotent operation when overwrite=False
    """
    
    template_name = "catalog/copy_year.html"
    form_class = CopyYearForm
    success_url = reverse_lazy("catalog:assignment_list")
    
    def form_valid(self, form):
        from_year = form.cleaned_data["from_year"]
        to_year = form.cleaned_data["to_year"]
        only_active = form.cleaned_data["only_active"]
        overwrite = form.cleaned_data["overwrite"]
        
        # Get source assignments
        source_qs = TeachingAssignment.objects.filter(year=from_year)
        if only_active:
            source_qs = source_qs.filter(is_active=True)
        
        source_assignments = list(source_qs.select_related("subject", "teacher"))
        
        if not source_assignments:
            messages.warning(
                self.request,
                f"Keine Zuweisungen im Jahr {from_year} gefunden."
            )
            return super().form_valid(form)
        
        # Handle overwrite
        deactivated_count = 0
        if overwrite:
            # Deactivate existing assignments in target year (safer than deleting)
            deactivated_count = TeachingAssignment.objects.filter(
                year=to_year,
                is_active=True
            ).update(is_active=False)
        
        # Copy assignments
        created_count = 0
        skipped_count = 0
        
        for assignment in source_assignments:
            # Check if assignment already exists in target year
            existing = TeachingAssignment.objects.filter(
                year=to_year,
                subject=assignment.subject,
                teacher=assignment.teacher,
            ).first()
            
            if existing:
                if overwrite and not existing.is_active:
                    # Reactivate if it was deactivated
                    existing.is_active = True
                    existing.save(update_fields=["is_active"])
                    created_count += 1
                else:
                    skipped_count += 1
            else:
                # Create new assignment
                TeachingAssignment.objects.create(
                    year=to_year,
                    subject=assignment.subject,
                    teacher=assignment.teacher,
                    is_active=True,
                )
                created_count += 1
        
        # Build success message
        msg_parts = [f"Kopieren von {from_year} nach {to_year} abgeschlossen."]
        if created_count > 0:
            msg_parts.append(f"{created_count} Zuweisungen erstellt/aktiviert.")
        if skipped_count > 0:
            msg_parts.append(f"{skipped_count} bereits vorhanden (übersprungen).")
        if deactivated_count > 0:
            msg_parts.append(f"{deactivated_count} vorher deaktiviert.")
        
        messages.success(self.request, " ".join(msg_parts))
        
        # Redirect to assignment list for target year
        return redirect(f"{self.success_url}?year={to_year}")
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        
        # Available years with assignment counts
        from django.db.models import Count, Q
        year_stats = (
            TeachingAssignment.objects
            .values("year")
            .annotate(total=Count("id"), active=Count("id", filter=Q(is_active=True)))
            .order_by("-year")
        )
        context["year_stats"] = list(year_stats)
        context["current_year"] = date.today().year
        
        return context


# =============================================================================
# Price Option Views
# =============================================================================

class PriceListView(CatalogAdminMixin, ListView):
    """List price options with year/subject filters."""
    
    model = PriceOption
    template_name = "catalog/prices_list.html"
    context_object_name = "prices"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet[PriceOption]:
        qs = PriceOption.objects.select_related("subject")
        
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
        
        # Filter by subject
        subject_id = self.request.GET.get("subject")
        if subject_id:
            try:
                qs = qs.filter(subject_id=int(subject_id))
            except ValueError:
                pass
        
        # Filter by active status
        show_inactive = self.request.GET.get("show_inactive")
        if not show_inactive:
            qs = qs.filter(is_active=True)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        
        # Current filters
        context["current_year"] = self.request.GET.get("year", str(date.today().year))
        context["current_subject"] = self.request.GET.get("subject", "")
        context["show_inactive"] = self.request.GET.get("show_inactive", "")
        
        # Available years for dropdown
        years = PriceOption.objects.values_list("year", flat=True).distinct()
        current_year = date.today().year
        year_list = sorted(set(years) | {current_year, current_year + 1}, reverse=True)
        context["available_years"] = year_list
        
        # Available subjects for filtering
        context["subjects"] = Subject.objects.filter(is_active=True)
        
        return context


class PriceCreateView(CatalogAdminMixin, CreateView):
    """Create a new price option."""
    
    model = PriceOption
    form_class = PriceOptionForm
    template_name = "catalog/price_form.html"
    success_url = reverse_lazy("catalog:price_list")
    
    def form_valid(self, form):
        try:
            response = super().form_valid(form)
            messages.success(self.request, "Preis wurde erfolgreich erstellt.")
            return response
        except IntegrityError:
            messages.error(
                self.request,
                "Dieser Preis existiert bereits (gleicher Betrag, Fach und Jahr)."
            )
            return self.form_invalid(form)

    def get_form_kwargs(self) -> dict[str, Any]:
        """Pass require_comment flag to form."""
        kwargs = super().get_form_kwargs()
        # Check if require_comment=1 in query params
        if self.request.GET.get("require_comment") == "1":
            kwargs["require_comment"] = True
        return kwargs

    def get_success_url(self) -> str:
        """Redirect to 'next' if provided, otherwise to price list."""
        next_url = self.request.POST.get("next") or self.request.GET.get("next", "")
        if next_url:
            return next_url
        return super().get_success_url()
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        from decimal import Decimal, InvalidOperation
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Preis"
        context["next"] = self.request.GET.get("next", "")
        context["require_comment"] = self.request.GET.get("require_comment") == "1"
        
        # Find duplicate prices with same year+subject+amount for hint
        year = self.request.GET.get("year")
        subject_id = self.request.GET.get("subject")
        amount_str = self.request.GET.get("amount")
        
        existing_prices = []
        if year and subject_id and amount_str:
            try:
                year_int = int(year)
                subject_int = int(subject_id)
                amount_dec = Decimal(amount_str)
                existing_prices = list(
                    PriceOption.objects.filter(
                        year=year_int,
                        subject_id=subject_int,
                        amount=amount_dec,
                        is_active=True,
                    ).select_related("subject")
                )
            except (ValueError, InvalidOperation):
                pass
        context["existing_prices"] = existing_prices
        
        return context
    
    def get_initial(self):
        """Prefill form fields from query parameters."""
        from decimal import Decimal, InvalidOperation
        initial = super().get_initial()
        # Pre-fill year from query parameter
        year = self.request.GET.get("year")
        if year:
            try:
                initial["year"] = int(year)
            except ValueError:
                pass
        # Pre-fill subject from query parameter
        subject = self.request.GET.get("subject")
        if subject:
            try:
                initial["subject"] = int(subject)
            except ValueError:
                pass
        # Pre-fill amount from query parameter
        amount = self.request.GET.get("amount")
        if amount:
            try:
                initial["amount"] = Decimal(amount)
            except InvalidOperation:
                pass
        return initial


class PriceUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing price option."""
    
    model = PriceOption
    form_class = PriceOptionForm
    template_name = "catalog/price_form.html"
    success_url = reverse_lazy("catalog:price_list")
    
    def form_valid(self, form):
        try:
            response = super().form_valid(form)
            messages.success(self.request, "Preis wurde erfolgreich aktualisiert.")
            return response
        except IntegrityError:
            messages.error(
                self.request,
                "Dieser Preis existiert bereits (gleicher Betrag, Fach und Jahr)."
            )
            return self.form_invalid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Preis bearbeiten: {self.object}"
        return context
    
    def get_success_url(self):
        # Redirect back with year filter preserved
        year = self.object.year
        return f"{reverse_lazy('catalog:price_list')}?year={year}"


class PriceDeleteView(CatalogAdminMixin, DeleteView):
    """Delete a price option."""
    
    model = PriceOption
    success_url = reverse_lazy("catalog:price_list")
    
    def post(self, request, *args, **kwargs):
        price = self.get_object()
        year = price.year
        price_str = str(price)
        response = super().post(request, *args, **kwargs)
        messages.success(request, f"Preis '{price_str}' wurde gelöscht.")
        # Redirect with year filter
        return redirect(f"{self.success_url}?year={year}")


# =============================================================================
# Copy Prices Year View
# =============================================================================

class CopyPricesYearView(CatalogAdminMixin, FormView):
    """
    Copy price options from one year to another.
    
    Features:
    - Copy only active prices (optional)
    - Overwrite existing prices in target year (optional)
    - Idempotent operation when overwrite=False
    """
    
    template_name = "catalog/copy_prices_year.html"
    form_class = CopyPricesYearForm
    success_url = reverse_lazy("catalog:price_list")
    
    def form_valid(self, form):
        from_year = form.cleaned_data["from_year"]
        to_year = form.cleaned_data["to_year"]
        only_active = form.cleaned_data["only_active"]
        overwrite = form.cleaned_data["overwrite"]
        
        # Get source prices
        source_qs = PriceOption.objects.filter(year=from_year)
        if only_active:
            source_qs = source_qs.filter(is_active=True)
        
        source_prices = list(source_qs.select_related("subject"))
        
        if not source_prices:
            messages.warning(
                self.request,
                f"Keine Preise im Jahr {from_year} gefunden."
            )
            return super().form_valid(form)
        
        # Handle overwrite
        deactivated_count = 0
        if overwrite:
            # Deactivate existing prices in target year (safer than deleting)
            deactivated_count = PriceOption.objects.filter(
                year=to_year,
                is_active=True
            ).update(is_active=False)
        
        # Copy prices
        created_count = 0
        skipped_count = 0
        updated_count = 0
        
        for price in source_prices:
            # Check if price already exists in target year
            existing = PriceOption.objects.filter(
                year=to_year,
                subject=price.subject,
                amount=price.amount,
                comment=price.comment,
            ).first()
            
            if existing:
                if overwrite and not existing.is_active:
                    # Reactivate if it was deactivated
                    existing.is_active = True
                    existing.save(update_fields=["is_active"])
                    updated_count += 1
                else:
                    skipped_count += 1
            else:
                # Create new price
                PriceOption.objects.create(
                    year=to_year,
                    subject=price.subject,
                    amount=price.amount,
                    comment=price.comment,
                    is_active=True,
                )
                created_count += 1
        
        # Build success message
        msg_parts = [f"Kopieren von {from_year} nach {to_year} abgeschlossen."]
        if created_count > 0:
            msg_parts.append(f"{created_count} Preise erstellt.")
        if updated_count > 0:
            msg_parts.append(f"{updated_count} Preise reaktiviert.")
        if skipped_count > 0:
            msg_parts.append(f"{skipped_count} bereits vorhanden (übersprungen).")
        if deactivated_count > 0:
            msg_parts.append(f"{deactivated_count} vorher deaktiviert.")
        
        messages.success(self.request, " ".join(msg_parts))
        
        # Redirect to price list for target year
        return redirect(f"{self.success_url}?year={to_year}")
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        
        # Available years with price counts
        from django.db.models import Count, Q
        year_stats = (
            PriceOption.objects
            .values("year")
            .annotate(total=Count("id"), active=Count("id", filter=Q(is_active=True)))
            .order_by("-year")
        )
        context["year_stats"] = list(year_stats)
        context["current_year"] = date.today().year
        
        return context


# =============================================================================
# Discount Views
# =============================================================================

class DiscountListView(CatalogAdminMixin, ListView):
    """List all discounts with filtering."""
    
    model = Discount
    template_name = "catalog/discounts_list.html"
    context_object_name = "discounts"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet:
        from .models import Discount
        qs = Discount.objects.all()
        
        # Filter by active status
        show_inactive = self.request.GET.get("show_inactive")
        if not show_inactive:
            qs = qs.filter(is_active=True)
        
        # Filter by kind
        kind = self.request.GET.get("kind", "").strip()
        if kind:
            qs = qs.filter(kind=kind)
        
        # Search by description
        search = self.request.GET.get("search", "").strip()
        if search:
            qs = qs.filter(description__icontains=search)
        
        return qs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        from .models import DiscountKind
        context = super().get_context_data(**kwargs)
        context["show_inactive"] = self.request.GET.get("show_inactive", "")
        context["kind"] = self.request.GET.get("kind", "")
        context["search"] = self.request.GET.get("search", "")
        context["kind_choices"] = DiscountKind.choices
        return context


class DiscountCreateView(CatalogAdminMixin, CreateView):
    """Create a new discount."""
    
    model = Discount
    template_name = "catalog/discount_form.html"
    success_url = reverse_lazy("catalog:discount_list")
    
    def get_form_class(self):
        from .forms import DiscountForm
        return DiscountForm
    
    def form_valid(self, form):
        messages.success(self.request, "Rabatt wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Rabatt"
        return context


class DiscountUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing discount."""
    
    model = Discount
    template_name = "catalog/discount_form.html"
    success_url = reverse_lazy("catalog:discount_list")
    
    def get_form_class(self):
        from .forms import DiscountForm
        return DiscountForm
    
    def form_valid(self, form):
        messages.success(self.request, "Rabatt wurde erfolgreich aktualisiert.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Rabatt bearbeiten: {self.object}"
        return context


# =============================================================================
# Family Discount Views
# =============================================================================

class FamilyDiscountListView(CatalogAdminMixin, ListView):
    """List family discounts with year and family filtering."""
    
    model = FamilyDiscount
    template_name = "catalog/family_discounts_list.html"
    context_object_name = "family_discounts"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet:
        from .models import FamilyDiscount
        from apps.karteien.models import KarteiRecord
        
        # Subquery to get parent_name from any KarteiRecord with same (year, family_id)
        parent_subquery = KarteiRecord.objects.filter(
            year=OuterRef('year'),
            family_id=OuterRef('family_id')
        ).order_by('pkid').values('parent_name')[:1]
        
        qs = FamilyDiscount.objects.select_related('discount').annotate(
            parent_name=Subquery(parent_subquery)
        )
        
        # Filter by year
        year = self.request.GET.get("year", "").strip()
        if year:
            try:
                qs = qs.filter(year=int(year))
            except ValueError:
                pass
        
        # Filter by family_id
        family_id = self.request.GET.get("family_id", "").strip()
        if family_id:
            qs = qs.filter(family_id__icontains=family_id)
        
        return qs.order_by('-year', 'family_id', 'start_month')
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        from .models import FamilyDiscount
        context = super().get_context_data(**kwargs)
        context["year"] = self.request.GET.get("year", "")
        context["family_id"] = self.request.GET.get("family_id", "")
        context["current_year"] = date.today().year
        
        # Available years
        years = FamilyDiscount.objects.values_list('year', flat=True).distinct().order_by('-year')
        context["available_years"] = list(years)
        
        return context


class FamilyDiscountCreateView(CatalogAdminMixin, CreateView):
    """Create a new family discount."""
    
    model = FamilyDiscount
    template_name = "catalog/family_discount_form.html"
    success_url = reverse_lazy("catalog:family_discount_list")
    
    def get_form_class(self):
        from .forms import FamilyDiscountForm
        return FamilyDiscountForm
    
    def get_initial(self) -> dict[str, Any]:
        initial = super().get_initial()
        # Pre-fill from query params
        if self.request.GET.get("year"):
            try:
                initial["year"] = int(self.request.GET.get("year"))
            except ValueError:
                pass
        if self.request.GET.get("family_id"):
            initial["family_id"] = self.request.GET.get("family_id")
        return initial
    
    def form_valid(self, form):
        messages.success(self.request, "Familienrabatt wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_success_url(self):
        # Redirect back with year filter
        year = self.object.year
        return f"{reverse_lazy('catalog:family_discount_list')}?year={year}"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Familienrabatt"
        return context


class FamilyDiscountUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing family discount."""
    
    model = FamilyDiscount
    template_name = "catalog/family_discount_form.html"
    success_url = reverse_lazy("catalog:family_discount_list")
    
    def get_form_class(self):
        from .forms import FamilyDiscountForm
        return FamilyDiscountForm
    
    def form_valid(self, form):
        messages.success(self.request, "Familienrabatt wurde erfolgreich aktualisiert.")
        return super().form_valid(form)
    
    def get_success_url(self):
        year = self.object.year
        return f"{reverse_lazy('catalog:family_discount_list')}?year={year}"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Familienrabatt bearbeiten: {self.object.family_id} ({self.object.year})"
        return context


class FamilyDiscountDeleteView(CatalogAdminMixin, DeleteView):
    """Delete a family discount."""
    
    model = FamilyDiscount
    template_name = "catalog/family_discount_confirm_delete.html"
    success_url = reverse_lazy("catalog:family_discount_list")
    
    def form_valid(self, form):
        year = self.object.year
        messages.success(self.request, "Familienrabatt wurde erfolgreich gelöscht.")
        response = super().form_valid(form)
        return redirect(f"{self.success_url}?year={year}")


# =============================================================================
# Record Discount Views
# =============================================================================

class RecordDiscountListView(CatalogAdminMixin, ListView):
    """List record discounts with filtering."""
    
    model = RecordDiscount
    template_name = "catalog/record_discounts_list.html"
    context_object_name = "record_discounts"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet:
        from .models import RecordDiscount
        qs = RecordDiscount.objects.select_related('discount', 'record').all()
        
        # Filter by record PKID
        record_pk = self.request.GET.get("record_pk", "").strip()
        if record_pk:
            try:
                qs = qs.filter(record__pkid=int(record_pk))
            except ValueError:
                pass
        
        # Filter by year (through record)
        year = self.request.GET.get("year", "").strip()
        if year:
            try:
                qs = qs.filter(record__year=int(year))
            except ValueError:
                pass
        
        # Filter by family_id (through record)
        family_id = self.request.GET.get("family_id", "").strip()
        if family_id:
            qs = qs.filter(record__family_id__icontains=family_id)
        
        return qs.order_by('-record__year', 'record__pkid', 'start_month')
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["record_pk"] = self.request.GET.get("record_pk", "")
        context["year"] = self.request.GET.get("year", "")
        context["family_id"] = self.request.GET.get("family_id", "")
        context["current_year"] = date.today().year
        # Pass next parameter for return navigation
        context["next"] = self.request.GET.get("next", "")
        return context


class RecordDiscountCreateView(CatalogAdminMixin, CreateView):
    """Create a new record discount."""
    
    model = RecordDiscount
    template_name = "catalog/record_discount_form.html"
    success_url = reverse_lazy("catalog:record_discount_list")
    
    def get_form_class(self):
        from .forms import RecordDiscountForm
        return RecordDiscountForm
    
    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        # Pass record if provided in query params
        record_pk = self.request.GET.get("record_pk")
        if record_pk:
            from apps.karteien.models import KarteiRecord
            try:
                kwargs['record'] = KarteiRecord.objects.get(pkid=int(record_pk))
            except (ValueError, KarteiRecord.DoesNotExist):
                pass
        return kwargs
    
    def form_valid(self, form):
        messages.success(self.request, "Eintrag-Rabatt wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_success_url(self):
        # Check for next parameter first
        next_url = self.request.GET.get("next") or self.request.POST.get("next", "")
        if next_url:
            return next_url
        record_pk = self.object.record.pkid
        return f"{reverse_lazy('catalog:record_discount_list')}?record_pk={record_pk}"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Eintrag-Rabatt"
        
        # Show record info if provided
        record_pk = self.request.GET.get("record_pk")
        if record_pk:
            from apps.karteien.models import KarteiRecord
            try:
                context["target_record"] = KarteiRecord.objects.get(pkid=int(record_pk))
            except (ValueError, KarteiRecord.DoesNotExist):
                pass
        
        # Pass next parameter for form and back navigation
        context["next"] = self.request.GET.get("next", "")
        
        return context


class RecordDiscountUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing record discount."""
    
    model = RecordDiscount
    template_name = "catalog/record_discount_form.html"
    success_url = reverse_lazy("catalog:record_discount_list")
    
    def get_form_class(self):
        from .forms import RecordDiscountForm
        return RecordDiscountForm
    
    def form_valid(self, form):
        messages.success(self.request, "Eintrag-Rabatt wurde erfolgreich aktualisiert.")
        return super().form_valid(form)
    
    def get_success_url(self):
        # Check for next parameter first
        next_url = self.request.GET.get("next") or self.request.POST.get("next", "")
        if next_url:
            return next_url
        record_pk = self.object.record.pkid
        return f"{reverse_lazy('catalog:record_discount_list')}?record_pk={record_pk}"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Eintrag-Rabatt bearbeiten: #{self.object.record.pkid}"
        context["target_record"] = self.object.record
        # Pass next parameter for form and back navigation
        context["next"] = self.request.GET.get("next", "")
        return context


class RecordDiscountDeleteView(CatalogAdminMixin, DeleteView):
    """Delete a record discount."""
    
    model = RecordDiscount
    template_name = "catalog/record_discount_confirm_delete.html"
    success_url = reverse_lazy("catalog:record_discount_list")
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        # Pass next parameter for back navigation
        context["next"] = self.request.GET.get("next", "")
        return context
    
    def form_valid(self, form):
        record_pk = self.object.record.pkid
        # Check for next parameter first
        next_url = self.request.GET.get("next") or self.request.POST.get("next", "")
        messages.success(self.request, "Eintrag-Rabatt wurde erfolgreich gelöscht.")
        response = super().form_valid(form)
        if next_url:
            return redirect(next_url)
        return redirect(f"{self.success_url}?record_pk={record_pk}")


# =============================================================================
# FamilyID Reservation Views (Admin only)
# =============================================================================

from django.db import transaction
from django.views import View
from apps.familyid_reservations.models import FamilyIdReservation
from apps.familyid_reservations.services import get_next_family_id


class FamilyIdReservationListView(CatalogAdminMixin, ListView):
    """
    List active (unused) FamilyID reservations.
    
    Shows all reserved FamilyIDs that haven't been used yet,
    with options to reserve a new one or cancel existing.
    """
    
    model = FamilyIdReservation
    template_name = "catalog/familyid_reservations.html"
    context_object_name = "reservations"
    paginate_by = 50
    
    def get_queryset(self) -> QuerySet:
        """Return active reservations, ordered by newest first."""
        return FamilyIdReservation.objects.filter(
            is_used=False
        ).select_related('reserved_by').order_by('-reserved_at')
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["next_family_id_preview"] = get_next_family_id()
        return context


class ReserveNextFamilyIdView(CatalogAdminMixin, View):
    """
    Reserve the next available FamilyID.
    
    Uses atomic transaction with retry logic to handle race conditions
    when multiple admins try to reserve at the same time.
    """
    
    MAX_RETRIES = 3
    
    def post(self, request):
        """Handle POST request to reserve next FamilyID."""
        for attempt in range(self.MAX_RETRIES):
            try:
                with transaction.atomic():
                    candidate = get_next_family_id()
                    reservation = FamilyIdReservation.objects.create(
                        family_id=candidate,
                        reserved_by=request.user
                    )
                    messages.success(
                        request,
                        f"FamilyID '{reservation.family_id}' wurde erfolgreich reserviert."
                    )
                    return redirect("catalog:familyid_reservation_list")
            except IntegrityError:
                # Race condition: another admin reserved this ID
                # Retry with new candidate
                if attempt < self.MAX_RETRIES - 1:
                    continue
                messages.error(
                    request,
                    "Reservierung fehlgeschlagen: Bitte versuchen Sie es erneut."
                )
                return redirect("catalog:familyid_reservation_list")
        
        # Should not reach here, but safety fallback
        messages.error(request, "Reservierung fehlgeschlagen nach mehreren Versuchen.")
        return redirect("catalog:familyid_reservation_list")
    
    def get(self, request):
        """GET not allowed, redirect to list."""
        return redirect("catalog:familyid_reservation_list")


class CancelFamilyIdReservationView(CatalogAdminMixin, View):
    """
    Cancel (delete) an existing FamilyID reservation.
    
    Only unused reservations can be cancelled.
    """
    
    def post(self, request, pk):
        """Handle POST request to cancel a reservation."""
        from django.shortcuts import get_object_or_404
        
        reservation = get_object_or_404(FamilyIdReservation, pk=pk)
        
        if reservation.is_used:
            messages.error(
                request,
                f"Reservierung '{reservation.family_id}' kann nicht storniert werden, "
                "da sie bereits verwendet wurde."
            )
            return redirect("catalog:familyid_reservation_list")
        
        family_id = reservation.family_id
        reservation.delete()
        
        messages.success(
            request,
            f"Reservierung '{family_id}' wurde erfolgreich storniert."
        )
        return redirect("catalog:familyid_reservation_list")
    
    def get(self, request, pk):
        """GET not allowed, redirect to list."""
        return redirect("catalog:familyid_reservation_list")


# =============================================================================
# Sync From Legacy View
# =============================================================================

class SyncFromLegacyView(CatalogAdminMixin, FormView):
    """
    Sync catalog entries from legacy KarteiRecord text fields.
    
    Scans legacy fields (subject1, subject2, extra1-3, teacher1/2_legacy_name)
    for a given year and adds missing entries to the catalog.
    
    Features:
    - Preview of missing subjects, teachers, and assignments
    - Idempotent: repeated runs on the same year don't create duplicates
    - Handles unparseable teacher names gracefully
    """
    
    template_name = "catalog/sync_from_legacy.html"
    form_class = SyncFromLegacyForm
    success_url = reverse_lazy("catalog:sync_from_legacy")
    
    def _normalize_name(self, name: str) -> str:
        """
        Normalize a subject/teacher name for matching.
        
        - Strips leading/trailing whitespace
        - Collapses multiple whitespace to single space
        - Returns casefolded version for comparison
        """
        import re
        if not name:
            return ""
        # Strip and collapse whitespace
        name = name.strip()
        name = re.sub(r"\s+", " ", name)
        return name.casefold()
    
    def _clean_name(self, name: str) -> str:
        """
        Clean a name for storage (preserves original case).
        
        - Strips leading/trailing whitespace
        - Collapses multiple whitespace to single space
        """
        import re
        if not name:
            return ""
        name = name.strip()
        name = re.sub(r"\s+", " ", name)
        return name
    
    def _parse_teacher_name(self, full_name: str) -> tuple[str, str] | None:
        """
        Parse a legacy teacher name into (last_name, first_name).
        
        Supports two formats:
        - "Nachname Vorname" -> last_name = all but last token, first_name = last token
        - "Nachname, Vorname" -> split by first comma
        
        Returns None if unparseable (single word or empty).
        """
        if not full_name:
            return None
        
        # Clean and normalize whitespace
        cleaned = self._clean_name(full_name)
        if not cleaned:
            return None
        
        # Check for comma format: "Nachname, Vorname"
        if "," in cleaned:
            parts = cleaned.split(",", 1)
            if len(parts) == 2:
                last_name = parts[0].strip()
                first_name = parts[1].strip()
                if last_name and first_name:
                    return (last_name, first_name)
        
        # Space format: "Nachname Vorname" (last token is first_name)
        parts = cleaned.split()
        if len(parts) >= 2:
            first_name = parts[-1]
            last_name = " ".join(parts[:-1])
            return (last_name, first_name)
        
        # Single word or empty - cannot parse
        return None
    
    def _get_legacy_data(self, year: int) -> dict:
        """
        Extract unique subjects, teachers, and assignments from legacy data for a year.
        
        Returns a dict with:
        - legacy_subjects: set of unique cleaned subject names
        - legacy_teachers: dict mapping (last_name, first_name) -> set of original full names
        - legacy_assignments: set of (subject_normalized, last_name, first_name) tuples
        - unparsed_teachers: set of teacher names that couldn't be parsed
        """
        from apps.karteien.models import KarteiRecord
        
        records = KarteiRecord.objects.filter(year=year).values(
            "subject1", "subject2", "extra1", "extra2", "extra3",
            "teacher1_legacy_name", "teacher2_legacy_name"
        )
        
        legacy_subjects = set()
        legacy_teachers = {}  # (last_name, first_name) -> set of original names
        legacy_assignments = set()  # (subject_norm, last_name, first_name)
        unparsed_teachers = set()
        # Map normalized subject name to human-readable cleaned name (for preview display)
        subject_display_by_norm = {}
        
        for rec in records:
            # Collect all subjects (including extras)
            for field in ["subject1", "subject2", "extra1", "extra2", "extra3"]:
                val = rec.get(field, "") or ""
                cleaned = self._clean_name(val)
                if cleaned:
                    legacy_subjects.add(cleaned)
                    # Store display name mapping (first occurrence wins)
                    norm = self._normalize_name(cleaned)
                    if norm and norm not in subject_display_by_norm:
                        subject_display_by_norm[norm] = cleaned
            
            # Process semester 1: subject1 + teacher1
            subj1 = self._clean_name(rec.get("subject1", "") or "")
            teacher1_raw = (rec.get("teacher1_legacy_name", "") or "").strip()
            if teacher1_raw:
                parsed = self._parse_teacher_name(teacher1_raw)
                if parsed:
                    last_name, first_name = parsed
                    key = (last_name, first_name)
                    if key not in legacy_teachers:
                        legacy_teachers[key] = set()
                    legacy_teachers[key].add(teacher1_raw)
                    # Assignment: only if subject is also non-empty
                    if subj1:
                        legacy_assignments.add((self._normalize_name(subj1), last_name, first_name))
                else:
                    unparsed_teachers.add(teacher1_raw)
            
            # Process semester 2: subject2 + teacher2
            subj2 = self._clean_name(rec.get("subject2", "") or "")
            teacher2_raw = (rec.get("teacher2_legacy_name", "") or "").strip()
            if teacher2_raw:
                parsed = self._parse_teacher_name(teacher2_raw)
                if parsed:
                    last_name, first_name = parsed
                    key = (last_name, first_name)
                    if key not in legacy_teachers:
                        legacy_teachers[key] = set()
                    legacy_teachers[key].add(teacher2_raw)
                    # Assignment: only if subject is also non-empty
                    if subj2:
                        legacy_assignments.add((self._normalize_name(subj2), last_name, first_name))
                else:
                    unparsed_teachers.add(teacher2_raw)
        
        return {
            "legacy_subjects": legacy_subjects,
            "legacy_teachers": legacy_teachers,
            "legacy_assignments": legacy_assignments,
            "unparsed_teachers": unparsed_teachers,
            "subject_display_by_norm": subject_display_by_norm,
        }
    
    def _find_missing_entries(self, year: int, legacy_data: dict) -> dict:
        """
        Compare legacy data against catalog and find missing entries.
        
        Returns dict with:
        - missing_subjects: list of cleaned subject names not in catalog
        - missing_teachers: list of (last_name, first_name) tuples not in catalog
        - missing_assignments: list of (subject_name, last_name, first_name) tuples
        - existing_subjects_map: dict of normalized_name -> Subject
        - existing_teachers_map: dict of (last_name, first_name) -> Teacher
        """
        legacy_subjects = legacy_data["legacy_subjects"]
        legacy_teachers = legacy_data["legacy_teachers"]
        legacy_assignments = legacy_data["legacy_assignments"]
        
        # Build map of existing subjects by normalized name
        existing_subjects = Subject.objects.all()
        existing_subjects_map = {}
        for subj in existing_subjects:
            norm = self._normalize_name(subj.name)
            existing_subjects_map[norm] = subj
        
        # Find missing subjects
        missing_subjects = []
        for subj_name in sorted(legacy_subjects):
            norm = self._normalize_name(subj_name)
            if norm and norm not in existing_subjects_map:
                missing_subjects.append(subj_name)
        
        # Build map of existing teachers by (last_name, first_name)
        existing_teachers = Teacher.objects.all()
        existing_teachers_map = {}
        for teacher in existing_teachers:
            key = (teacher.last_name, teacher.first_name)
            existing_teachers_map[key] = teacher
        
        # Find missing teachers
        missing_teachers = []
        for key in sorted(legacy_teachers.keys()):
            if key not in existing_teachers_map:
                missing_teachers.append(key)
        
        # Find missing assignments
        existing_assignments = TeachingAssignment.objects.filter(year=year).select_related(
            "subject", "teacher"
        )
        existing_assignment_set = set()
        for assign in existing_assignments:
            subj_norm = self._normalize_name(assign.subject.name)
            key = (subj_norm, assign.teacher.last_name, assign.teacher.first_name)
            existing_assignment_set.add(key)
        
        missing_assignments = []
        for assign_key in sorted(legacy_assignments):
            if assign_key not in existing_assignment_set:
                # Only add if we can find/create both subject and teacher
                subj_norm, last_name, first_name = assign_key
                missing_assignments.append((subj_norm, last_name, first_name))
        
        return {
            "missing_subjects": missing_subjects,
            "missing_teachers": missing_teachers,
            "missing_assignments": missing_assignments,
            "existing_subjects_map": existing_subjects_map,
            "existing_teachers_map": existing_teachers_map,
        }
    
    def _sync_entries(self, year: int, legacy_data: dict, missing_data: dict) -> dict:
        """
        Create missing catalog entries.
        
        Order: Subjects first, then Teachers, then Assignments.
        
        Returns stats dict with counts and any errors.
        """
        subjects_created = 0
        teachers_created = 0
        assignments_created = 0
        errors = []
        
        # Get mutable copies of existing maps
        subjects_map = dict(missing_data["existing_subjects_map"])
        teachers_map = dict(missing_data["existing_teachers_map"])
        
        # 1. Create missing subjects
        for subj_name in missing_data["missing_subjects"]:
            norm = self._normalize_name(subj_name)
            if norm in subjects_map:
                continue  # Already exists or was just created
            try:
                subj = Subject.objects.create(name=subj_name, is_active=True)
                subjects_map[norm] = subj
                subjects_created += 1
            except IntegrityError:
                # Race condition or duplicate - reload from DB
                try:
                    subj = Subject.objects.get(name__iexact=subj_name)
                    subjects_map[norm] = subj
                except Subject.DoesNotExist:
                    errors.append(f"Fach konnte nicht erstellt werden: {subj_name}")
        
        # 2. Create missing teachers
        for (last_name, first_name) in missing_data["missing_teachers"]:
            key = (last_name, first_name)
            if key in teachers_map:
                continue
            try:
                teacher = Teacher.objects.create(
                    last_name=last_name,
                    first_name=first_name,
                    is_active=True
                )
                teachers_map[key] = teacher
                teachers_created += 1
            except IntegrityError:
                # Race condition or duplicate - reload from DB
                try:
                    teacher = Teacher.objects.get(last_name=last_name, first_name=first_name)
                    teachers_map[key] = teacher
                except Teacher.DoesNotExist:
                    errors.append(f"Lehrer konnte nicht erstellt werden: {last_name}, {first_name}")
        
        # 3. Create missing assignments
        for (subj_norm, last_name, first_name) in missing_data["missing_assignments"]:
            # Find subject
            subj = subjects_map.get(subj_norm)
            if not subj:
                # Try to find by normalized name again (in case subject exists with different casing)
                for norm, s in subjects_map.items():
                    if norm == subj_norm:
                        subj = s
                        break
            if not subj:
                # Subject not found - skip
                continue
            
            # Find teacher
            teacher_key = (last_name, first_name)
            teacher = teachers_map.get(teacher_key)
            if not teacher:
                # Teacher not found - skip
                continue
            
            # Check if assignment already exists
            exists = TeachingAssignment.objects.filter(
                year=year, subject=subj, teacher=teacher
            ).exists()
            if exists:
                continue
            
            try:
                TeachingAssignment.objects.create(
                    year=year,
                    subject=subj,
                    teacher=teacher,
                    is_active=True
                )
                assignments_created += 1
            except IntegrityError:
                # Already exists (race condition)
                pass
        
        return {
            "subjects_created": subjects_created,
            "teachers_created": teachers_created,
            "assignments_created": assignments_created,
            "errors": errors,
        }
    
    def get_form_kwargs(self):
        """Add available years to form kwargs."""
        from apps.karteien.models import KarteiRecord
        
        kwargs = super().get_form_kwargs()
        available_years = list(
            KarteiRecord.objects.values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        kwargs["available_years"] = available_years
        return kwargs
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        from apps.karteien.models import KarteiRecord
        
        # Get available years from KarteiRecord
        available_years = list(
            KarteiRecord.objects.values_list("year", flat=True)
            .distinct()
            .order_by("-year")
        )
        context["available_years"] = available_years
        context["current_year"] = date.today().year
        
        # If year is in GET params, show preview
        year_param = self.request.GET.get("year")
        if year_param:
            try:
                year = int(year_param)
                legacy_data = self._get_legacy_data(year)
                missing_data = self._find_missing_entries(year, legacy_data)
                
                context["preview_year"] = year
                context["preview"] = {
                    "legacy_subjects_count": len(legacy_data["legacy_subjects"]),
                    "legacy_teachers_count": len(legacy_data["legacy_teachers"]),
                    "legacy_assignments_count": len(legacy_data["legacy_assignments"]),
                    "missing_subjects": missing_data["missing_subjects"],
                    "missing_subjects_count": len(missing_data["missing_subjects"]),
                    "missing_teachers": [
                        {"last_name": ln, "first_name": fn}
                        for ln, fn in missing_data["missing_teachers"]
                    ],
                    "missing_teachers_count": len(missing_data["missing_teachers"]),
                    "missing_assignments": [
                        {
                            "subject_norm": sn,
                            "subject_name": legacy_data["subject_display_by_norm"].get(sn, sn),
                            "last_name": ln,
                            "first_name": fn,
                        }
                        for sn, ln, fn in missing_data["missing_assignments"]
                    ],
                    "missing_assignments_count": len(missing_data["missing_assignments"]),
                    "unparsed_teachers": sorted(legacy_data["unparsed_teachers"]),
                    "unparsed_teachers_count": len(legacy_data["unparsed_teachers"]),
                }
                
                # Set year in form initial
                context["form"].initial["year"] = year
            except (ValueError, TypeError):
                pass
        
        return context
    
    def form_valid(self, form):
        """Handle POST - create missing entries."""
        year = form.cleaned_data["year"]
        
        # Get legacy data and find missing entries
        legacy_data = self._get_legacy_data(year)
        missing_data = self._find_missing_entries(year, legacy_data)
        
        # Check if there's anything to sync
        total_missing = (
            len(missing_data["missing_subjects"]) +
            len(missing_data["missing_teachers"]) +
            len(missing_data["missing_assignments"])
        )
        
        if total_missing == 0:
            messages.info(
                self.request,
                f"Keine fehlenden Einträge für Jahr {year} gefunden. Der Katalog ist vollständig."
            )
            return redirect(f"{self.success_url}?year={year}")
        
        # Sync entries
        stats = self._sync_entries(year, legacy_data, missing_data)
        
        # Build success message
        msg_parts = []
        if stats["subjects_created"] > 0:
            msg_parts.append(f"{stats['subjects_created']} Fächer hinzugefügt")
        if stats["teachers_created"] > 0:
            msg_parts.append(f"{stats['teachers_created']} Lehrer hinzugefügt")
        if stats["assignments_created"] > 0:
            msg_parts.append(f"{stats['assignments_created']} Zuweisungen hinzugefügt")
        
        if msg_parts:
            messages.success(
                self.request,
                f"Synchronisation für Jahr {year} abgeschlossen: " + ", ".join(msg_parts) + "."
            )
        else:
            messages.info(
                self.request,
                f"Keine neuen Einträge für Jahr {year} erstellt (alle bereits vorhanden)."
            )
        
        # Warning for unparsed teachers
        if legacy_data["unparsed_teachers"]:
            unparsed_list = ", ".join(sorted(legacy_data["unparsed_teachers"])[:5])
            more = len(legacy_data["unparsed_teachers"]) - 5
            if more > 0:
                unparsed_list += f" und {more} weitere"
            messages.warning(
                self.request,
                f"Einige Lehrernamen konnten nicht geparst werden: {unparsed_list}. "
                "Diese müssen manuell hinzugefügt werden."
            )
        
        # Show any errors
        for error in stats.get("errors", []):
            messages.error(self.request, error)
        
        return redirect(f"{self.success_url}?year={year}")


# =============================================================================
# Subject Category Views
# =============================================================================

class SubjectCategoryListView(CatalogAdminMixin, ListView):
    """List subject categories for a given year."""

    model = SubjectCategory
    template_name = "catalog/category_list.html"
    context_object_name = "categories"
    paginate_by = 50

    def get_year(self) -> int:
        return int(self.kwargs["year"])

    def get_queryset(self) -> QuerySet[SubjectCategory]:
        year = self.get_year()
        # Bootstrap default categories on first access
        ensure_default_categories(year)

        qs = (
            SubjectCategory.objects
            .filter(year=year)
            .annotate(links_count=Count("links"))
            .order_by("kind", "name")
        )

        # Optional kind filter
        kind = self.request.GET.get("kind", "").strip()
        if kind in (CategoryKind.GROUP, CategoryKind.INDIVIDUAL):
            qs = qs.filter(kind=kind)

        return qs

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["year"] = self.get_year()
        context["selected_kind"] = self.request.GET.get("kind", "")
        context["kind_choices"] = CategoryKind.choices
        return context


class SubjectCategoryCreateView(CatalogAdminMixin, CreateView):
    """Create a new subject category for a given year."""

    model = SubjectCategory
    form_class = SubjectCategoryForm
    template_name = "catalog/category_form.html"

    def get_year(self) -> int:
        return int(self.kwargs["year"])

    def get_success_url(self) -> str:
        return reverse("catalog:category_list", kwargs={"year": self.get_year()})

    def form_valid(self, form):
        form.instance.year = self.get_year()
        messages.success(self.request, "Kategorie wurde erfolgreich erstellt.")
        return super().form_valid(form)

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["year"] = self.get_year()
        context["is_create"] = True
        context["page_title"] = "Neue Kategorie"
        return context


class SubjectCategoryUpdateView(CatalogAdminMixin, UpdateView):
    """Edit an existing subject category."""

    model = SubjectCategory
    form_class = SubjectCategoryForm
    template_name = "catalog/category_form.html"

    def get_year(self) -> int:
        return int(self.kwargs["year"])

    def get_queryset(self) -> QuerySet[SubjectCategory]:
        return SubjectCategory.objects.filter(year=self.get_year())

    def get_success_url(self) -> str:
        return reverse("catalog:category_list", kwargs={"year": self.get_year()})

    def get_form(self, form_class=None):
        form = super().get_form(form_class)
        # Disallow changing kind when links exist
        if self.object.links.exists():
            form.fields["kind"].disabled = True
            form.fields["kind"].help_text = (
                "Art kann nicht geändert werden, solange Fächer zugeordnet sind."
            )
        return form

    def form_valid(self, form):
        messages.success(self.request, "Kategorie wurde erfolgreich aktualisiert.")
        return super().form_valid(form)

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["year"] = self.get_year()
        context["is_create"] = False
        context["page_title"] = f"Kategorie bearbeiten: {self.object.name}"
        return context


class SubjectCategoryDeleteView(CatalogAdminMixin, DeleteView):
    """Soft-delete a subject category (set is_active=False) and remove links."""

    model = SubjectCategory
    template_name = "catalog/category_confirm_delete.html"
    context_object_name = "category"

    def get_year(self) -> int:
        return int(self.kwargs["year"])

    def get_queryset(self) -> QuerySet[SubjectCategory]:
        return SubjectCategory.objects.filter(year=self.get_year())

    def get_success_url(self) -> str:
        return reverse("catalog:category_list", kwargs={"year": self.get_year()})

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["year"] = self.get_year()
        context["linked_subjects"] = (
            SubjectCategoryLink.objects
            .filter(category=self.object)
            .select_related("subject")
        )
        return context

    def form_valid(self, form=None):
        category = self.get_object()
        # Remove all links for this category
        deleted_count, _ = SubjectCategoryLink.objects.filter(category=category).delete()
        # Soft-delete: deactivate instead of removing
        category.is_active = False
        category.save(update_fields=["is_active"])
        if deleted_count:
            messages.success(
                self.request,
                f"Kategorie «{category.name}» deaktiviert. "
                f"{deleted_count} Fachzuordnung(en) entfernt.",
            )
        else:
            messages.success(
                self.request,
                f"Kategorie «{category.name}» wurde deaktiviert.",
            )
        return redirect(self.get_success_url())


# =============================================================================
# Subject ↔ Category Link Management Views
# =============================================================================

class SubjectCategoryLinksView(CatalogAdminMixin, TemplateView):
    """
    Manage subject links for a category.

    GET – display current links and a form to add new subjects.
    POST – create a new SubjectCategoryLink (+ DisciplineGroup / DurationEntry
    for GROUP categories).
    """

    template_name = "catalog/category_subjects.html"

    def get_year(self) -> int:
        return int(self.kwargs["year"])

    def get_category(self) -> SubjectCategory:
        return get_object_or_404(
            SubjectCategory,
            pk=self.kwargs["pk"],
            year=self.get_year(),
        )

    # ---- helpers ----

    def _available_subjects(self, year: int) -> QuerySet[Subject]:
        """Subjects not yet linked to any category in this year."""
        linked_ids = (
            SubjectCategoryLink.objects
            .filter(year=year)
            .values_list("subject_id", flat=True)
        )
        return (
            Subject.objects
            .filter(is_active=True)
            .exclude(pk__in=linked_ids)
            .order_by("name")
        )

    # ---- GET ----

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        category = self.get_category()
        year = self.get_year()

        context["category"] = category
        context["year"] = year
        context["links"] = (
            SubjectCategoryLink.objects
            .filter(category=category)
            .select_related("subject")
            .order_by("subject__name")
        )
        context["available_subjects"] = self._available_subjects(year)
        context["is_group"] = category.kind == CategoryKind.GROUP
        context["month_choices"] = list(range(1, 13))
        return context

    # ---- POST ----

    def post(self, request, *args, **kwargs):
        category = self.get_category()
        year = self.get_year()

        subject_id = request.POST.get("subject")
        if not subject_id:
            messages.error(request, "Bitte ein Fach auswählen.")
            return redirect(
                reverse(
                    "catalog:category_subjects",
                    kwargs={"year": year, "pk": category.pk},
                )
            )

        subject = get_object_or_404(Subject, pk=subject_id)

        # Duplicate check
        if SubjectCategoryLink.objects.filter(subject=subject, year=year).exists():
            messages.warning(
                request,
                f"Dieses Fach ist bereits einer Kategorie im Jahr {year} zugeordnet.",
            )
            return redirect(
                reverse(
                    "catalog:category_subjects",
                    kwargs={"year": year, "pk": category.pk},
                )
            )

        # Create link
        SubjectCategoryLink.objects.create(subject=subject, category=category)

        # GROUP-specific: ensure DisciplineGroup + initial DurationEntry
        if category.kind == CategoryKind.GROUP:
            activation_month = int(request.POST.get("activation_month", 1))
            initial_duration = int(request.POST.get("initial_duration", 45))

            group, created = DisciplineGroup.objects.get_or_create(
                subject=subject,
                year=year,
                defaults={"category": category},
            )
            if not created:
                group.category = category
                group.is_active = True
                group.save(update_fields=["category", "is_active"])

            # Create DurationEntry only if not already present for this month
            if not DurationEntry.objects.filter(
                group=group,
                effective_from_month=activation_month,
            ).exists():
                DurationEntry.objects.create(
                    group=group,
                    effective_from_month=activation_month,
                    duration_minutes=initial_duration,
                )

        messages.success(
            request,
            f"Fach «{subject.name}» wurde der Kategorie «{category.name}» zugeordnet.",
        )
        return redirect(
            reverse(
                "catalog:category_subjects",
                kwargs={"year": year, "pk": category.pk},
            )
        )


class SubjectCategoryUnlinkView(CatalogAdminMixin, View):
    """
    Remove a Subject ↔ Category link.

    POST-only. If the category is GROUP, deactivates the corresponding
    DisciplineGroup instead of deleting it.
    """

    def post(self, request, year: int, pk: int, link_pk: int):
        link = get_object_or_404(
            SubjectCategoryLink,
            pk=link_pk,
            category__pk=pk,
            category__year=year,
        )
        category = link.category
        subject = link.subject

        # Remove link
        link.delete()

        # Deactivate DisciplineGroup if GROUP category
        if category.kind == CategoryKind.GROUP:
            DisciplineGroup.objects.filter(
                subject=subject,
                year=year,
            ).update(is_active=False)

        messages.success(
            request,
            f"Fach «{subject.name}» wurde aus der Kategorie «{category.name}» entfernt.",
        )
        return redirect(
            reverse(
                "catalog:category_subjects",
                kwargs={"year": year, "pk": category.pk},
            )
        )


# =============================================================================
# DisciplineGroup Views (PROMPT_141.2)
# =============================================================================

class DisciplineGroupListView(CatalogEditorMixin, ListView):
    """List active DisciplineGroups for a given year."""

    template_name = "catalog/group_list.html"
    context_object_name = "groups"

    def get_queryset(self) -> QuerySet:
        year = self.kwargs["year"]
        qs = DisciplineGroup.objects.filter(year=year, is_active=True).select_related(
            "subject", "category",
        )
        category_id = self.request.GET.get("category")
        if category_id:
            qs = qs.filter(category_id=category_id)
        return qs.order_by("subject__name")

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        ctx = super().get_context_data(**kwargs)
        year = self.kwargs["year"]
        ctx["year"] = year
        ctx["categories"] = SubjectCategory.objects.filter(
            year=year, is_active=True,
        ).order_by("name")
        ctx["selected_category"] = self.request.GET.get("category", "")
        return ctx


class DisciplineGroupDetailView(CatalogEditorMixin, TemplateView):
    """Detail view for a single DisciplineGroup with duration/size entries."""

    template_name = "catalog/group_detail.html"

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        ctx = super().get_context_data(**kwargs)
        year = self.kwargs["year"]
        pk = self.kwargs["pk"]
        group = get_object_or_404(
            DisciplineGroup.objects.select_related("subject", "category"),
            pk=pk,
            year=year,
        )

        duration_entries = list(group.duration_entries.order_by("effective_from_month"))
        size_entries = list(group.size_entries.order_by("effective_from_month"))
        records_sem1 = _get_group_records(group, semester=1)
        records_sem2 = _get_group_records(group, semester=2)
        records_for_semester = {
            1: records_sem1,
            2: records_sem2,
        }

        from apps.karteien.billing import get_semester_for_month
        from apps.karteien.models import (
            get_contract_type_for_month,
            is_billable_in_month,
        )
        from .group_size_service import _is_month_active_for_slot

        # Prefetch category link once (same for all months)
        cat_link = SubjectCategoryLink.objects.filter(
            subject=group.subject,
            year=year,
            category__is_active=True,
        ).select_related("category").first()

        # Prefetch contract-type entries for all involved records once.
        ct_entries_by_pk: dict[int, list] = {}
        for rec in {r.pk: r for r in (records_sem1 + records_sem2)}.values():
            ct_entries_by_pk[rec.pk] = list(rec.contract_type_entries.all())

        monthly_summary = []
        for m in range(1, 13):
            size_info = get_group_size_for_month(group, m)
            semester = get_semester_for_month(m, year)

            # Contract type marker for this month from all billable records.
            # monthly/yearly/mixed/None
            contract_types: set[str] = set()
            monthly_sample_record = None
            yearly_sample_record = None
            for rec in records_for_semester.get(semester, []):
                if not _is_month_active_for_slot(rec, semester, m):
                    continue
                if not is_billable_in_month(rec, m):
                    continue
                is_monthly = get_contract_type_for_month(
                    rec, m, entries=ct_entries_by_pk.get(rec.pk),
                )
                if is_monthly:
                    contract_types.add("monthly")
                    if monthly_sample_record is None:
                        monthly_sample_record = rec
                else:
                    contract_types.add("yearly")
                    if yearly_sample_record is None:
                        yearly_sample_record = rec

            if len(contract_types) == 1:
                contract_type_marker = next(iter(contract_types))
            elif len(contract_types) > 1:
                contract_type_marker = "mixed"
            else:
                contract_type_marker = None

            suggested_monthly = None
            if monthly_sample_record is not None:
                suggested_monthly = calculate_suggested_price_group(
                    monthly_sample_record,
                    m,
                    link=cat_link,
                    group=group,
                    duration_entries=duration_entries,
                    contract_type_entries=ct_entries_by_pk.get(monthly_sample_record.pk),
                )

            suggested_yearly = None
            if yearly_sample_record is not None:
                suggested_yearly = calculate_suggested_price_group(
                    yearly_sample_record,
                    m,
                    link=cat_link,
                    group=group,
                    duration_entries=duration_entries,
                    contract_type_entries=ct_entries_by_pk.get(yearly_sample_record.pk),
                )

            if contract_type_marker == "monthly":
                suggested_primary = suggested_monthly
            elif contract_type_marker == "yearly":
                suggested_primary = suggested_yearly
            else:
                suggested_primary = None

            monthly_summary.append({
                "month": m,
                "duration": get_duration_for_month(group, m, entries=duration_entries),
                "manual_size": size_info["manual_size"],
                "auto_size": size_info["auto_size"],
                "effective_size": size_info["size"],
                "billable_count": size_info["billable_count"],
                "is_manual": size_info["is_manual"],
                "suggested_price": suggested_primary["price"] if suggested_primary else None,
                "suggested_base_price": suggested_primary["base_price"] if suggested_primary else None,
                "suggested_scaling": suggested_primary["scaling_applied"] if suggested_primary else False,
                "suggested_contract_type": contract_type_marker,
                "suggested_monthly_price": suggested_monthly["price"] if suggested_monthly else None,
                "suggested_monthly_base_price": suggested_monthly["base_price"] if suggested_monthly else None,
                "suggested_monthly_scaling": suggested_monthly["scaling_applied"] if suggested_monthly else False,
                "suggested_yearly_price": suggested_yearly["price"] if suggested_yearly else None,
                "suggested_yearly_base_price": suggested_yearly["base_price"] if suggested_yearly else None,
                "suggested_yearly_scaling": suggested_yearly["scaling_applied"] if suggested_yearly else False,
            })

        ctx["year"] = year
        ctx["group"] = group
        ctx["duration_entries"] = duration_entries
        ctx["size_entries"] = size_entries
        ctx["monthly_summary"] = monthly_summary
        ctx["sem1_stats"] = _get_group_semester_stats(group, semester=1)
        ctx["sem2_stats"] = _get_group_semester_stats(group, semester=2)
        ctx["duration_form"] = DurationEntryForm()
        ctx["size_form"] = GroupSizeEntryForm()
        ctx["warnings"] = get_group_warnings(group)
        return ctx


class DurationEntryCreateView(CatalogAdminMixin, View):
    """POST-only: create a DurationEntry for a group."""

    def post(self, request, year: int, pk: int):
        group = get_object_or_404(DisciplineGroup, pk=pk, year=year)
        form = DurationEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.group = group
            entry.changed_by = request.user
            try:
                entry.save()
                messages.success(
                    request,
                    f"Dauer-Eintrag für Monat {entry.effective_from_month} wurde erstellt.",
                )
            except IntegrityError:
                messages.error(
                    request,
                    f"Für Monat {form.cleaned_data['effective_from_month']} existiert bereits ein Dauer-Eintrag.",
                )
        else:
            for field, errors in form.errors.items():
                for error in errors:
                    messages.error(request, f"{field}: {error}")
        return redirect(
            reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
        )


class GroupSizeEntryCreateView(CatalogAdminMixin, View):
    """POST-only: create a GroupSizeEntry for a group."""

    def post(self, request, year: int, pk: int):
        group = get_object_or_404(DisciplineGroup, pk=pk, year=year)
        form = GroupSizeEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.group = group
            entry.changed_by = request.user
            try:
                entry.save()
                messages.success(
                    request,
                    f"Größen-Eintrag für Monat {entry.effective_from_month} wurde erstellt.",
                )
            except IntegrityError:
                messages.error(
                    request,
                    f"Für Monat {form.cleaned_data['effective_from_month']} existiert bereits ein Größen-Eintrag.",
                )
        else:
            for field, errors in form.errors.items():
                for error in errors:
                    messages.error(request, f"{field}: {error}")
        return redirect(
            reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
        )


class DisciplineGroupToggleScalingView(CatalogAdminMixin, View):
    """POST-only: toggle auto_scaling_enabled on a group."""

    def post(self, request, year: int, pk: int):
        group = get_object_or_404(DisciplineGroup, pk=pk, year=year)

        # Enabling is only allowed if the group has students in at least one month.
        if not group.auto_scaling_enabled:
            has_any_students = any(
                get_group_size_for_month(group, m)["size"] > 0
                for m in range(1, 13)
            )
            if not has_any_students:
                messages.error(
                    request,
                    "Auto-Skalierung kann nicht aktiviert werden: "
                    "Die Gruppe ist in allen Monaten leer (Groesse = 0).",
                )
                return redirect(
                    reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
                )

        group.auto_scaling_enabled = not group.auto_scaling_enabled
        group.save(update_fields=["auto_scaling_enabled"])
        state = "aktiviert" if group.auto_scaling_enabled else "deaktiviert"
        messages.success(
            request,
            f"Auto-Scaling für «{group}» wurde {state}.",
        )
        return redirect(
            reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
        )


class GroupSizeApiView(CatalogAdminMixin, View):
    """GET JSON: group size info for a single month."""

    def get(self, request, year: int, pk: int):
        group = get_object_or_404(DisciplineGroup, pk=pk, year=year)

        try:
            month = int(request.GET.get("month", 0))
        except (TypeError, ValueError):
            return JsonResponse({"error": "Invalid month parameter."}, status=400)

        if not 1 <= month <= 12:
            return JsonResponse({"error": "month must be 1-12."}, status=400)

        data = get_group_size_for_month(group, month)
        return JsonResponse(data)


# =============================================================================
# Group Legacy Preparation
# =============================================================================

class DisciplineGroupPrepareLegacyView(CatalogAdminMixin, View):
    """
    POST-only: prepare legacy records for category pricing in a semester.

    Steps per eligible record (status NORMAL):
    1) Set subject*_ref to the group's subject when missing.
    2) Convert LEGACY -> AUTO (full 12-month recalculation).
    3) Create/update PendingChange and set status=PENDING.
    """

    def post(self, request, year: int, pk: int):
        group = get_object_or_404(
            DisciplineGroup.objects.select_related("subject"),
            pk=pk,
            year=year,
        )

        try:
            semester = int(request.POST.get("semester", 0))
        except (TypeError, ValueError):
            semester = 0
        if semester not in (1, 2):
            messages.error(request, "Ungültiges Halbjahr.")
            return redirect(reverse("catalog:group_detail", kwargs={"year": year, "pk": pk}))

        comment = (request.POST.get("comment") or "").strip()

        from apps.approvals.services import create_or_update_pending_change
        from apps.karteien.billing import _normalize_subject_name, recalculate_legacy_to_auto
        from apps.karteien.models import KarteiRecord, MonthsMode, RecordStatus

        records = _get_group_records(group, semester=semester)
        if not records:
            messages.warning(request, "Keine Datensätze für diese Gruppe / dieses Halbjahr gefunden.")
            return redirect(reverse("catalog:group_detail", kwargs={"year": year, "pk": pk}))

        ref_field = "subject1_ref" if semester == 1 else "subject2_ref"
        ref_field_id = f"{ref_field}_id"
        legacy_field = "subject1" if semester == 1 else "subject2"
        norm_subject = _normalize_subject_name(group.subject.name)

        total = len(records)
        prepared = 0
        refs_set = 0
        converted_auto = 0
        skipped_status = 0
        skipped_unchanged = 0
        skipped_unmatched_legacy = 0
        failed = 0

        for record in records:
            if record.status != RecordStatus.NORMAL:
                skipped_status += 1
                continue

            changed = False
            old_base_amounts = dict(record.base_amounts or {})

            # Ensure subject*_ref is set for this semester.
            if getattr(record, ref_field_id) is None:
                legacy_value = (getattr(record, legacy_field) or "").strip()
                if _normalize_subject_name(legacy_value) == norm_subject:
                    setattr(record, ref_field, group.subject)
                    refs_set += 1
                    changed = True
                else:
                    skipped_unmatched_legacy += 1
                    continue

            # Convert LEGACY -> AUTO so category pricing can be applied.
            if record.months_mode == MonthsMode.LEGACY:
                try:
                    recalculate_legacy_to_auto(
                        record,
                        touched_months=set(range(1, 13)),
                        hours_amounts=(record.hours_amounts or {}),
                    )
                    converted_auto += 1
                    changed = True
                except Exception:
                    failed += 1
                    continue

            if not changed:
                skipped_unchanged += 1
                continue

            admin_comment = (
                f"[Gruppe: {group.subject.name}] LEGACY-Vorbereitung ({semester}. Halbjahr)"
            )
            if comment:
                admin_comment = f"{admin_comment}: {comment}"

            pending = create_or_update_pending_change(
                record, admin_comment=admin_comment,
            )

            pending_snapshot = dict(pending.snapshot or {})
            pending_snapshot["_old_base_amounts"] = old_base_amounts
            pending.snapshot = pending_snapshot
            pending.save(update_fields=["snapshot"])

            KarteiRecord.objects.filter(pk=record.pk).update(
                **{
                    ref_field_id: getattr(record, ref_field_id),
                    "months_mode": record.months_mode,
                    "base_amounts": record.base_amounts,
                    "hours_amounts": record.hours_amounts,
                    "legacy_base_amounts_enabled": record.legacy_base_amounts_enabled,
                    "status": RecordStatus.PENDING,
                }
            )
            prepared += 1

        if prepared:
            messages.success(
                request,
                f"Vorbereitung abgeschlossen ({semester}. Halbjahr): "
                f"{prepared}/{total} Datensatz/Datensätze als PENDING eingereicht "
                f"(subject_ref gesetzt: {refs_set}, LEGACY→AUTO: {converted_auto}).",
            )
        else:
            messages.warning(
                request,
                f"Keine Datensätze vorbereitet ({semester}. Halbjahr).",
            )

        if skipped_status:
            messages.info(
                request,
                f"Übersprungen wegen Status != NORMAL: {skipped_status}.",
            )
        if skipped_unchanged:
            messages.info(
                request,
                f"Bereits vorbereitet / keine Änderungen nötig: {skipped_unchanged}.",
            )
        if skipped_unmatched_legacy:
            messages.warning(
                request,
                f"Legacy-Fachname passte nicht eindeutig zur Gruppe: {skipped_unmatched_legacy}.",
            )
        if failed:
            messages.error(
                request,
                f"Fehler bei der Umstellung LEGACY→AUTO: {failed}.",
            )

        return redirect(
            reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
        )


# =============================================================================
# Bulk Apply Category Price (Preview + Apply)
# =============================================================================

class BulkApplyPreviewView(CatalogEditorMixin, TemplateView):
    """GET-only: preview bulk category-price application for group records."""

    template_name = "catalog/group_bulk_apply_preview.html"

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        ctx = super().get_context_data(**kwargs)
        year = self.kwargs["year"]
        pk = self.kwargs["pk"]
        group = get_object_or_404(
            DisciplineGroup.objects.select_related("subject"),
            pk=pk,
            year=year,
        )

        semester = int(self.request.GET.get("semester", 1))
        from_month = int(self.request.GET.get("from_month", 1))

        from apps.karteien.models import KarteiRecord, MonthsMode, RecordStatus

        # Find records linked to this group's subject in the matching semester
        # (with legacy fallback for records that still have only text subjects).
        records = _get_group_records(group, semester=semester)

        from copy import deepcopy
        from apps.karteien.category_price import apply_category_price_to_record

        preview_rows: list[dict[str, Any]] = []
        for record in records:
            row: dict[str, Any] = {
                "record": record,
                "eligible": False,
                "skip_reason": "",
                "diff": None,
                "error": None,
            }

            # Eligibility checks
            if record.status != RecordStatus.NORMAL:
                row["skip_reason"] = f"Status: {record.get_status_display()}"
            elif record.months_mode == MonthsMode.LEGACY:
                row["skip_reason"] = "Abrechnungsmodus: Legacy"
            else:
                row["eligible"] = True
                # Apply on a deep copy to avoid DB changes
                try:
                    record_copy = deepcopy(record)
                    diff = apply_category_price_to_record(
                        record_copy, semester=semester, from_month=from_month,
                    )
                    row["diff"] = diff
                except Exception as exc:
                    row["eligible"] = False
                    row["error"] = str(exc)

            preview_rows.append(row)

        ctx["year"] = year
        ctx["group"] = group
        ctx["semester"] = semester
        ctx["from_month"] = from_month
        ctx["comment"] = self.request.GET.get("comment", "")
        ctx["preview_rows"] = preview_rows
        ctx["eligible_count"] = sum(1 for r in preview_rows if r["eligible"])
        ctx["skip_count"] = sum(1 for r in preview_rows if not r["eligible"])
        return ctx


class BulkApplyCategoryPriceView(CatalogEditorMixin, View):
    """POST-only: bulk apply category prices and create PendingChanges."""

    def post(self, request, year: int, pk: int):
        group = get_object_or_404(
            DisciplineGroup.objects.select_related("subject"),
            pk=pk,
            year=year,
        )

        semester = int(request.POST.get("semester", 1))
        from_month = int(request.POST.get("from_month", 1))
        comment = request.POST.get("comment", "").strip()

        if not comment:
            messages.error(request, "Bitte geben Sie einen Kommentar ein.")
            return redirect(
                reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
            )

        # Operator: block past months (current/future months only).
        user = request.user
        if user.has_past_months_restrictions:
            from apps.karteien.validators import get_allowed_months

            allowed_fields, _reason = get_allowed_months(year)
            if f"month_{from_month}" not in allowed_fields:
                messages.error(
                    request,
                    f"Sie dürfen vergangene Monate nicht ändern (Monat {from_month})."
                )
                return redirect(
                    reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
                )

        from apps.karteien.models import KarteiRecord, MonthsMode, RecordStatus
        from apps.karteien.category_price import apply_category_price_to_record
        from apps.approvals.services import create_or_update_pending_change

        # Find records linked to this group's subject in the matching semester
        # (with legacy fallback for records that still have only text subjects).
        records = _get_group_records(group, semester=semester)

        updated_count = 0
        skipped_count = 0

        for record in records:
            # Skip ineligible records
            if record.status != RecordStatus.NORMAL:
                skipped_count += 1
                continue
            if record.months_mode == MonthsMode.LEGACY:
                skipped_count += 1
                continue

            old_base_amounts = dict(record.base_amounts or {})

            # Apply category price in-memory
            try:
                apply_category_price_to_record(
                    record, semester=semester, from_month=from_month,
                )
            except Exception:
                skipped_count += 1
                continue

            # Create pending change
            admin_comment = f"[Gruppe: {group.subject.name}] {comment}"
            pending = create_or_update_pending_change(
                record, admin_comment=admin_comment
            )
            pending_snapshot = dict(pending.snapshot or {})
            pending_snapshot["_old_base_amounts"] = old_base_amounts
            pending.snapshot = pending_snapshot
            pending.save(update_fields=["snapshot"])

            # Save safe fields via queryset update (no full record.save())
            KarteiRecord.objects.filter(pk=record.pk).update(
                status=RecordStatus.PENDING,
                base_amounts=record.base_amounts,
                months_mode=record.months_mode,
            )
            updated_count += 1

        messages.success(
            request,
            f"{updated_count} Datensatz/Datensätze aktualisiert, "
            f"{skipped_count} übersprungen.",
        )
        return redirect(
            reverse("catalog:group_detail", kwargs={"year": year, "pk": pk})
        )


# =============================================================================
# Copy Categories Between Years
# =============================================================================


class CopyCategoriesView(CatalogAdminMixin, FormView):
    """
    Copy subject categories (with links and discipline groups) from one year
    to another.

    GET  → form with source_year / target_year / overwrite checkbox.
    POST → call ``copy_categories_between_years`` and show result.
    """

    template_name = "catalog/copy_categories_year.html"
    form_class = CopyCategoriesYearForm

    def form_valid(self, form):
        from_year = form.cleaned_data["from_year"]
        to_year = form.cleaned_data["to_year"]
        overwrite = form.cleaned_data["overwrite"]

        result = copy_categories_between_years(
            source_year=from_year,
            target_year=to_year,
            overwrite=overwrite,
        )

        # Warnings only (no categories copied)
        if result.warnings and result.categories_created == 0:
            for w in result.warnings:
                messages.warning(self.request, w)
            return self.form_invalid(form)

        # Build success message
        parts = [f"Kopieren von {from_year} nach {to_year} abgeschlossen."]
        if result.categories_created:
            parts.append(f"{result.categories_created} Kategorien erstellt.")
        if result.categories_skipped:
            parts.append(
                f"{result.categories_skipped} Kategorien übersprungen."
            )
        if result.links_created:
            parts.append(f"{result.links_created} Verknüpfungen erstellt.")
        if result.links_skipped:
            parts.append(
                f"{result.links_skipped} Verknüpfungen übersprungen."
            )
        if result.groups_created:
            parts.append(f"{result.groups_created} Gruppen erstellt.")

        messages.success(self.request, " ".join(parts))

        # Show individual warnings (e.g. duplicate subject links)
        for w in result.warnings:
            messages.warning(self.request, w)

        return redirect(
            reverse("catalog:category_list", kwargs={"year": to_year})
        )

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)

        year_stats = (
            SubjectCategory.objects.values("year")
            .annotate(
                total=Count("id"),
                active=Count("id", filter=Q(is_active=True)),
            )
            .order_by("-year")
        )
        context["year_stats"] = list(year_stats)
        context["current_year"] = date.today().year
        return context


class CopyCategoriesPreviewView(CatalogAdminMixin, TemplateView):
    """
    Preview which categories / links would be created when copying between
    years.  Read-only — does **not** write to the database.
    """

    template_name = "catalog/copy_categories_preview.html"

    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)

        source_year = self.request.GET.get("from_year")
        target_year = self.request.GET.get("to_year")

        if not source_year or not target_year:
            context["error"] = "Bitte Quell- und Zieljahr angeben."
            return context

        try:
            source_year = int(source_year)
            target_year = int(target_year)
        except (ValueError, TypeError):
            context["error"] = "Ungültige Jahrwerte."
            return context

        if source_year == target_year:
            context["error"] = (
                "Quell- und Zieljahr dürfen nicht identisch sein."
            )
            return context

        context["source_year"] = source_year
        context["target_year"] = target_year

        # Source categories (active only)
        source_cats = list(
            SubjectCategory.objects.filter(
                year=source_year, is_active=True
            ).order_by("name")
        )

        # Source links grouped by category
        source_links = (
            SubjectCategoryLink.objects.filter(year=source_year)
            .select_related("subject", "category")
            .order_by("category__name", "subject__name")
        )
        links_by_cat: dict[int, list] = {}
        for link in source_links:
            links_by_cat.setdefault(link.category_id, []).append(link)

        # Subjects already linked in target year
        already_linked = set(
            SubjectCategoryLink.objects.filter(year=target_year).values_list(
                "subject_id", flat=True
            )
        )
        context["already_linked_subject_ids"] = already_linked

        # Build template-friendly list: category + its links + conflict info
        categories_with_links = []
        for cat in source_cats:
            cat_links = links_by_cat.get(cat.pk, [])
            annotated_links = []
            for link in cat_links:
                annotated_links.append({
                    "subject_name": link.subject.name,
                    "subject_id": link.subject_id,
                    "already_linked": link.subject_id in already_linked,
                })
            categories_with_links.append({
                "category": cat,
                "links": annotated_links,
            })
        context["source_categories"] = categories_with_links

        # Existing target categories
        target_cats = list(
            SubjectCategory.objects.filter(year=target_year).order_by("name")
        )
        context["target_categories"] = target_cats

        # Warnings
        warnings: list[str] = []
        if target_cats:
            active_count = sum(1 for c in target_cats if c.is_active)
            warnings.append(
                f"Im Zieljahr {target_year} existieren bereits "
                f"{len(target_cats)} Kategorien ({active_count} aktiv). "
                "Ohne \"Überschreiben\" wird keine Kopie durchgeführt."
            )
        if not source_cats:
            warnings.append(
                f"Keine aktiven Kategorien im Quelljahr {source_year} "
                "gefunden."
            )
        context["warnings"] = warnings

        return context

