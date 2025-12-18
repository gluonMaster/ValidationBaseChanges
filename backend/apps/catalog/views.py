"""
Views for catalog app.

This module contains views for managing:
- Teachers (Lehrer)
- Subjects (Fächer)
- Teaching Assignments (Zuweisungen)
- Price Options (Preise)
- Copying assignments and prices between years

Access is restricted to Admin role only.
"""

from __future__ import annotations

from datetime import date
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db import IntegrityError
from django.db.models import QuerySet
from django.http import HttpResponse
from django.shortcuts import redirect
from django.urls import reverse_lazy
from django.views.generic import (
    CreateView,
    DeleteView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
)

from .forms import (
    CopyYearForm,
    CopyPricesYearForm,
    PriceOptionForm,
    SubjectForm,
    TeacherForm,
    TeachingAssignmentForm,
)
from .models import Subject, Teacher, TeachingAssignment, PriceOption, Discount, FamilyDiscount, RecordDiscount


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
    
    def form_valid(self, form):
        messages.success(self.request, "Lehrer wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Lehrer"
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
    
    def form_valid(self, form):
        messages.success(self.request, "Fach wurde erfolgreich erstellt.")
        return super().form_valid(form)
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neues Fach"
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
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neue Zuweisung"
        return context
    
    def get_initial(self):
        initial = super().get_initial()
        # Pre-fill year from query parameter
        year = self.request.GET.get("year")
        if year:
            try:
                initial["year"] = int(year)
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
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = True
        context["page_title"] = "Neuer Preis"
        return context
    
    def get_initial(self):
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
        qs = FamilyDiscount.objects.select_related('discount').all()
        
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
        record_pk = self.object.record.pkid
        return f"{reverse_lazy('catalog:record_discount_list')}?record_pk={record_pk}"
    
    def get_context_data(self, **kwargs) -> dict[str, Any]:
        context = super().get_context_data(**kwargs)
        context["is_create"] = False
        context["page_title"] = f"Eintrag-Rabatt bearbeiten: #{self.object.record.pkid}"
        context["target_record"] = self.object.record
        return context


class RecordDiscountDeleteView(CatalogAdminMixin, DeleteView):
    """Delete a record discount."""
    
    model = RecordDiscount
    template_name = "catalog/record_discount_confirm_delete.html"
    success_url = reverse_lazy("catalog:record_discount_list")
    
    def form_valid(self, form):
        record_pk = self.object.record.pkid
        messages.success(self.request, "Eintrag-Rabatt wurde erfolgreich gelöscht.")
        response = super().form_valid(form)
        return redirect(f"{self.success_url}?record_pk={record_pk}")
