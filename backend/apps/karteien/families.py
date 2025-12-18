"""
Family Dashboard views for the karteien app.

This module contains:
- FamilyDashboardView: Shows all records for a family in a year, with discount management
- FamilyDiscountCreateView: Add a FamilyDiscount
- FamilyDiscountEditView: Edit a FamilyDiscount
- FamilyDiscountDeleteView: Delete a FamilyDiscount
- ApplyDiscountsView: Apply discounts to all NORMAL+AUTO records, creating PendingChanges

Access is restricted to Admin role only.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.db import transaction
from django.http import HttpRequest, HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views import View

from apps.approvals.services import build_snapshot, create_or_update_pending_change_from_snapshot
from apps.catalog.models import Discount, FamilyDiscount

from .billing import calculate_month_values, build_base_amounts, ZERO
from .models import KarteiRecord, RecordStatus, MonthsMode, MONTH_FIELD_NAMES


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
# Family Dashboard View
# =============================================================================

class FamilyDashboardView(AdminOnlyMixin, View):
    """
    Family Dashboard showing all records for a family in a year.
    
    Displays:
    - List of all KarteiRecord entries for the family/year
    - Family discounts (FamilyDiscount) with add/edit/delete controls
    - "Apply Discounts" button to create PendingChange for affected records
    
    URL: /karteien/family/?year=YYYY&family_id=XXX
    """
    
    template_name = "karteien/family_dashboard.html"
    
    def get(self, request: HttpRequest) -> HttpResponse:
        """Display family dashboard."""
        year = request.GET.get("year")
        family_id = request.GET.get("family_id")
        
        # Validate required parameters
        if not year or not family_id:
            messages.error(request, "Jahr und FamilyID müssen angegeben werden.")
            return redirect("karteien:record_list")
        
        try:
            year = int(year)
        except ValueError:
            messages.error(request, "Ungültiges Jahr.")
            return redirect("karteien:record_list")
        
        # Get all records for this family in this year
        records = KarteiRecord.objects.filter(
            year=year,
            family_id=family_id,
        ).order_by("child_name", "subject1")
        
        if not records.exists():
            messages.warning(
                request,
                f"Keine Datensätze für Familie '{family_id}' im Jahr {year} gefunden."
            )
            return redirect("karteien:record_list")
        
        # Get family discounts
        family_discounts = FamilyDiscount.objects.filter(
            year=year,
            family_id=family_id,
        ).select_related("discount").order_by("start_month")
        
        # Get all available discounts for dropdown
        available_discounts = Discount.objects.filter(is_active=True).order_by("kind", "-value")
        
        # Calculate summary statistics
        total_records = records.count()
        normal_auto_records = records.filter(
            status=RecordStatus.NORMAL,
            months_mode=MonthsMode.AUTO,
        ).count()
        pending_records = records.filter(status=RecordStatus.PENDING).count()
        declined_records = records.filter(status=RecordStatus.DECLINED).count()
        legacy_records = records.filter(months_mode=MonthsMode.LEGACY).count()
        
        context = {
            "year": year,
            "family_id": family_id,
            "records": records,
            "family_discounts": family_discounts,
            "available_discounts": available_discounts,
            "total_records": total_records,
            "normal_auto_records": normal_auto_records,
            "pending_records": pending_records,
            "declined_records": declined_records,
            "legacy_records": legacy_records,
        }
        
        return render(request, self.template_name, context)


# =============================================================================
# FamilyDiscount CRUD Views
# =============================================================================

class FamilyDiscountCreateView(AdminOnlyMixin, View):
    """
    Create a new FamilyDiscount.
    
    POST /karteien/family/discounts/create/
    Form fields: year, family_id, discount, start_month, end_month
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Create a new FamilyDiscount."""
        year = request.POST.get("year")
        family_id = request.POST.get("family_id")
        discount_id = request.POST.get("discount")
        start_month = request.POST.get("start_month", "1")
        end_month = request.POST.get("end_month", "12")
        
        # Validate
        if not all([year, family_id, discount_id]):
            messages.error(request, "Alle Pflichtfelder müssen ausgefüllt sein.")
            return self._redirect_back(request, year, family_id)
        
        try:
            year = int(year)
            discount_id = int(discount_id)
            start_month = int(start_month)
            end_month = int(end_month)
        except ValueError:
            messages.error(request, "Ungültige Eingabe.")
            return self._redirect_back(request, year, family_id)
        
        # Validate month range
        if start_month < 1 or start_month > 12 or end_month < 1 or end_month > 12:
            messages.error(request, "Monate müssen zwischen 1 und 12 liegen.")
            return self._redirect_back(request, year, family_id)
        
        if start_month > end_month:
            messages.error(request, "Startmonat darf nicht größer als Endmonat sein.")
            return self._redirect_back(request, year, family_id)
        
        # Get discount
        try:
            discount = Discount.objects.get(pk=discount_id)
        except Discount.DoesNotExist:
            messages.error(request, "Rabatt nicht gefunden.")
            return self._redirect_back(request, year, family_id)
        
        # Create FamilyDiscount
        FamilyDiscount.objects.create(
            year=year,
            family_id=family_id,
            discount=discount,
            start_month=start_month,
            end_month=end_month,
        )
        
        messages.success(
            request,
            f"Familienrabatt '{discount}' wurde hinzugefügt."
        )
        
        return self._redirect_back(request, year, family_id)
    
    def _redirect_back(self, request: HttpRequest, year, family_id) -> HttpResponse:
        """Redirect back to family dashboard."""
        url = reverse("karteien:family_dashboard")
        return redirect(f"{url}?year={year}&family_id={family_id}")


class FamilyDiscountEditView(AdminOnlyMixin, View):
    """
    Edit an existing FamilyDiscount.
    
    POST /karteien/family/discounts/<pk>/edit/
    """
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Update a FamilyDiscount."""
        family_discount = get_object_or_404(FamilyDiscount, pk=pk)
        
        discount_id = request.POST.get("discount")
        start_month = request.POST.get("start_month", "1")
        end_month = request.POST.get("end_month", "12")
        
        try:
            discount_id = int(discount_id)
            start_month = int(start_month)
            end_month = int(end_month)
        except ValueError:
            messages.error(request, "Ungültige Eingabe.")
            return self._redirect_back(request, family_discount)
        
        # Validate month range
        if start_month < 1 or start_month > 12 or end_month < 1 or end_month > 12:
            messages.error(request, "Monate müssen zwischen 1 und 12 liegen.")
            return self._redirect_back(request, family_discount)
        
        if start_month > end_month:
            messages.error(request, "Startmonat darf nicht größer als Endmonat sein.")
            return self._redirect_back(request, family_discount)
        
        # Get discount
        try:
            discount = Discount.objects.get(pk=discount_id)
        except Discount.DoesNotExist:
            messages.error(request, "Rabatt nicht gefunden.")
            return self._redirect_back(request, family_discount)
        
        # Update
        family_discount.discount = discount
        family_discount.start_month = start_month
        family_discount.end_month = end_month
        family_discount.save()
        
        messages.success(request, "Familienrabatt wurde aktualisiert.")
        
        return self._redirect_back(request, family_discount)
    
    def _redirect_back(self, request: HttpRequest, family_discount: FamilyDiscount) -> HttpResponse:
        """Redirect back to family dashboard."""
        url = reverse("karteien:family_dashboard")
        return redirect(f"{url}?year={family_discount.year}&family_id={family_discount.family_id}")


class FamilyDiscountDeleteView(AdminOnlyMixin, View):
    """
    Delete a FamilyDiscount.
    
    POST /karteien/family/discounts/<pk>/delete/
    """
    
    def post(self, request: HttpRequest, pk: int) -> HttpResponse:
        """Delete a FamilyDiscount."""
        family_discount = get_object_or_404(FamilyDiscount, pk=pk)
        
        year = family_discount.year
        family_id = family_discount.family_id
        discount_str = str(family_discount.discount)
        
        family_discount.delete()
        
        messages.success(
            request,
            f"Familienrabatt '{discount_str}' wurde gelöscht."
        )
        
        url = reverse("karteien:family_dashboard")
        return redirect(f"{url}?year={year}&family_id={family_id}")


# =============================================================================
# Apply Discounts View
# =============================================================================

class ApplyDiscountsView(AdminOnlyMixin, View):
    """
    Apply family discounts to all affected records.
    
    For each NORMAL + AUTO record in the family/year:
    1. Calculate new month values based on current discounts
    2. If values changed, create/update PendingChange
    3. Set record status to PENDING
    
    Records with LEGACY mode, PENDING status, or DECLINED status are skipped.
    
    POST /karteien/family/apply-discounts/
    Form fields: year, family_id
    """
    
    def post(self, request: HttpRequest) -> HttpResponse:
        """Apply discounts to family records."""
        year = request.POST.get("year")
        family_id = request.POST.get("family_id")
        
        # Validate
        if not year or not family_id:
            messages.error(request, "Jahr und FamilyID müssen angegeben werden.")
            return redirect("karteien:record_list")
        
        try:
            year = int(year)
        except ValueError:
            messages.error(request, "Ungültiges Jahr.")
            return redirect("karteien:record_list")
        
        # Get all NORMAL + AUTO records for this family/year
        eligible_records = KarteiRecord.objects.filter(
            year=year,
            family_id=family_id,
            status=RecordStatus.NORMAL,
            months_mode=MonthsMode.AUTO,
        )
        
        # Get family discounts once
        family_discounts = list(
            FamilyDiscount.objects.filter(
                year=year,
                family_id=family_id,
            ).select_related("discount")
        )
        
        # Statistics
        processed_count = 0
        pending_created = 0
        skipped_no_change = 0
        skipped_no_refs = 0
        errors = []
        
        with transaction.atomic():
            for record in eligible_records:
                try:
                    result = self._process_record(record, family_discounts)
                    processed_count += 1
                    
                    if result == "pending":
                        pending_created += 1
                    elif result == "no_change":
                        skipped_no_change += 1
                    elif result == "no_refs":
                        skipped_no_refs += 1
                        
                except Exception as e:
                    errors.append(f"Record {record.id}: {str(e)}")
        
        # Report results
        if pending_created > 0:
            messages.success(
                request,
                f"{pending_created} Datensätze wurden zu PENDING geändert und warten auf Superadmin-Prüfung."
            )
        
        if skipped_no_change > 0:
            messages.info(
                request,
                f"{skipped_no_change} Datensätze unverändert (keine Betragsänderung)."
            )
        
        if skipped_no_refs > 0:
            messages.warning(
                request,
                f"{skipped_no_refs} Datensätze übersprungen (keine Preis-Referenzen)."
            )
        
        # Count other skipped records
        total_family_records = KarteiRecord.objects.filter(
            year=year,
            family_id=family_id,
        ).count()
        
        other_skipped = total_family_records - processed_count
        if other_skipped > 0:
            # Get breakdown
            pending_count = KarteiRecord.objects.filter(
                year=year, family_id=family_id, status=RecordStatus.PENDING
            ).exclude(pk__in=eligible_records).count()
            declined_count = KarteiRecord.objects.filter(
                year=year, family_id=family_id, status=RecordStatus.DECLINED
            ).count()
            legacy_count = KarteiRecord.objects.filter(
                year=year, family_id=family_id, months_mode=MonthsMode.LEGACY
            ).count()
            
            skip_parts = []
            if pending_count > 0:
                skip_parts.append(f"{pending_count} PENDING")
            if declined_count > 0:
                skip_parts.append(f"{declined_count} DECLINED")
            if legacy_count > 0:
                skip_parts.append(f"{legacy_count} LEGACY")
            
            if skip_parts:
                messages.info(
                    request,
                    f"Übersprungen: {', '.join(skip_parts)}"
                )
        
        if errors:
            for error in errors[:5]:  # Show max 5 errors
                messages.error(request, error)
            if len(errors) > 5:
                messages.error(request, f"... und {len(errors) - 5} weitere Fehler.")
        
        # Redirect back to family dashboard
        url = reverse("karteien:family_dashboard")
        return redirect(f"{url}?year={year}&family_id={family_id}")
    
    def _process_record(
        self,
        record: KarteiRecord,
        family_discounts: list[FamilyDiscount],
    ) -> str:
        """
        Process a single record for discount application.
        
        Returns:
            "pending" - PendingChange created, record marked PENDING
            "no_change" - No value changes, skipped
            "no_refs" - No price references, skipped
        """
        # Check if record has price references (either catalog refs or legacy prices)
        has_price_1 = record.price1_ref_id or record.price1
        has_price_2 = record.price2_ref_id or record.price2
        
        if not has_price_1 and not has_price_2:
            return "no_refs"
        
        # Get record-level discounts
        record_discounts = list(
            record.record_discounts.select_related("discount")
        )
        
        # Calculate new month values with current discounts
        base_amounts = build_base_amounts(record)
        new_month_values, flags = calculate_month_values(
            record,
            family_discounts=family_discounts,
            record_discounts=record_discounts,
            base_amounts=base_amounts,
        )
        
        # Compare with current values
        has_changes = False
        for month_num in range(1, 13):
            field_key = f"month_{month_num}"
            current_value = getattr(record, field_key) or ZERO
            new_value = new_month_values.get(field_key, ZERO)
            
            # Normalize for comparison
            if current_value is None:
                current_value = ZERO
            if isinstance(current_value, (int, float, str)):
                current_value = Decimal(str(current_value))
            if isinstance(new_value, (int, float, str)):
                new_value = Decimal(str(new_value))
            
            # Compare normalized values
            if current_value.quantize(Decimal("0.01")) != new_value.quantize(Decimal("0.01")):
                has_changes = True
                break
        
        if not has_changes:
            return "no_change"
        
        # Build snapshot with new values
        # Create a temporary modified record for snapshot
        snapshot_record = KarteiRecord(
            pkid=record.pkid,
            id=record.id,
            year=record.year,
            family_id=record.family_id,
            parent_name=record.parent_name,
            child_name=record.child_name,
            birthdate=record.birthdate,
            address=record.address,
            phone=record.phone,
            mobile=record.mobile,
            email=record.email,
            subject1=record.subject1,
            price1=record.price1,
            subject2=record.subject2,
            price2=record.price2,
            extra1=record.extra1,
            extra2=record.extra2,
            extra3=record.extra3,
        )
        
        # Set new month values on snapshot record
        for month_num in range(1, 13):
            field_key = f"month_{month_num}"
            setattr(snapshot_record, field_key, new_month_values.get(field_key, ZERO))
        
        # Build snapshot
        snapshot = build_snapshot(snapshot_record)
        
        # Create/update PendingChange
        create_or_update_pending_change_from_snapshot(record, snapshot)
        
        # Update record status to PENDING
        record.status = RecordStatus.PENDING
        record.save(update_fields=["status", "updated_at"])
        
        return "pending"
