"""
Kosten (costs) report view for families.

This module contains:
- FamilyKostenReportView: HTML report showing monthly costs for a family
  across the last 3 years available in the database.
- FamilyKostenFragmentView: Partial HTML fragment for use in Offcanvas panel.

Access is restricted to Superadmin role only.

Report structure:
- One section per year (latest 3 years with data for the family)
- One row per child/record
- Columns split by semester according to SemesterConfig for each year
- Summary row with totals per month
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Any

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views import View

from .billing import get_semester_month_ranges
from .models import KarteiRecord, RecordStatus, MONTH_FIELD_NAMES


# =============================================================================
# Constants
# =============================================================================

# Zero value for calculations
ZERO = Decimal("0.00")

# Month names in German (short form)
MONTH_NAMES_DE = (
    "Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
    "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"
)


# =============================================================================
# Permission Mixin
# =============================================================================

class SuperadminOnlyMixin(LoginRequiredMixin, UserPassesTestMixin):
    """
    Mixin that restricts access to Superadmin users only.
    """
    
    def test_func(self) -> bool:
        """Check if user is Superadmin."""
        user = self.request.user
        return user.is_authenticated and user.is_superadmin
    
    def handle_no_permission(self) -> HttpResponse:
        """Redirect to login or show error for unauthorized users."""
        if not self.request.user.is_authenticated:
            return super().handle_no_permission()
        
        messages.error(
            self.request,
            "Dieser Bericht ist nur für Superadmin verfügbar."
        )
        return redirect("karteien:record_list")


# =============================================================================
# Family Kosten Report View
# =============================================================================

class FamilyKostenReportView(SuperadminOnlyMixin, View):
    """
    Family Kosten (costs) report showing monthly charges for a family.
    
    Displays:
    - Last 3 years that have data for the family
    - One row per child/record (child_name)
    - Columns split by semester per year (SemesterConfig boundary)
    - Summary row with monthly totals
    - For current year, future months are empty and excluded from totals
    
    URL: /karteien/family-kosten/?family_id=XXX
    """
    
    template_name = "karteien/family_kosten_report.html"
    
    def get(self, request: HttpRequest) -> HttpResponse:
        """Display family Kosten report."""
        family_id = request.GET.get("family_id", "").strip()
        
        # Validate required parameter
        if not family_id:
            messages.error(request, "FamilyID muss angegeben werden.")
            return redirect("karteien:record_list")
        
        # Get current date for determining future months
        today = date.today()
        current_year = today.year
        current_month = today.month
        
        # Find the last 3 distinct years in the database
        all_years = (
            KarteiRecord.objects
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")[:3]
        )
        all_years = list(all_years)
        
        if not all_years:
            messages.warning(request, "Keine Daten in der Datenbank vorhanden.")
            return redirect("karteien:record_list")
        
        # Build data for each year
        years_data = []
        
        for year in all_years:
            sem1_months, sem2_months = get_semester_month_ranges(year)
            sem1_len = len(sem1_months)

            # Get records for this family in this year
            # Filter: status in (NORMAL, PENDING) - exclude DECLINED
            records = KarteiRecord.objects.filter(
                family_id=family_id,
                year=year,
                status__in=[RecordStatus.NORMAL, RecordStatus.PENDING],
            ).order_by("child_name", "subject1")
            
            if not records.exists():
                # Skip years where family has no records
                continue
            
            # Check if any records are PENDING
            has_pending_records = records.filter(status=RecordStatus.PENDING).exists()
            
            # Build rows for each record
            rows = []
            # Monthly sums (1-12)
            month_sums = [ZERO] * 12
            
            for record in records:
                row = self._build_record_row(
                    record,
                    year,
                    current_year,
                    current_month,
                    sem1_len,
                )
                rows.append(row)
                
                # Accumulate monthly sums
                for i, month_val in enumerate(row["months"]):
                    if month_val is not None:
                        month_sums[i] += month_val
            
            # For current year, nullify future month sums
            if year == current_year:
                for i in range(current_month, 12):
                    month_sums[i] = None
            
            # Calculate year total (sum of all months with values)
            year_total = sum(v for v in month_sums if v is not None)
            
            years_data.append({
                "year": year,
                "rows": rows,
                "month_sums": month_sums,
                "month_sums_sem1": month_sums[:sem1_len],
                "month_sums_sem2": month_sums[sem1_len:],
                "month_names_sem1": MONTH_NAMES_DE[:sem1_len],
                "month_names_sem2": MONTH_NAMES_DE[sem1_len:],
                "year_total": year_total,
                "has_pending_records": has_pending_records,
            })
        
        if not years_data:
            messages.warning(
                request,
                f"Keine Datensätze für FamilyID '{family_id}' in den letzten Jahren gefunden."
            )
            return redirect("karteien:record_list")
        
        # Get parent name from first record (for display)
        first_record = years_data[0]["rows"][0]["record"] if years_data[0]["rows"] else None
        parent_name = first_record.parent_name if first_record else ""
        
        context = {
            "family_id": family_id,
            "parent_name": parent_name,
            "years_data": years_data,
            "month_names": MONTH_NAMES_DE,
        }
        
        return render(request, self.template_name, context)
    
    def _build_record_row(
        self,
        record: KarteiRecord,
        year: int,
        current_year: int,
        current_month: int,
        sem1_len: int,
    ) -> dict[str, Any]:
        """
        Build a row dict for a single record.
        
        Returns:
            dict with keys:
            - record: the KarteiRecord instance
            - child_name: str
            - subject1: str
            - subject2: str
            - months: list of 12 Decimal values (None for future months)
            - months_sem1: list of semester-1 months
            - months_sem2: list of semester-2 months
            - is_pending: bool
            - row_total: Decimal (sum of non-None months)
        """
        # Get month values (month_1 to month_12)
        months = []
        for i in range(1, 13):
            field_name = f"month_{i}"
            value = getattr(record, field_name, None)
            
            # For current year, future months should be None (displayed as "-")
            if year == current_year and i > current_month:
                value = None
            elif value is None:
                # Past or present months: treat None as 0.00
                value = ZERO
            
            months.append(value)
        
        # Calculate row total (sum of non-None months)
        row_total = sum(v for v in months if v is not None)
        
        return {
            "record": record,
            "child_name": record.child_name or "-",
            "subject1": record.subject1 or "-",
            "subject2": record.subject2 or "-",
            "months": months,
            "months_sem1": months[:sem1_len],
            "months_sem2": months[sem1_len:],
            "is_pending": record.status == RecordStatus.PENDING,
            "row_total": row_total,
        }


# =============================================================================
# Family Kosten Fragment View (for Offcanvas)
# =============================================================================

class FamilyKostenFragmentView(SuperadminOnlyMixin, View):
    """
    Family Kosten fragment view for embedding in Offcanvas panel.
    
    Returns partial HTML without the base layout, suitable for AJAX loading.
    Uses the same data logic as FamilyKostenReportView.
    
    URL: /api/karteien/family-kosten-fragment/?family_id=XXX
    """
    
    template_name = "karteien/_family_kosten_report_fragment.html"
    
    def handle_no_permission(self) -> HttpResponse:
        """Return JSON error for unauthorized AJAX requests."""
        if not self.request.user.is_authenticated:
            return JsonResponse({"error": "Nicht angemeldet"}, status=401)
        return JsonResponse({"error": "Nur für Superadmin"}, status=403)
    
    def get(self, request: HttpRequest) -> HttpResponse:
        """Return family Kosten report as HTML fragment."""
        family_id = request.GET.get("family_id", "").strip()
        
        # Validate required parameter
        if not family_id:
            return HttpResponse(
                '<div class="alert alert-danger">FamilyID fehlt.</div>',
                content_type="text/html"
            )
        
        # Get current date for determining future months
        today = date.today()
        current_year = today.year
        current_month = today.month
        
        # Find the last 3 distinct years in the database
        all_years = (
            KarteiRecord.objects
            .values_list("year", flat=True)
            .distinct()
            .order_by("-year")[:3]
        )
        all_years = list(all_years)
        
        if not all_years:
            return HttpResponse(
                '<div class="alert alert-warning">Keine Daten in der Datenbank vorhanden.</div>',
                content_type="text/html"
            )
        
        # Build data for each year (reuse logic from FamilyKostenReportView)
        years_data = []
        
        for year in all_years:
            sem1_months, sem2_months = get_semester_month_ranges(year)
            sem1_len = len(sem1_months)

            records = KarteiRecord.objects.filter(
                family_id=family_id,
                year=year,
                status__in=[RecordStatus.NORMAL, RecordStatus.PENDING],
            ).order_by("child_name", "subject1")
            
            if not records.exists():
                continue
            
            has_pending_records = records.filter(status=RecordStatus.PENDING).exists()
            
            rows = []
            month_sums = [ZERO] * 12
            
            for record in records:
                row = self._build_record_row(
                    record,
                    year,
                    current_year,
                    current_month,
                    sem1_len,
                )
                rows.append(row)
                
                for i, month_val in enumerate(row["months"]):
                    if month_val is not None:
                        month_sums[i] += month_val
            
            if year == current_year:
                for i in range(current_month, 12):
                    month_sums[i] = None
            
            year_total = sum(v for v in month_sums if v is not None)
            
            years_data.append({
                "year": year,
                "rows": rows,
                "month_sums": month_sums,
                "month_sums_sem1": month_sums[:sem1_len],
                "month_sums_sem2": month_sums[sem1_len:],
                "month_names_sem1": MONTH_NAMES_DE[:sem1_len],
                "month_names_sem2": MONTH_NAMES_DE[sem1_len:],
                "year_total": year_total,
                "has_pending_records": has_pending_records,
            })
        
        if not years_data:
            return HttpResponse(
                f'<div class="alert alert-warning">Keine Datensätze für FamilyID "{family_id}" gefunden.</div>',
                content_type="text/html"
            )
        
        # Get parent name from first record
        first_record = years_data[0]["rows"][0]["record"] if years_data[0]["rows"] else None
        parent_name = first_record.parent_name if first_record else ""
        
        context = {
            "family_id": family_id,
            "parent_name": parent_name,
            "years_data": years_data,
            "month_names": MONTH_NAMES_DE,
        }
        
        return render(request, self.template_name, context)
    
    def _build_record_row(
        self,
        record: KarteiRecord,
        year: int,
        current_year: int,
        current_month: int,
        sem1_len: int,
    ) -> dict[str, Any]:
        """Build a row dict for a single record (same logic as FamilyKostenReportView)."""
        months = []
        for i in range(1, 13):
            field_name = f"month_{i}"
            value = getattr(record, field_name, None)
            
            if year == current_year and i > current_month:
                value = None
            elif value is None:
                value = ZERO
            
            months.append(value)
        
        row_total = sum(v for v in months if v is not None)
        
        return {
            "record": record,
            "child_name": record.child_name or "-",
            "subject1": record.subject1 or "-",
            "subject2": record.subject2 or "-",
            "months": months,
            "months_sem1": months[:sem1_len],
            "months_sem2": months[sem1_len:],
            "is_pending": record.status == RecordStatus.PENDING,
            "row_total": row_total,
        }
