"""
API views for the karteien app.

This module provides JSON API endpoints for:
- Month breakdown details for billing explanation
- Live search for record list filtering
- Autocomplete for parent/child names
- Prefill form data from existing records
- Subject dependents (teachers/prices) for wizard filtering
"""

from __future__ import annotations

from datetime import date

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.db.models import Q
from django.http import JsonResponse, HttpRequest
from django.shortcuts import get_object_or_404
from django.template.loader import render_to_string
from django.views.decorators.http import require_GET

from apps.catalog.models import PriceOption, Subject, Teacher, TeachingAssignment

from .billing import get_month_breakdown
from .models import KarteiRecord, RecordStatus, MonthsMode


# =============================================================================
# Autocomplete and Prefill API
# =============================================================================

# Number of years to search back for fallback
FALLBACK_YEARS_DEPTH = 2

# Maximum results for autocomplete
AUTOCOMPLETE_LIMIT = 10


def _check_can_edit_kartei(user) -> JsonResponse | None:
    """Check if user can edit kartei. Returns JsonResponse error or None if ok."""
    if not user.can_edit_kartei:
        return JsonResponse({
            'error': 'Access denied',
            'code': 'ACCESS_DENIED',
        }, status=403)
    return None


def _search_records_by_field(field: str, query: str, target_year: int, limit: int) -> dict:
    """
    Search KarteiRecord by a given field (parent_name or child_name).
    
    First searches in target_year, then falls back to previous years if no results.
    
    Returns dict with:
    - target_year: the requested year
    - results_year: the year where results were found
    - is_fallback: True if results are from a different year
    - results: list of matching records with label and pkid
    """
    filter_kwargs = {f"{field}__icontains": query}
    
    # Try target year first
    qs = KarteiRecord.objects.filter(
        year=target_year,
        **filter_kwargs
    ).order_by(field, 'family_id')[:limit]
    
    results = list(qs)
    results_year = target_year
    is_fallback = False
    
    # If no results in target year, try fallback to previous years
    if not results:
        for offset in range(1, FALLBACK_YEARS_DEPTH + 1):
            fallback_year = target_year - offset
            qs = KarteiRecord.objects.filter(
                year=fallback_year,
                **filter_kwargs
            ).order_by(field, 'family_id')[:limit]
            
            results = list(qs)
            if results:
                results_year = fallback_year
                is_fallback = True
                break
    
    # Build result labels
    result_list = []
    for record in results:
        if field == 'parent_name':
            name = record.parent_name
        else:
            name = record.child_name
        
        # Include year prefix for fallback results
        if is_fallback:
            label = f"[{results_year}] {name} (FamilyID {record.family_id})"
        else:
            label = f"{name} (FamilyID {record.family_id})"
        
        result_list.append({
            'record_pkid': record.pkid,
            'label': label,
            'family_id': record.family_id,
            'parent_name': record.parent_name,
            'child_name': record.child_name,
        })
    
    return {
        'target_year': target_year,
        'results_year': results_year,
        'is_fallback': is_fallback,
        'results': result_list,
    }


@login_required
@require_GET
def autocomplete_parents_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint for parent name autocomplete.
    
    GET /api/karteien/autocomplete/parents/?year=2025&q=anna
    
    Searches parent_name field. First in target year, then fallback to previous years.
    
    Returns JSON with:
    - target_year: requested year
    - results_year: year where results were found
    - is_fallback: true if results are from a different year
    - results: array of {record_pkid, label, family_id, parent_name, child_name}
    """
    user = request.user
    
    # Access control
    error = _check_can_edit_kartei(user)
    if error:
        return error
    
    # Get parameters
    query = request.GET.get('q', '').strip()
    year_str = request.GET.get('year', '')
    
    # Validate query length
    if len(query) < 2:
        return JsonResponse({
            'target_year': None,
            'results_year': None,
            'is_fallback': False,
            'results': [],
        })
    
    # Parse year
    try:
        target_year = int(year_str)
    except (ValueError, TypeError):
        target_year = date.today().year
    
    # Search
    result = _search_records_by_field(
        field='parent_name',
        query=query,
        target_year=target_year,
        limit=AUTOCOMPLETE_LIMIT,
    )
    
    return JsonResponse(result)


@login_required
@require_GET
def autocomplete_children_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint for child name autocomplete.
    
    GET /api/karteien/autocomplete/children/?year=2025&q=max
    
    Searches child_name field. First in target year, then fallback to previous years.
    
    Returns JSON with:
    - target_year: requested year
    - results_year: year where results were found
    - is_fallback: true if results are from a different year
    - results: array of {record_pkid, label, family_id, parent_name, child_name}
    """
    user = request.user
    
    # Access control
    error = _check_can_edit_kartei(user)
    if error:
        return error
    
    # Get parameters
    query = request.GET.get('q', '').strip()
    year_str = request.GET.get('year', '')
    
    # Validate query length
    if len(query) < 2:
        return JsonResponse({
            'target_year': None,
            'results_year': None,
            'is_fallback': False,
            'results': [],
        })
    
    # Parse year
    try:
        target_year = int(year_str)
    except (ValueError, TypeError):
        target_year = date.today().year
    
    # Search
    result = _search_records_by_field(
        field='child_name',
        query=query,
        target_year=target_year,
        limit=AUTOCOMPLETE_LIMIT,
    )
    
    return JsonResponse(result)


@login_required
@require_GET
def prefill_from_record_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint for getting prefill data from a specific record.
    
    GET /api/karteien/prefill-from-record/?record_pkid=123&target_year=2025
    
    Returns JSON with:
    - target_year: requested year
    - source_year: year of the source record
    - is_fallback: true if source_year != target_year
    - data: form field values to prefill
    """
    user = request.user
    
    # Access control
    error = _check_can_edit_kartei(user)
    if error:
        return error
    
    # Get parameters
    record_pkid_str = request.GET.get('record_pkid', '')
    target_year_str = request.GET.get('target_year', '')
    
    # Validate record_pkid
    try:
        record_pkid = int(record_pkid_str)
    except (ValueError, TypeError):
        return JsonResponse({
            'error': 'Invalid or missing record_pkid',
            'code': 'INVALID_RECORD_PKID',
        }, status=400)
    
    # Parse target year
    try:
        target_year = int(target_year_str)
    except (ValueError, TypeError):
        target_year = date.today().year
    
    # Get record
    try:
        record = KarteiRecord.objects.get(pkid=record_pkid)
    except KarteiRecord.DoesNotExist:
        return JsonResponse({
            'error': 'Record not found',
            'code': 'NOT_FOUND',
        }, status=404)
    
    # Build prefill data
    source_year = record.year
    is_fallback = source_year != target_year
    
    # Format birthdate for form input (YYYY-MM-DD)
    birthdate_str = ''
    if record.birthdate:
        birthdate_str = record.birthdate.isoformat()
    
    data = {
        'family_id': record.family_id,
        'parent_name': record.parent_name,
        'child_name': record.child_name,
        'birthdate': birthdate_str,
        'address': record.address,
        'phone': record.phone,
        'mobile': record.mobile,
        'email': record.email,
        'sepa_marker': record.sepa_marker if hasattr(record, 'sepa_marker') else '',
    }
    
    return JsonResponse({
        'target_year': target_year,
        'source_year': source_year,
        'is_fallback': is_fallback,
        'data': data,
    })


@login_required
@require_GET
def month_breakdown_api(request: HttpRequest, pk: int) -> JsonResponse:
    """
    API endpoint for getting month billing breakdown.
    
    GET /api/karteien/<pk>/month-breakdown/?month=5
    
    Access control:
    - Admin/Operator: Can view any record they can normally access
    - Superadmin: Can view any record
    - User: Can view records (read-only access)
    
    Returns:
        JSON with breakdown details or error.
    """
    user = request.user
    
    # Get month parameter
    month_str = request.GET.get('month')
    if not month_str:
        return JsonResponse({
            'error': 'Missing month parameter',
            'code': 'MISSING_MONTH',
        }, status=400)
    
    try:
        month = int(month_str)
        if month < 1 or month > 12:
            raise ValueError("Month must be 1-12")
    except ValueError:
        return JsonResponse({
            'error': 'Invalid month parameter. Must be 1-12.',
            'code': 'INVALID_MONTH',
        }, status=400)
    
    # Get record
    record = get_object_or_404(KarteiRecord, pk=pk)
    
    # Access control based on role
    # Admin/Operator can edit kartei, so they can view breakdown
    # Superadmin can view everything
    # User role can view records (read-only)
    can_access = (
        user.can_edit_kartei or  # Admin/Operator
        user.is_superadmin or    # Superadmin
        user.is_user_role        # User (read-only access)
    )
    
    if not can_access:
        return JsonResponse({
            'error': 'Access denied',
            'code': 'ACCESS_DENIED',
        }, status=403)
    
    # Get breakdown
    breakdown = get_month_breakdown(record, month)
    
    return JsonResponse(breakdown)


@login_required
@require_GET
def live_search_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint for live search filtering of Kartei records.
    
    GET /api/karteien/live-search/?year=2025&family_id=...&parent=...&child=...&status=...&page=1
    
    Access control:
    - Admin, Operator, Superadmin, and User can access (same as list view)
    
    Returns:
        JSON with:
        - rows_html: HTML for table body
        - pagination_html: HTML for pagination
        - total_count: Number of results
    """
    user = request.user
    
    # Access control: Users who can view Kartei (Admin/Operator/Superadmin/User)
    if not (user.can_edit_kartei or user.is_superadmin or user.is_user_role):
        return JsonResponse({
            'error': 'Access denied',
            'code': 'ACCESS_DENIED',
        }, status=403)
    
    # Build queryset with filters (same logic as KarteiRecordListView)
    qs = KarteiRecord.objects.all()
    
    # Default to current year if not specified
    year = request.GET.get("year")
    if year:
        try:
            qs = qs.filter(year=int(year))
        except ValueError:
            pass
    else:
        qs = qs.filter(year=date.today().year)
    
    # Filter by FamilyID
    family_id = request.GET.get("family_id")
    if family_id:
        qs = qs.filter(family_id__icontains=family_id)
    
    # Filter by Parent
    parent = request.GET.get("parent")
    if parent:
        qs = qs.filter(parent_name__icontains=parent)
    
    # Filter by Child
    child = request.GET.get("child")
    if child:
        qs = qs.filter(child_name__icontains=child)
    
    # Filter by Status
    status = request.GET.get("status")
    if status:
        if status == "PENDING":
            qs = qs.filter(status=RecordStatus.PENDING)
        elif status == "DECLINED":
            qs = qs.filter(status=RecordStatus.DECLINED)
        elif status == "NORMAL":
            qs = qs.filter(status=RecordStatus.NORMAL)
    
    # Filter by contract type (monthly/yearly)
    contract_type = request.GET.get("contract_type")
    if contract_type:
        if contract_type == "monthly":
            qs = qs.filter(is_monthly_contract=True)
        elif contract_type == "yearly":
            qs = qs.filter(is_monthly_contract=False)
    
    # Filter by contract status (active/terminated)
    contract_status = request.GET.get("contract_status")
    if contract_status:
        if contract_status == "active":
            qs = qs.filter(is_contract_terminated=False)
        elif contract_status == "terminated":
            qs = qs.filter(is_contract_terminated=True)
    
    # Order results
    qs = qs.order_by("family_id", "parent_name", "child_name")
    
    # Paginate (same as list view: 50 per page)
    paginator = Paginator(qs, 50)
    page = request.GET.get("page", 1)
    
    try:
        page_obj = paginator.page(page)
    except PageNotAnInteger:
        page_obj = paginator.page(1)
    except EmptyPage:
        page_obj = paginator.page(paginator.num_pages)
    
    # Build context for templates
    current_year = year if year else date.today().year
    context = {
        'records': page_obj.object_list,
        'page_obj': page_obj,
        'current_year': current_year,
        'user': user,
    }
    
    # Render partial templates
    rows_html = render_to_string(
        'karteien/_record_list_table.html',
        context,
        request=request
    )
    pagination_html = render_to_string(
        'karteien/_record_list_pagination.html',
        context,
        request=request
    )
    
    return JsonResponse({
        'rows_html': rows_html,
        'pagination_html': pagination_html,
        'total_count': paginator.count,
    })


# =============================================================================
# Subject Dependents API (for wizard dynamic filtering)
# =============================================================================

def _check_is_admin(user) -> JsonResponse | None:
    """Check if user is admin. Returns JsonResponse error or None if ok."""
    if not user.is_authenticated or not user.is_admin_role:
        return JsonResponse({
            'error': 'Access denied',
            'code': 'ACCESS_DENIED',
        }, status=403)
    return None


@login_required
@require_GET
def subject_dependents_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint to get teachers and prices for a given subject and year.
    
    Used by the New Family Wizard to dynamically filter dropdowns.
    
    Query params:
    - year: The year to filter by (required)
    - subject_id: The subject ID to filter by (required)
    
    Returns JSON:
    {
        "teachers": [{"id": 1, "label": "Nachname, Vorname"}, ...],
        "prices": [{"id": 1, "label": "123.45 € (comment)"}, ...]
    }
    
    Admin-only access.
    """
    # Check admin access
    error = _check_is_admin(request.user)
    if error:
        return error
    
    # Parse parameters
    try:
        year = int(request.GET.get('year', 0))
        subject_id = int(request.GET.get('subject_id', 0))
    except (TypeError, ValueError):
        return JsonResponse({
            'error': 'Invalid parameters',
            'code': 'INVALID_PARAMS',
        }, status=400)
    
    if not year or not subject_id:
        return JsonResponse({
            'error': 'Missing required parameters: year and subject_id',
            'code': 'MISSING_PARAMS',
        }, status=400)
    
    # Get teachers linked to this subject for this year via TeachingAssignment
    teacher_ids = TeachingAssignment.objects.filter(
        year=year,
        subject_id=subject_id,
        is_active=True,
        teacher__is_active=True,
    ).values_list('teacher_id', flat=True).distinct()
    
    teachers = Teacher.objects.filter(
        id__in=teacher_ids,
        is_active=True,
    ).order_by('last_name', 'first_name')
    
    teachers_list = [
        {'id': t.id, 'label': str(t)}
        for t in teachers
    ]
    
    # Get prices for this subject and year
    prices = PriceOption.objects.filter(
        year=year,
        subject_id=subject_id,
        is_active=True,
    ).order_by('amount')
    
    prices_list = [
        {
            'id': p.id,
            'label': f"{p.amount} € ({p.comment})" if p.comment else f"{p.amount} €"
        }
        for p in prices
    ]
    
    return JsonResponse({
        'teachers': teachers_list,
        'prices': prices_list,
    })


# =============================================================================
# Billing Preview API (for live preview in edit form)
# =============================================================================

@login_required
def billing_preview_api(request: HttpRequest, pk: int) -> JsonResponse:
    """
    API endpoint for billing preview (live preview in edit form).
    
    POST /api/karteien/<pk>/billing-preview/
    
    Calculates what the month values would be based on form inputs,
    without saving anything. Returns preview data with source information
    for each month (legacy vs calculated).
    
    Access: Admin only (users who can edit kartei).
    
    Request body (JSON or form data):
    - subject1_ref: Subject reference ID for semester 1
    - price1_ref: Price reference ID for semester 1
    - start_month_1: Start month for semester 1
    - apply_from_month_1: Apply price changes from this month (semester 1)
    - subject2_ref: Subject reference ID for semester 2
    - price2_ref: Price reference ID for semester 2
    - start_month_2: Start month for semester 2
    - apply_from_month_2: Apply price changes from this month (semester 2)
    - hours_month_1..hours_month_12: Hours for hourly subjects
    
    Returns JSON:
    {
        "months": {
            "1": {"value": "46.00", "source": "legacy"|"calculated", "reason": "..."},
            ...
        },
        "changed_months": [7, 8, 9, ...],
        "warnings": ["..."]
    }
    """
    from decimal import Decimal
    import json
    
    from apps.catalog.models import FamilyDiscount, RecordDiscount
    
    from .billing import (
        is_per_hour_subject,
        ZERO, round_money_up
    )
    
    user = request.user
    
    # Access control - Admin only
    error = _check_can_edit_kartei(user)
    if error:
        return error
    
    # Get record
    record = get_object_or_404(KarteiRecord, pk=pk)
    
    # Parse request data (supports both JSON and form data)
    if request.content_type and 'application/json' in request.content_type:
        try:
            data = json.loads(request.body)
        except json.JSONDecodeError:
            return JsonResponse({
                'error': 'Invalid JSON',
                'code': 'INVALID_JSON',
            }, status=400)
    else:
        data = request.POST.dict() if request.method == 'POST' else request.GET.dict()
    
    # Helper to parse int or None
    def parse_int_or_none(val):
        if val is None or val == '' or val == 'None':
            return None
        try:
            return int(val)
        except (ValueError, TypeError):
            return None
    
    # Helper to parse Decimal or None
    def parse_decimal_or_none(val):
        if val is None or val == '' or val == 'None':
            return None
        try:
            return Decimal(str(val))
        except:
            return None
    
    # Extract form parameters
    subject1_ref_id = parse_int_or_none(data.get('subject1_ref'))
    price1_ref_id = parse_int_or_none(data.get('price1_ref'))
    start_month_1 = parse_int_or_none(data.get('start_month_1')) or 1
    apply_from_month_1 = parse_int_or_none(data.get('apply_from_month_1'))
    
    subject2_ref_id = parse_int_or_none(data.get('subject2_ref'))
    price2_ref_id = parse_int_or_none(data.get('price2_ref'))
    start_month_2 = parse_int_or_none(data.get('start_month_2')) or 7
    apply_from_month_2 = parse_int_or_none(data.get('apply_from_month_2'))
    
    # Extract hours
    hours_amounts = {}
    for m in range(1, 13):
        key = f'hours_month_{m}'
        val = data.get(key)
        if val is not None and val != '':
            hours_amounts[f'month_{m}'] = val
    
    # Build a "virtual" record state for calculation
    # We'll temporarily modify record attributes for preview (not saved)
    original_values = {}
    
    # Store original month values for comparison
    for m in range(1, 13):
        original_values[f'month_{m}'] = getattr(record, f'month_{m}', None)
    
    # Get subject names for preview
    subject1_name = None
    subject2_name = None
    price1_amount = None
    price2_amount = None
    price1_is_legacy = False
    price2_is_legacy = False
    
    if subject1_ref_id:
        try:
            subject1 = Subject.objects.get(pk=subject1_ref_id)
            subject1_name = subject1.name
        except Subject.DoesNotExist:
            pass
    else:
        subject1_name = record.subject1
    
    if subject2_ref_id:
        try:
            subject2 = Subject.objects.get(pk=subject2_ref_id)
            subject2_name = subject2.name
        except Subject.DoesNotExist:
            pass
    else:
        subject2_name = record.subject2
    
    if price1_ref_id:
        try:
            price1_option = PriceOption.objects.get(pk=price1_ref_id)
            price1_amount = price1_option.amount
        except PriceOption.DoesNotExist:
            pass
    else:
        price1_amount = record.price1
        price1_is_legacy = True
    
    if price2_ref_id:
        try:
            price2_option = PriceOption.objects.get(pk=price2_ref_id)
            price2_amount = price2_option.amount
        except PriceOption.DoesNotExist:
            pass
    else:
        price2_amount = record.price2
        price2_is_legacy = True
    
    # Build result structure
    months_result = {}
    changed_months = []
    warnings = []
    
    # Check if record is LEGACY mode
    is_legacy_mode = record.months_mode == MonthsMode.LEGACY
    
    # Determine which months would be affected by changes
    # For LEGACY mode, we need to track which months are being recalculated
    
    # Fetch discounts
    family_discounts = list(
        FamilyDiscount.objects.filter(
            year=record.year,
            family_id=record.family_id,
        ).select_related('discount')
    )
    
    record_discounts = list(
        record.record_discounts.select_related('discount')
    ) if record.pk else []
    
    # For each month, determine source and calculate value
    for month_num in range(1, 13):
        field_key = f'month_{month_num}'
        
        # Determine semester
        semester = 1 if month_num <= 6 else 2
        
        # Get subject/price for this semester
        if semester == 1:
            subj_name = subject1_name
            price = price1_amount
            start_m = start_month_1
            apply_from = apply_from_month_1
            is_price_legacy = price1_is_legacy
        else:
            subj_name = subject2_name
            price = price2_amount
            start_m = start_month_2
            apply_from = apply_from_month_2
            is_price_legacy = price2_is_legacy
        
        # Get hours for this month
        hours_val = hours_amounts.get(field_key)
        if hours_val is None:
            # Try to get from record's hours_amounts
            record_hours = getattr(record, 'hours_amounts', None) or {}
            hours_val = record_hours.get(field_key)
        
        # Determine if this month should be recalculated or kept as legacy
        current_value = original_values[field_key]
        
        # Month is "affected" if:
        # 1. apply_from is set and month >= apply_from
        # 2. Or if this is a new calculation (AUTO mode)
        is_affected = False
        reason_parts = []
        
        if is_legacy_mode:
            # For LEGACY records, only recalculate if:
            # - apply_from is set and month >= apply_from
            # - OR hours were provided for hourly subjects
            if apply_from is not None and month_num >= apply_from:
                is_affected = True
                reason_parts.append(f"Neuberechnung ab Monat {apply_from}")
            elif is_per_hour_subject(subj_name) and hours_val:
                is_affected = True
                reason_parts.append("Stunden eingegeben")
        else:
            # AUTO mode - all months are calculated
            is_affected = True
        
        if not is_affected:
            # Keep legacy value
            months_result[str(month_num)] = {
                'value': str(current_value) if current_value else '0.00',
                'source': 'legacy',
                'reason': 'Legacy-Wert, nicht neu berechnet',
            }
            continue
        
        # Calculate new value for this month
        if month_num < start_m:
            # Before start month
            new_value = ZERO
            reason = f"Vor Startmonat ({start_m})"
        elif not subj_name or price is None:
            # No subject or price
            new_value = ZERO
            reason = "Kein Fach oder Preis ausgewählt"
        else:
            # Calculate base amount
            if is_per_hour_subject(subj_name):
                # Hourly subject
                hours_decimal = parse_decimal_or_none(hours_val) or ZERO
                if hours_decimal == ZERO:
                    new_value = ZERO
                    reason = f"Stundenfach ({subj_name}), keine Stunden eingegeben"
                else:
                    base = hours_decimal * Decimal(str(price))
                    new_value = round_money_up(base)
                    reason_parts.append(f"Stundenfach: {hours_decimal} Std × {price} €")
            else:
                # Monthly subject
                base = Decimal(str(price))
                new_value = round_money_up(base)
                reason_parts.append(f"Monatspreis: {price} €")
            
            # Add price source info
            if is_price_legacy:
                reason_parts.append("(Basis aus Legacy-Preis, nicht Katalog)")
            
            # Apply discounts (simplified - for full breakdown use month_breakdown_api)
            # We'll note if discounts apply but not recalculate in detail here
            if new_value > ZERO and (family_discounts or record_discounts):
                reason_parts.append("Rabatte werden angewendet")
            
            reason = "; ".join(reason_parts) if reason_parts else "Berechnet"
        
        # Check if value changed
        if current_value is not None:
            try:
                current_dec = Decimal(str(current_value))
                if current_dec != new_value:
                    changed_months.append(month_num)
            except:
                changed_months.append(month_num)
        else:
            changed_months.append(month_num)
        
        months_result[str(month_num)] = {
            'value': str(new_value),
            'source': 'calculated',
            'reason': reason,
        }
    
    # Add warnings
    if is_legacy_mode and not any(parse_int_or_none(data.get(f'apply_from_month_{s}')) for s in [1, 2]):
        warnings.append("Legacy-Modus: Wählen Sie 'Preis anwenden ab Monat', um Monate neu zu berechnen.")
    
    return JsonResponse({
        'months': months_result,
        'changed_months': sorted(changed_months),
        'warnings': warnings,
        'record_months_mode': record.months_mode,
    })
