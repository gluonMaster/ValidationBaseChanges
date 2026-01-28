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
from django.db.models import Count, Exists, OuterRef, Q
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


def _parse_months_csv_for_preview(csv_str: str) -> set[int]:
    """
    Parse a months CSV string into a set of month numbers.
    
    Args:
        csv_str: CSV string like "7,8,12" or empty string.
        
    Returns:
        Set of month numbers. Empty set if input is empty or invalid.
    """
    if not csv_str or not csv_str.strip():
        return set()
    
    months_set: set[int] = set()
    for part in csv_str.split(','):
        part = part.strip()
        if part:
            try:
                months_set.add(int(part))
            except (ValueError, TypeError):
                pass
    
    return months_set


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
    year_param = request.GET.get("year")
    system_year = date.today().year
    
    if year_param:
        try:
            selected_year = int(year_param)
            # Save to session for persistence (AJAX calls update session)
            request.session["karteien_selected_year"] = selected_year
        except ValueError:
            # Invalid year - fall back to session or system year
            selected_year = request.session.get("karteien_selected_year", system_year)
    else:
        # No year in GET - use session or default to system year
        selected_year = request.session.get("karteien_selected_year", system_year)
    
    # Filter by selected year
    qs = qs.filter(year=selected_year)
    
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
    
    # Filter by contract status (active/terminated/active_sepa/terminated_sepa)
    contract_status = request.GET.get("contract_status")
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
    subject_query = request.GET.get("subject_query", "").strip()
    subject_semester = request.GET.get("subject_semester", "")
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
    teacher_query = request.GET.get("teacher_query", "").strip()
    teacher_semester = request.GET.get("teacher_semester", "")
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
    context = {
        'records': page_obj.object_list,
        'page_obj': page_obj,
        'current_year': selected_year,
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
    for each month (legacy vs calculated). Applies all discounts (family + record).
    
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
    
    from apps.catalog.models import FamilyDiscount, RecordDiscount, DiscountKind
    
    from .billing import (
        is_per_hour_subject,
        is_nachhilfe_subject,
        build_base_amounts,
        calculate_month_values,
        collect_discounts_for_month,
        get_semester_for_month,
        get_subject_name_for_semester,
        ZERO, round_money_up, MAX_PERCENT_DISCOUNT
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
    end_month_1 = parse_int_or_none(data.get('end_month_1'))  # None = until end of semester
    months_csv_1 = data.get('months_csv_1', '') or ''  # CSV of specific months
    apply_from_month_1 = parse_int_or_none(data.get('apply_from_month_1'))
    apply_to_month_1 = parse_int_or_none(data.get('apply_to_month_1'))
    
    subject2_ref_id = parse_int_or_none(data.get('subject2_ref'))
    price2_ref_id = parse_int_or_none(data.get('price2_ref'))
    start_month_2 = parse_int_or_none(data.get('start_month_2')) or 7
    end_month_2 = parse_int_or_none(data.get('end_month_2'))  # None = until end of semester
    months_csv_2 = data.get('months_csv_2', '') or ''  # CSV of specific months
    apply_from_month_2 = parse_int_or_none(data.get('apply_from_month_2'))
    apply_to_month_2 = parse_int_or_none(data.get('apply_to_month_2'))
    
    # Extract hours
    hours_amounts = {}
    for m in range(1, 13):
        key = f'hours_month_{m}'
        val = data.get(key)
        if val is not None and val != '':
            hours_amounts[f'month_{m}'] = val
    
    # Extract discounts_disabled fields
    discounts_disabled_str = data.get('discounts_disabled', '')
    discounts_disabled = discounts_disabled_str in ('true', 'True', '1', True)
    
    # Parse discounts_disabled_months from von/bis/csv
    discounts_disabled_months: list[int] = []
    
    von_str = data.get('discounts_disabled_von', '')
    bis_str = data.get('discounts_disabled_bis', '')
    csv_str = data.get('discounts_disabled_csv', '')
    
    von_month = parse_int_or_none(von_str)
    bis_month = parse_int_or_none(bis_str)
    
    months_set: set[int] = set()
    
    # Add range months
    if von_month is not None and bis_month is not None:
        if 1 <= von_month <= 12 and 1 <= bis_month <= 12:
            start = min(von_month, bis_month)
            end = max(von_month, bis_month)
            for m in range(start, end + 1):
                months_set.add(m)
    elif von_month is not None and 1 <= von_month <= 12:
        months_set.add(von_month)
    elif bis_month is not None and 1 <= bis_month <= 12:
        months_set.add(bis_month)
    
    # Parse CSV
    if csv_str:
        for part in csv_str.split(','):
            part = part.strip()
            if part:
                try:
                    month = int(part)
                    if 1 <= month <= 12:
                        months_set.add(month)
                except (ValueError, TypeError):
                    pass
    
    discounts_disabled_months = sorted(months_set)
    
    # Extract contract status fields for termination preview
    contract_status_str = data.get('contract_status', '')
    contract_status_effective_month_str = data.get('contract_status_effective_month', '')
    
    # Parse contract_status_effective_month
    preview_is_terminated = contract_status_str == 'terminated'
    preview_terminated_from_month = parse_int_or_none(contract_status_effective_month_str)
    if preview_terminated_from_month is not None:
        if not (1 <= preview_terminated_from_month <= 12):
            preview_terminated_from_month = None
    
    # Store original month values for comparison
    original_values = {}
    for m in range(1, 13):
        original_values[f'month_{m}'] = getattr(record, f'month_{m}', None)
    
    # Check if record is LEGACY mode
    is_legacy_mode = record.months_mode == MonthsMode.LEGACY
    
    # Fetch discounts for the record
    family_discounts = list(
        FamilyDiscount.objects.filter(
            year=record.year,
            family_id=record.family_id,
        ).select_related('discount')
    )
    
    record_discounts = list(
        record.record_discounts.select_related('discount')
    ) if record.pk else []
    
    # Determine which months would be affected by changes
    # For LEGACY mode, compute affected_months using same logic as detect_meaningful_changes
    # For AUTO mode, ALL months are always calculated
    affected_months = set()
    
    if is_legacy_mode:
        # LEGACY mode: compute touched months based on meaningful changes
        # Mirrors logic from detect_meaningful_changes for consistency with save behavior
        
        # Helper to normalize subject names for comparison
        def _normalize_subject_name(name: str | None) -> str:
            if not name:
                return ''
            return name.strip().lower()
        
        # Get original values from record for comparison
        old_price1_ref_id = record.price1_ref_id
        old_price2_ref_id = record.price2_ref_id
        old_subject1_ref_id = record.subject1_ref_id
        old_subject2_ref_id = record.subject2_ref_id
        old_start_1 = record.start_month_1 or 1
        old_start_2 = record.start_month_2 or 7
        old_discounts_disabled = record.discounts_disabled
        old_is_terminated = record.is_contract_terminated
        old_hours_amounts = record.hours_amounts or {}
        
        # Check price1_ref change - touch months from apply_from_month_1 (default=1) to 6
        # Any change to price1_ref (including linking to same legacy value) is meaningful
        if price1_ref_id != old_price1_ref_id:
            # Use apply_from_month_1 if provided, else default to 1
            effective_apply_from = apply_from_month_1 if apply_from_month_1 is not None else 1
            # Also respect apply_to
            for m in range(effective_apply_from, 7):
                if apply_to_month_1 is None or m <= apply_to_month_1:
                    affected_months.add(m)
        
        # Check price2_ref change - touch months from apply_from_month_2 (default=7) to 12
        # Any change to price2_ref (including linking to same legacy value) is meaningful
        if price2_ref_id != old_price2_ref_id:
            # Use apply_from_month_2 if provided, else default to 7
            effective_apply_from = apply_from_month_2 if apply_from_month_2 is not None else 7
            # Also respect apply_to
            for m in range(effective_apply_from, 13):
                if apply_to_month_2 is None or m <= apply_to_month_2:
                    affected_months.add(m)
        
        # Check subject1_ref change - touch months from min(old_start, new_start) to 6
        # Any change to subject1_ref (including linking to same legacy value) is meaningful
        if subject1_ref_id != old_subject1_ref_id:
            effective_start = min(old_start_1, start_month_1)
            affected_months.update(range(effective_start, 7))
        
        # Check subject2_ref change - touch months from min(old_start, new_start) to 12
        # Any change to subject2_ref (including linking to same legacy value) is meaningful
        if subject2_ref_id != old_subject2_ref_id:
            effective_start = min(old_start_2, start_month_2)
            affected_months.update(range(effective_start, 13))
        
        # Check start_month_1 change - only touch months between old and new
        if start_month_1 != old_start_1:
            min_start = min(old_start_1, start_month_1)
            max_start = max(old_start_1, start_month_1)
            affected_months.update(range(min_start, max_start))
        
        # Check start_month_2 change - only touch months between old and new
        if start_month_2 != old_start_2:
            min_start = min(old_start_2, start_month_2)
            max_start = max(old_start_2, start_month_2)
            affected_months.update(range(min_start, max_start))
        
        # Check end_month_1 change - touch months between old and new end
        old_end_1 = record.end_month_1
        if end_month_1 != old_end_1:
            eff_old_end_1 = old_end_1 if old_end_1 is not None else 6
            eff_new_end_1 = end_month_1 if end_month_1 is not None else 6
            min_end = min(eff_old_end_1, eff_new_end_1)
            max_end = max(eff_old_end_1, eff_new_end_1)
            affected_months.update(range(min_end + 1, max_end + 1))
        
        # Check end_month_2 change - touch months between old and new end
        old_end_2 = record.end_month_2
        if end_month_2 != old_end_2:
            eff_old_end_2 = old_end_2 if old_end_2 is not None else 12
            eff_new_end_2 = end_month_2 if end_month_2 is not None else 12
            min_end = min(eff_old_end_2, eff_new_end_2)
            max_end = max(eff_old_end_2, eff_new_end_2)
            affected_months.update(range(min_end + 1, max_end + 1))
        
        # Check months_csv_1 change - touch symmetric difference
        old_csv_1 = record.months_csv_1 or ''
        if months_csv_1 != old_csv_1:
            old_months_1 = _parse_months_csv_for_preview(old_csv_1)
            new_months_1 = _parse_months_csv_for_preview(months_csv_1)
            if not old_months_1 or not new_months_1:
                affected_months.update(range(1, 7))
            else:
                affected_months.update(old_months_1.symmetric_difference(new_months_1))
        
        # Check months_csv_2 change - touch symmetric difference
        old_csv_2 = record.months_csv_2 or ''
        if months_csv_2 != old_csv_2:
            old_months_2 = _parse_months_csv_for_preview(old_csv_2)
            new_months_2 = _parse_months_csv_for_preview(months_csv_2)
            if not old_months_2 or not new_months_2:
                affected_months.update(range(7, 13))
            else:
                affected_months.update(old_months_2.symmetric_difference(new_months_2))
        
        # Check discounts_disabled change - affects all months
        if discounts_disabled != old_discounts_disabled:
            affected_months.update(range(1, 13))
        
        # Check contract termination change - affects all months
        if contract_status_str:
            if preview_is_terminated != old_is_terminated:
                affected_months.update(range(1, 13))
        
        # Check hours changes for per-hour subjects
        from decimal import Decimal as D
        for month_num in range(1, 13):
            field_key = f'month_{month_num}'
            new_hours = hours_amounts.get(field_key, '0.00')
            old_hours = old_hours_amounts.get(field_key, '0.00')
            try:
                new_val = D(str(new_hours or '0.00'))
                old_val = D(str(old_hours or '0.00'))
                if new_val != old_val:
                    affected_months.add(month_num)
            except (ValueError, TypeError):
                pass
    else:
        # AUTO mode: ALL months are always calculated
        # apply_from/apply_to only affects base_amounts (price history), not which months are shown
        affected_months = set(range(1, 13))
    
    # Temporarily modify record attributes for preview (not saved)
    # Store originals to restore later (though not strictly necessary since we don't save)
    orig_subject1_ref_id = record.subject1_ref_id
    orig_subject2_ref_id = record.subject2_ref_id
    orig_price1_ref_id = record.price1_ref_id
    orig_price2_ref_id = record.price2_ref_id
    orig_start_month_1 = record.start_month_1
    orig_start_month_2 = record.start_month_2
    orig_end_month_1 = record.end_month_1
    orig_end_month_2 = record.end_month_2
    orig_months_csv_1 = record.months_csv_1
    orig_months_csv_2 = record.months_csv_2
    orig_hours_amounts = record.hours_amounts
    orig_base_amounts = record.base_amounts
    orig_discounts_disabled = record.discounts_disabled
    orig_discounts_disabled_months = record.discounts_disabled_months
    orig_is_contract_terminated = record.is_contract_terminated
    orig_contract_terminated_from_month = record.contract_terminated_from_month
    
    # Apply preview parameters to record (in-memory only)
    if subject1_ref_id is not None:
        record.subject1_ref_id = subject1_ref_id
    if subject2_ref_id is not None:
        record.subject2_ref_id = subject2_ref_id
    if price1_ref_id is not None:
        record.price1_ref_id = price1_ref_id
    if price2_ref_id is not None:
        record.price2_ref_id = price2_ref_id
    record.start_month_1 = start_month_1
    record.start_month_2 = start_month_2
    # Apply end_month values for preview (if provided in request, or keep original)
    # end_month_* can be None (meaning "until end of semester")
    if 'end_month_1' in data:
        record.end_month_1 = end_month_1
    if 'end_month_2' in data:
        record.end_month_2 = end_month_2
    # Apply months_csv values for preview
    if 'months_csv_1' in data:
        record.months_csv_1 = months_csv_1
    if 'months_csv_2' in data:
        record.months_csv_2 = months_csv_2
    
    # Apply discounts_disabled fields
    record.discounts_disabled = discounts_disabled
    record.discounts_disabled_months = discounts_disabled_months
    
    # Apply contract termination preview
    # Only override if contract_status was provided in the request
    if contract_status_str:
        record.is_contract_terminated = preview_is_terminated
        record.contract_terminated_from_month = preview_terminated_from_month if preview_is_terminated else None
    
    # Merge hours_amounts with record's existing hours
    merged_hours = dict(orig_hours_amounts or {})
    merged_hours.update(hours_amounts)
    record.hours_amounts = merged_hours
    
    # Build base amounts using the billing module
    base_amounts = build_base_amounts(
        record,
        apply_from_month_1=apply_from_month_1,
        apply_from_month_2=apply_from_month_2,
        apply_to_month_1=apply_to_month_1,
        apply_to_month_2=apply_to_month_2,
        hours_amounts=merged_hours,
    )
    
    # Temporarily set base_amounts on record for calculate_month_values
    record.base_amounts = base_amounts
    
    # Calculate final values with discounts applied
    final_values, calc_flags = calculate_month_values(
        record,
        family_discounts=family_discounts,
        record_discounts=record_discounts,
        base_amounts=base_amounts,
    )
    
    # Restore original record state
    record.subject1_ref_id = orig_subject1_ref_id
    record.subject2_ref_id = orig_subject2_ref_id
    record.price1_ref_id = orig_price1_ref_id
    record.price2_ref_id = orig_price2_ref_id
    record.start_month_1 = orig_start_month_1
    record.start_month_2 = orig_start_month_2
    record.end_month_1 = orig_end_month_1
    record.end_month_2 = orig_end_month_2
    record.months_csv_1 = orig_months_csv_1
    record.months_csv_2 = orig_months_csv_2
    record.hours_amounts = orig_hours_amounts
    record.base_amounts = orig_base_amounts
    record.discounts_disabled = orig_discounts_disabled
    record.discounts_disabled_months = orig_discounts_disabled_months
    record.is_contract_terminated = orig_is_contract_terminated
    record.contract_terminated_from_month = orig_contract_terminated_from_month
    
    # Get price info for reason building
    price1_is_legacy = price1_ref_id is None and record.price1 is not None
    price2_is_legacy = price2_ref_id is None and record.price2 is not None
    
    # Get subject names for preview (using the values from form or record)
    def get_subject_name(semester: int) -> str | None:
        """Get subject name for a semester based on form inputs or record."""
        if semester == 1:
            if subject1_ref_id:
                try:
                    return Subject.objects.get(pk=subject1_ref_id).name
                except Subject.DoesNotExist:
                    pass
            if record.subject1_ref_id:
                try:
                    return Subject.objects.get(pk=record.subject1_ref_id).name
                except Subject.DoesNotExist:
                    pass
            return record.subject1
        else:
            if subject2_ref_id:
                try:
                    return Subject.objects.get(pk=subject2_ref_id).name
                except Subject.DoesNotExist:
                    pass
            if record.subject2_ref_id:
                try:
                    return Subject.objects.get(pk=record.subject2_ref_id).name
                except Subject.DoesNotExist:
                    pass
            return record.subject2
    
    # Build result structure
    months_result = {}
    changed_months = []
    warnings = []
    
    # For each month, build response with reason
    for month_num in range(1, 13):
        field_key = f'month_{month_num}'
        current_value = original_values[field_key]
        
        # Check if this month is affected (recalculated)
        if month_num not in affected_months:
            # Keep legacy value
            months_result[str(month_num)] = {
                'value': str(current_value) if current_value else '0.00',
                'source': 'legacy',
                'reason': 'Legacy-Wert, nicht neu berechnet',
            }
            continue
        
        # Month is being calculated
        semester = 1 if month_num <= 6 else 2
        subj_name = get_subject_name(semester)
        is_price_legacy = price1_is_legacy if semester == 1 else price2_is_legacy
        start_m = start_month_1 if semester == 1 else start_month_2
        
        base = base_amounts.get(field_key, ZERO)
        final = final_values.get(field_key, ZERO)
        
        # Build reason string
        reason_parts = []
        
        # Check if month is zeroed due to contract termination
        if month_num in calc_flags.terminated_months:
            termination_month = calc_flags.termination_from_month
            reason = f"Vertrag gekündigt ab Monat {termination_month}"
        elif month_num < start_m:
            reason = f"Vor Startmonat ({start_m})"
        elif not subj_name:
            reason = "Kein Fach ausgewählt"
        elif base == ZERO:
            if is_per_hour_subject(subj_name):
                reason = f"Stundenfach ({subj_name}), keine Stunden eingegeben"
            else:
                reason = "Kein Preis ausgewählt"
        else:
            # Build base description
            if is_per_hour_subject(subj_name):
                hours_val = merged_hours.get(field_key)
                hours_decimal = parse_decimal_or_none(hours_val) or ZERO
                price_val = None
                if hours_decimal > 0:
                    price_val = base / hours_decimal
                    reason_parts.append(f"Basis: {hours_decimal} Std × {price_val:.2f} € = {base:.2f} €")
                else:
                    reason_parts.append(f"Basis: {base:.2f} €")
            else:
                reason_parts.append(f"Basis: {base:.2f} €")
            
            # Add legacy price note
            if is_price_legacy:
                reason_parts.append("(Legacy-Preis)")
            
            # Check discount application
            # Use the preview values for discounts_disabled
            is_nachhilfe = is_nachhilfe_subject(subj_name)
            
            # Check if discounts are disabled for this specific month
            discounts_skipped_for_month = False
            if discounts_disabled:
                if not discounts_disabled_months:
                    # Empty list = all months disabled
                    discounts_skipped_for_month = True
                    reason_parts.append("Rabatte deaktiviert (alle Monate)")
                elif month_num in discounts_disabled_months:
                    # This month is in the disabled list
                    discounts_skipped_for_month = True
                    reason_parts.append(f"Rabatte deaktiviert für Monat {month_num}")
            
            if not discounts_skipped_for_month and is_nachhilfe:
                discounts_skipped_for_month = True
                reason_parts.append("Nachhilfe: keine Rabatte")
            
            if not discounts_skipped_for_month and base > ZERO:
                # Get discount breakdown for this month
                percent_sum, fixed_sum = collect_discounts_for_month(
                    month_num, family_discounts, record_discounts
                )
                
                if percent_sum > ZERO or fixed_sum > ZERO:
                    discount_parts = []
                    if percent_sum > ZERO:
                        # Clamp percent as in calculate_month_values
                        clamped_percent = min(percent_sum, MAX_PERCENT_DISCOUNT)
                        percent_display = clamped_percent * 100
                        discount_parts.append(f"{percent_display:.0f}% Rabatt")
                    if fixed_sum > ZERO:
                        discount_parts.append(f"{fixed_sum:.2f} € Festrabatt")
                    
                    reason_parts.append(f"Rabatte: {', '.join(discount_parts)}")
                    reason_parts.append(f"Ergebnis: {final:.2f} €")
                else:
                    reason_parts.append("Keine Rabatte anwendbar")
            
            reason = "; ".join(reason_parts)
        
        # Check if value changed from original
        if current_value is not None:
            try:
                current_dec = Decimal(str(current_value))
                if current_dec != final:
                    changed_months.append(month_num)
            except:
                changed_months.append(month_num)
        else:
            if final != ZERO:
                changed_months.append(month_num)
        
        months_result[str(month_num)] = {
            'value': str(final),
            'source': 'calculated',
            'reason': reason,
        }
    
    # Add warnings
    # Only show warning in LEGACY mode if there are no meaningful changes detected
    # (i.e., no months would be recalculated on save)
    if is_legacy_mode and not affected_months:
        warnings.append("Legacy-Modus: Keine abrechnungsrelevanten Änderungen erkannt. Wählen Sie Fach/Preis oder benutzen Sie 'Monate neu berechnen'.")
    
    # Add clamp warnings from calculation flags
    if calc_flags.clamped_to_zero_months:
        months_str = ", ".join(str(m) for m in sorted(calc_flags.clamped_to_zero_months))
        warnings.append(f"Monate {months_str}: Rabatt übersteigt Basis, auf 0 € gesetzt.")
    
    if calc_flags.percent_discount_exceeded:
        warnings.append(f"Prozentrabatt auf 99% begrenzt (Original: {calc_flags.original_percent_sum * 100:.0f}%).")
    
    return JsonResponse({
        'months': months_result,
        'changed_months': sorted(changed_months),
        'warnings': warnings,
        'record_months_mode': record.months_mode,
    })


# =============================================================================
# Family Search API (Superadmin only - for Kosten Report)
# =============================================================================

# Limit for family search results
FAMILY_SEARCH_LIMIT = 20

# Characters used to split query into tokens
import re
_TOKEN_SPLIT_PATTERN = re.compile(r'[\s,;.]+')


def _check_superadmin(user) -> JsonResponse | None:
    """Check if user is superadmin. Returns JsonResponse error or None if ok."""
    if not user.is_superadmin:
        return JsonResponse({
            'error': 'Access denied',
            'code': 'ACCESS_DENIED',
        }, status=403)
    return None


@login_required
@require_GET
def family_search_api(request: HttpRequest) -> JsonResponse:
    """
    API endpoint for searching families by parent name.
    
    Superadmin-only endpoint for the Kosten-Report feature.
    Searches across ALL years and returns unique family_id results.
    
    GET /api/karteien/family-search/?q=робинович
    
    Query tokenization:
    - The query is split by whitespace and punctuation (spaces, commas, semicolons, dots)
    - Each token must match (AND logic) via case-insensitive icontains
    - Example: "Nikolay,Robinovich" is found by searching "robinovich"
    
    Returns JSON with:
    - query: the original search query
    - results: array of unique families:
        - family_id: the unique family identifier
        - parent_name: from the latest year record (preferring NORMAL over PENDING)
        - latest_year: the most recent year with records for this family
        - has_pending: true if any PENDING records exist for this family
    """
    user = request.user
    
    # Access control - superadmin only
    error = _check_superadmin(user)
    if error:
        return error
    
    # Get query parameter
    query = request.GET.get('q', '').strip()
    
    # Empty query - return empty results
    if not query:
        return JsonResponse({
            'query': query,
            'results': [],
        })
    
    # Split query into tokens (by spaces and punctuation)
    tokens = [t.strip() for t in _TOKEN_SPLIT_PATTERN.split(query) if t.strip()]
    
    # If no valid tokens, return empty
    if not tokens:
        return JsonResponse({
            'query': query,
            'results': [],
        })
    
    # Build filter: all tokens must match parent_name (AND logic)
    filter_q = Q()
    for token in tokens:
        filter_q &= Q(parent_name__icontains=token)
    
    # Query to get unique families with latest year record
    # Filter criteria:
    #   - Only NORMAL or PENDING status (exclude DECLINED for Kosten-Report)
    #   - Exclude empty/NULL family_id (unusable for reports)
    # Order by: family_id ASC, year DESC, status ASC ('' < 'PENDING')
    # This ensures NORMAL (status='') comes before PENDING for the same year
    # distinct("family_id") picks the first row per family_id
    qs = (
        KarteiRecord.objects
        .filter(filter_q)
        .filter(status__in=[RecordStatus.NORMAL, RecordStatus.PENDING])
        .exclude(family_id__isnull=True)
        .exclude(family_id='')
        .order_by('family_id', '-year', 'status')
        .distinct('family_id')
        [:FAMILY_SEARCH_LIMIT]
    )
    
    # Collect results
    results = []
    family_ids = []
    for record in qs:
        family_ids.append(record.family_id)
        results.append({
            'family_id': record.family_id,
            'parent_name': record.parent_name,
            'latest_year': record.year,
            'has_pending': False,  # Will be updated below
        })
    
    # If we have results, check which families have PENDING records
    if family_ids:
        pending_families = set(
            KarteiRecord.objects
            .filter(
                family_id__in=family_ids,
                status=RecordStatus.PENDING,
            )
            .values_list('family_id', flat=True)
            .distinct()
        )
        
        # Update has_pending flags
        for result in results:
            if result['family_id'] in pending_families:
                result['has_pending'] = True
    
    return JsonResponse({
        'query': query,
        'results': results,
    })
