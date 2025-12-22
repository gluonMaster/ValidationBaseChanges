"""
API views for the karteien app.

This module provides JSON API endpoints for:
- Month breakdown details for billing explanation
- Live search for record list filtering
- Autocomplete for parent/child names
- Prefill form data from existing records
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

from .billing import get_month_breakdown
from .models import KarteiRecord, RecordStatus


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
