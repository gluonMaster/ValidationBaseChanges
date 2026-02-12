"""
URL configuration for the karteien app.

Admin/Operator UI for managing Kartei records:
- List view with filters
- Detail view
- Create/Edit/Delete views
- Emergency months override (admin-only)
- Family Dashboard (admin-only)
- FamilyDiscount CRUD (admin-only)
- New Family Wizard (admin-only)
- API: Month breakdown
- API: Live search
- API: Autocomplete parents/children
- API: Prefill from record
"""

from django.urls import path

from . import views
from . import families
from . import api
from . import wizard
from . import kosten_report

app_name = 'karteien'

urlpatterns = [
    # List view
    path('', views.KarteiRecordListView.as_view(), name='record_list'),
    
    # Detail view
    path('<int:pk>/', views.KarteiRecordDetailView.as_view(), name='record_detail'),
    
    # Create view
    path('create/', views.KarteiRecordCreateView.as_view(), name='record_create'),
    
    # Update view
    path('<int:pk>/edit/', views.KarteiRecordUpdateView.as_view(), name='record_update'),
    
    # Delete view
    path('<int:pk>/delete/', views.KarteiRecordDeleteView.as_view(), name='record_delete'),
    
    # Emergency months override (admin-only)
    path('<int:pk>/months-override/', views.MonthsOverrideView.as_view(), name='months_override'),
    
    # Apply category price (admin/operator)
    path('<int:pk>/apply-price/preview/', views.apply_price_preview, name='record_apply_price_preview'),
    path('<int:pk>/apply-price/', views.ApplyCategoryPriceView.as_view(), name='record_apply_price'),

    # Contract type/status change (creates PendingChange)
    path('<int:pk>/contract-type/change/', views.ContractTypeChangeView.as_view(), name='record_contract_type_change'),
    path('<int:pk>/contract-status/change/', views.ContractStatusChangeView.as_view(), name='record_contract_status_change'),

    # Quick-set subject*_ref from legacy text match (creates PendingChange)
    path('<int:pk>/quick-set-subject-ref/', views.QuickSetSubjectRefView.as_view(), name='record_quick_set_subject_ref'),
    
    # API: Live search (for AJAX filtering)
    path('live-search/', api.live_search_api, name='live_search_api'),
    
    # API: Month breakdown
    path('<int:pk>/month-breakdown/', api.month_breakdown_api, name='month_breakdown_api'),
    
    # API: Billing preview (live preview in edit form)
    path('<int:pk>/billing-preview/', api.billing_preview_api, name='billing_preview_api'),
    
    # API: Autocomplete
    path('autocomplete/parents/', api.autocomplete_parents_api, name='autocomplete_parents'),
    path('autocomplete/children/', api.autocomplete_children_api, name='autocomplete_children'),
    
    # API: Prefill from record
    path('prefill-from-record/', api.prefill_from_record_api, name='prefill_from_record'),
    
    # API: Subject dependents (teachers/prices) for wizard filtering
    path('subject-dependents/', api.subject_dependents_api, name='subject_dependents'),
    
    # Family Dashboard (admin-only)
    path('family/', families.FamilyDashboardView.as_view(), name='family_dashboard'),
    
    # New Family Wizard (admin-only)
    path('family/new/', wizard.NewFamilyWizardView.as_view(), name='new_family_wizard'),
    
    # FamilyDiscount CRUD (admin-only)
    path('family/discounts/create/', families.FamilyDiscountCreateView.as_view(), name='family_discount_create'),
    path('family/discounts/<int:pk>/edit/', families.FamilyDiscountEditView.as_view(), name='family_discount_edit'),
    path('family/discounts/<int:pk>/delete/', families.FamilyDiscountDeleteView.as_view(), name='family_discount_delete'),
    
    # Apply discounts to family records (admin-only)
    path('family/apply-discounts/', families.ApplyDiscountsView.as_view(), name='apply_discounts'),
    
    # API: Family search for Kosten report (superadmin-only)
    path('family-search/', api.family_search_api, name='family_search_api'),
    
    # Family Kosten report (superadmin-only)
    path('family-kosten/', kosten_report.FamilyKostenReportView.as_view(), name='family_kosten_report'),
    
    # API: Family Kosten report fragment for Offcanvas (superadmin-only)
    path('family-kosten-fragment/', kosten_report.FamilyKostenFragmentView.as_view(), name='family_kosten_fragment'),
]
