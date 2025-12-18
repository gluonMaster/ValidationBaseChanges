"""
URL configuration for the approvals app.

Admin UI for managing declined records:
- Declined overview (list of DECLINED records)
- Apply fixes (move from DECLINED to PENDING)
- View pending changes

Superadmin UI for approving/declining pending changes:
- Pending overview (list of PENDING records for Superadmin)
- War/Ist view (detail comparison for a single pending change)
- Bulk actions (approve all, decline all)
- NeuList (new records view)
- Record history

Mirrors functionality from:
- Export_DeclinedTools.bas
- valid_ApproveFlow.bas
- valid_GrossGeschichteDecision.bas
- valid_NeuList.bas
"""

from django.urls import path

from . import views

app_name = 'approvals'

urlpatterns = [
    # ===========================================================================
    # Admin: Declined Overview
    # ===========================================================================
    path('declined/', views.DeclinedOverviewView.as_view(), name='declined_overview'),
    path('declined/<int:pk>/', views.DeclinedDetailView.as_view(), name='declined_detail'),
    path('declined/<int:pk>/edit/', views.DeclinedChangeEditView.as_view(), name='declined_edit'),
    
    # Apply fixes
    path('declined/<int:pk>/apply-fix/', views.ApplyDeclinedFixView.as_view(), name='apply_fix'),
    path('declined/apply-all/', views.ApplyAllDeclinedFixesView.as_view(), name='apply_all_fixes'),
    
    # Pending changes (info view for Admin)
    path('pending/', views.PendingChangesListView.as_view(), name='pending_list'),
    path('pending/<int:pk>/', views.PendingDetailView.as_view(), name='pending_detail'),
    
    # ===========================================================================
    # Superadmin: Pending Overview and Decisions
    # ===========================================================================
    path(
        'superadmin/pending/',
        views.SuperadminPendingOverviewView.as_view(),
        name='superadmin_pending_overview'
    ),
    path(
        'superadmin/pending/<int:pk>/',
        views.SuperadminWarIstView.as_view(),
        name='superadmin_war_ist'
    ),
    
    # Bulk actions
    path(
        'superadmin/approve-all/',
        views.SuperadminApproveAllView.as_view(),
        name='superadmin_approve_all'
    ),
    path(
        'superadmin/decline-all/',
        views.SuperadminDeclineAllView.as_view(),
        name='superadmin_decline_all'
    ),
    
    # ===========================================================================
    # Superadmin: NeuList (New Records)
    # ===========================================================================
    path(
        'superadmin/neu/',
        views.SuperadminNeuListView.as_view(),
        name='superadmin_neu_list'
    ),
    path(
        'superadmin/mark-seen/',
        views.SuperadminMarkSeenView.as_view(),
        name='superadmin_mark_seen'
    ),
    
    # ===========================================================================
    # Superadmin: Record History
    # ===========================================================================
    path(
        'superadmin/history/<int:pk>/',
        views.SuperadminRecordHistoryView.as_view(),
        name='superadmin_record_history'
    ),
]
