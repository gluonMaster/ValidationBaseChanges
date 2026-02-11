"""
URL configuration for catalog app.

Provides routes for managing Teachers, Subjects, TeachingAssignments, PriceOptions, and Discounts.
Only accessible to users with ADMIN role.
"""

from django.urls import path

from . import views

app_name = "catalog"

urlpatterns = [
    # Index page
    path("", views.CatalogIndexView.as_view(), name="index"),
    
    # Teachers CRUD
    path("teachers/", views.TeacherListView.as_view(), name="teacher_list"),
    path("teachers/create/", views.TeacherCreateView.as_view(), name="teacher_create"),
    path("teachers/<int:pk>/edit/", views.TeacherUpdateView.as_view(), name="teacher_edit"),
    
    # Subjects CRUD
    path("subjects/", views.SubjectListView.as_view(), name="subject_list"),
    path("subjects/create/", views.SubjectCreateView.as_view(), name="subject_create"),
    path("subjects/<int:pk>/edit/", views.SubjectUpdateView.as_view(), name="subject_edit"),
    
    # Teaching Assignments CRUD
    path("assignments/", views.AssignmentListView.as_view(), name="assignment_list"),
    path("assignments/create/", views.AssignmentCreateView.as_view(), name="assignment_create"),
    path("assignments/<int:pk>/delete/", views.AssignmentDeleteView.as_view(), name="assignment_delete"),
    
    # Copy year functionality (assignments)
    path("assignments/copy-year/", views.CopyYearView.as_view(), name="copy_year"),
    
    # Price Options CRUD
    path("prices/", views.PriceListView.as_view(), name="price_list"),
    path("prices/create/", views.PriceCreateView.as_view(), name="price_create"),
    path("prices/<int:pk>/edit/", views.PriceUpdateView.as_view(), name="price_edit"),
    path("prices/<int:pk>/delete/", views.PriceDeleteView.as_view(), name="price_delete"),
    
    # Copy year functionality (prices)
    path("prices/copy-year/", views.CopyPricesYearView.as_view(), name="copy_prices_year"),
    
    # Discounts CRUD
    path("discounts/", views.DiscountListView.as_view(), name="discount_list"),
    path("discounts/create/", views.DiscountCreateView.as_view(), name="discount_create"),
    path("discounts/<int:pk>/edit/", views.DiscountUpdateView.as_view(), name="discount_edit"),
    
    # Family Discounts CRUD
    path("family-discounts/", views.FamilyDiscountListView.as_view(), name="family_discount_list"),
    path("family-discounts/create/", views.FamilyDiscountCreateView.as_view(), name="family_discount_create"),
    path("family-discounts/<int:pk>/edit/", views.FamilyDiscountUpdateView.as_view(), name="family_discount_edit"),
    path("family-discounts/<int:pk>/delete/", views.FamilyDiscountDeleteView.as_view(), name="family_discount_delete"),
    
    # Record Discounts CRUD
    path("record-discounts/", views.RecordDiscountListView.as_view(), name="record_discount_list"),
    path("record-discounts/create/", views.RecordDiscountCreateView.as_view(), name="record_discount_create"),
    path("record-discounts/<int:pk>/edit/", views.RecordDiscountUpdateView.as_view(), name="record_discount_edit"),
    path("record-discounts/<int:pk>/delete/", views.RecordDiscountDeleteView.as_view(), name="record_discount_delete"),
    
    # FamilyID Reservations (Admin only)
    path("familyid-reservations/", views.FamilyIdReservationListView.as_view(), name="familyid_reservation_list"),
    path("familyid-reservations/reserve-next/", views.ReserveNextFamilyIdView.as_view(), name="familyid_reserve_next"),
    path("familyid-reservations/<int:pk>/cancel/", views.CancelFamilyIdReservationView.as_view(), name="familyid_reservation_cancel"),
    
    # Sync from Legacy tool (Admin only)
    path("sync-from-legacy/", views.SyncFromLegacyView.as_view(), name="sync_from_legacy"),
    
    # Subject Categories CRUD (per year)
    path("categories/<int:year>/", views.SubjectCategoryListView.as_view(), name="category_list"),
    path("categories/<int:year>/create/", views.SubjectCategoryCreateView.as_view(), name="category_create"),
    path("categories/<int:year>/<int:pk>/edit/", views.SubjectCategoryUpdateView.as_view(), name="category_edit"),
    path("categories/<int:year>/<int:pk>/delete/", views.SubjectCategoryDeleteView.as_view(), name="category_delete"),
    
    # Subject ↔ Category link management
    path("categories/<int:year>/<int:pk>/subjects/", views.SubjectCategoryLinksView.as_view(), name="category_subjects"),
    path("categories/<int:year>/<int:pk>/subjects/<int:link_pk>/remove/", views.SubjectCategoryUnlinkView.as_view(), name="category_unlink_subject"),
    
    # DisciplineGroup management
    path("groups/<int:year>/", views.DisciplineGroupListView.as_view(), name="group_list"),
    path("groups/<int:year>/<int:pk>/", views.DisciplineGroupDetailView.as_view(), name="group_detail"),
    path("groups/<int:year>/<int:pk>/duration/add/", views.DurationEntryCreateView.as_view(), name="group_duration_add"),
    path("groups/<int:year>/<int:pk>/size/add/", views.GroupSizeEntryCreateView.as_view(), name="group_size_add"),
    path("groups/<int:year>/<int:pk>/toggle-scaling/", views.DisciplineGroupToggleScalingView.as_view(), name="group_toggle_scaling"),

    # API endpoints
    path("api/groups/<int:year>/<int:pk>/size/", views.GroupSizeApiView.as_view(), name="group_size_api"),
]
