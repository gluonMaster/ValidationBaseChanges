"""
URL configuration for the accounts app.

User Cabinet UI for read-only access to records and history:
- Dashboard (home page for User role)
- Search records
- View record details (read-only)
- View record history (read-only)

All operations are strictly read-only for User role.
"""

from django.urls import path

from . import views

app_name = 'accounts'

urlpatterns = [
    # User Dashboard (home page for User role)
    path('', views.UserDashboardView.as_view(), name='user_dashboard'),
    
    # User Search
    path('search/', views.UserKarteiSearchView.as_view(), name='user_search'),
    
    # User Record Detail (read-only)
    path('record/<int:pk>/', views.UserKarteiDetailView.as_view(), name='user_record_detail'),
    
    # User Record History (read-only)
    path('record/<int:pk>/history/', views.UserRecordHistoryView.as_view(), name='user_record_history'),
]
