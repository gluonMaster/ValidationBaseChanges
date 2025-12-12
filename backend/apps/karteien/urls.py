"""
URL configuration for the karteien app.

Admin/Operator UI for managing Kartei records:
- List view with filters
- Detail view
- Create/Edit/Delete views
"""

from django.urls import path

from . import views

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
]
