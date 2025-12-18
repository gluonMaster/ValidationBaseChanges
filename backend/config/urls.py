"""
URL configuration for KindEltern Web project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.0/topics/http/urls/
"""

from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import path, include
from django.views.generic import RedirectView

from apps.accounts.views import role_based_redirect
from apps.catalog.views_api import TeachersApiView, PricesApiView

urlpatterns = [
    # Django admin
    path('admin/', admin.site.urls),
    
    # Authentication
    path('accounts/login/', auth_views.LoginView.as_view(template_name='accounts/login.html'), name='login'),
    path('accounts/logout/', auth_views.LogoutView.as_view(template_name='accounts/logged_out.html'), name='logout'),
    
    # Root redirect based on user role
    path('', role_based_redirect, name='home'),
    
    # Web UI - User cabinet (read-only access)
    path('user/', include('apps.accounts.urls')),
    
    # Web UI - Admin/Operator interfaces
    path('karteien/', include('apps.karteien.urls')),
    path('approvals/', include('apps.approvals.urls')),
    path('catalog/', include('apps.catalog.urls')),
    
    # API endpoints
    path('api/accounts/', include(('apps.accounts.urls', 'accounts_api'), namespace='accounts_api')),
    path('api/karteien/', include(('apps.karteien.urls', 'karteien_api'), namespace='karteien_api')),
    path('api/approvals/', include(('apps.approvals.urls', 'approvals_api'), namespace='approvals_api')),
    path('api/history/', include('apps.history.urls')),
    path('api/notifications/', include('apps.notifications.urls')),
    path('api/legacy-import/', include('apps.legacy_import.urls')),
    path('api/reporting/', include('apps.reporting.urls')),
    
    # Catalog API endpoints for dynamic form data
    path('api/catalog/teachers/', TeachersApiView.as_view(), name='api_catalog_teachers'),
    path('api/catalog/prices/', PricesApiView.as_view(), name='api_catalog_prices'),
]
