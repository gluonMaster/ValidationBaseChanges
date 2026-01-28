"""
Simple JSON API views for catalog data.

These endpoints provide dynamic data for the KarteiRecord form:
- Subjects (all active subjects)
- Teachers filtered by year and subject
- Prices filtered by year and subject

These are simple Django views (not DRF) for minimal overhead.
"""

from __future__ import annotations

import json

from django.http import JsonResponse
from django.views import View
from django.contrib.auth.mixins import LoginRequiredMixin

from .models import Subject, Teacher, PriceOption, TeachingAssignment


class SubjectsApiView(LoginRequiredMixin, View):
    """
    JSON API endpoint for all active subjects.
    
    GET /api/catalog/subjects/
    
    Returns:
        {
            "subjects": [
                {"id": 1, "name": "11A Klasse Di"},
                {"id": 2, "name": "Mathematik"},
                ...
            ]
        }
    """
    
    def get(self, request):
        subjects = Subject.objects.filter(
            is_active=True
        ).order_by("name")
        
        subject_list = [
            {
                "id": s.id,
                "name": s.name,
            }
            for s in subjects
        ]
        
        return JsonResponse({"subjects": subject_list})


class TeachersApiView(LoginRequiredMixin, View):
    """
    JSON API endpoint for teachers filtered by year and subject.
    
    GET /api/catalog/teachers/?year=2025&subject_id=123
    
    Returns:
        {
            "teachers": [
                {"id": 1, "name": "Müller, Hans"},
                {"id": 2, "name": "Schmidt, Anna"},
                ...
            ]
        }
    """
    
    def get(self, request):
        year = request.GET.get("year")
        subject_id = request.GET.get("subject_id")
        
        if not year:
            return JsonResponse({"error": "year parameter is required"}, status=400)
        
        try:
            year = int(year)
        except ValueError:
            return JsonResponse({"error": "year must be an integer"}, status=400)
        
        # Get teacher IDs from active TeachingAssignments for this year and subject
        assignment_filter = {
            "year": year,
            "is_active": True,
        }
        
        if subject_id:
            try:
                assignment_filter["subject_id"] = int(subject_id)
            except ValueError:
                return JsonResponse({"error": "subject_id must be an integer"}, status=400)
        
        teacher_ids = TeachingAssignment.objects.filter(
            **assignment_filter
        ).values_list("teacher_id", flat=True).distinct()
        
        # Get active teachers with these IDs
        teachers = Teacher.objects.filter(
            id__in=teacher_ids,
            is_active=True
        ).order_by("last_name", "first_name")
        
        teacher_list = [
            {
                "id": t.id,
                "name": str(t),  # "LastName, FirstName"
            }
            for t in teachers
        ]
        
        return JsonResponse({"teachers": teacher_list})


class PricesApiView(LoginRequiredMixin, View):
    """
    JSON API endpoint for prices filtered by year and subject.
    
    GET /api/catalog/prices/?year=2025&subject_id=123
    
    Returns:
        {
            "prices": [
                {"id": 1, "amount": "45.00", "unit": "€/Monat", "comment": "Standard"},
                {"id": 2, "amount": "60.00", "unit": "€/UE", "comment": "Individuell"},
                ...
            ]
        }
    """
    
    def get(self, request):
        year = request.GET.get("year")
        subject_id = request.GET.get("subject_id")
        
        if not year:
            return JsonResponse({"error": "year parameter is required"}, status=400)
        
        try:
            year = int(year)
        except ValueError:
            return JsonResponse({"error": "year must be an integer"}, status=400)
        
        # Build filter
        price_filter = {
            "year": year,
            "is_active": True,
        }
        
        if subject_id:
            try:
                price_filter["subject_id"] = int(subject_id)
            except ValueError:
                return JsonResponse({"error": "subject_id must be an integer"}, status=400)
        
        prices = PriceOption.objects.filter(
            **price_filter
        ).select_related("subject").order_by("amount")
        
        price_list = [
            {
                "id": p.id,
                "amount": str(p.amount),
                "unit": p.get_price_unit(),
                "comment": p.comment,
                "subject_name": p.subject.name if p.subject else "",
            }
            for p in prices
        ]
        
        return JsonResponse({"prices": price_list})
