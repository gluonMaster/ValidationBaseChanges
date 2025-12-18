from django.contrib import admin

from .models import (
    Teacher, Subject, TeachingAssignment, PriceOption,
    Discount, FamilyDiscount, RecordDiscount
)


@admin.register(Teacher)
class TeacherAdmin(admin.ModelAdmin):
    """Admin interface for Teacher model."""
    list_display = ['last_name', 'first_name', 'is_active']
    list_filter = ['is_active']
    search_fields = ['last_name', 'first_name']
    ordering = ['last_name', 'first_name']


@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    """Admin interface for Subject model."""
    list_display = ['name', 'is_active']
    list_filter = ['is_active']
    search_fields = ['name']
    ordering = ['name']


@admin.register(TeachingAssignment)
class TeachingAssignmentAdmin(admin.ModelAdmin):
    """Admin interface for TeachingAssignment model."""
    list_display = ['year', 'subject', 'teacher', 'is_active']
    list_filter = ['year', 'is_active', 'subject', 'teacher']
    search_fields = ['subject__name', 'teacher__last_name', 'teacher__first_name']
    ordering = ['-year', 'subject__name', 'teacher__last_name']
    autocomplete_fields = ['subject', 'teacher']


@admin.register(PriceOption)
class PriceOptionAdmin(admin.ModelAdmin):
    """Admin interface for PriceOption model."""
    list_display = ['year', 'subject', 'amount', 'price_unit_display', 'is_active']
    list_filter = ['year', 'subject', 'is_active']
    search_fields = ['subject__name', 'comment']
    ordering = ['-year', 'subject__name', 'amount']
    autocomplete_fields = ['subject']
    list_per_page = 50

    @admin.display(description='Einheit')
    def price_unit_display(self, obj):
        """Display the price unit (€/Monat or €/UE) based on subject."""
        return obj.get_price_unit()


# =============================================================================
# Discount Admin
# =============================================================================

@admin.register(Discount)
class DiscountAdmin(admin.ModelAdmin):
    """Admin interface for Discount model."""
    list_display = ['id', 'kind', 'value_display', 'description_short', 'is_active']
    list_filter = ['kind', 'is_active']
    search_fields = ['description']
    ordering = ['kind', '-value']
    list_per_page = 50

    @admin.display(description='Wert')
    def value_display(self, obj):
        """Display value with proper formatting."""
        if obj.kind == 'PERCENT':
            return f"{obj.value * 100:.0f}%"
        return f"{obj.value:.2f} €"

    @admin.display(description='Beschreibung')
    def description_short(self, obj):
        """Truncate description for list view."""
        if len(obj.description) > 50:
            return obj.description[:50] + '...'
        return obj.description


@admin.register(FamilyDiscount)
class FamilyDiscountAdmin(admin.ModelAdmin):
    """Admin interface for FamilyDiscount model."""
    list_display = ['year', 'family_id', 'discount', 'months_display', 'created_at']
    list_filter = ['year', 'discount__kind']
    search_fields = ['family_id', 'discount__description']
    ordering = ['-year', 'family_id', 'start_month']
    autocomplete_fields = ['discount']
    list_per_page = 50
    readonly_fields = ['created_at', 'updated_at']

    fieldsets = (
        (None, {
            'fields': ('year', 'family_id', 'discount')
        }),
        ('Zeitraum', {
            'fields': ('start_month', 'end_month', 'months'),
            'description': 'Wenn "Monate (Liste)" ausgefüllt ist, wird der Bereich "Von-Bis" ignoriert.'
        }),
        ('Metadaten', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )

    @admin.display(description='Monate')
    def months_display(self, obj):
        return obj.months_display()


@admin.register(RecordDiscount)
class RecordDiscountAdmin(admin.ModelAdmin):
    """Admin interface for RecordDiscount model."""
    list_display = ['record_pkid', 'record_year', 'record_family', 'discount', 'months_display', 'created_at']
    list_filter = ['record__year', 'discount__kind']
    search_fields = ['record__pkid', 'record__family_id', 'discount__description']
    ordering = ['-record__year', 'record__pkid', 'start_month']
    autocomplete_fields = ['discount']
    raw_id_fields = ['record']
    list_per_page = 50
    readonly_fields = ['created_at', 'updated_at']

    fieldsets = (
        (None, {
            'fields': ('record', 'discount')
        }),
        ('Zeitraum', {
            'fields': ('start_month', 'end_month', 'months'),
            'description': 'Wenn "Monate (Liste)" ausgefüllt ist, wird der Bereich "Von-Bis" ignoriert.'
        }),
        ('Metadaten', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )

    @admin.display(description='PKID', ordering='record__pkid')
    def record_pkid(self, obj):
        return obj.record.pkid

    @admin.display(description='Jahr', ordering='record__year')
    def record_year(self, obj):
        return obj.record.year

    @admin.display(description='FamilyID', ordering='record__family_id')
    def record_family(self, obj):
        return obj.record.family_id

    @admin.display(description='Monate')
    def months_display(self, obj):
        return obj.months_display()
