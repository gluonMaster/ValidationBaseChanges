from django.contrib import admin

from .models import (
    Teacher, Subject, TeachingAssignment, PriceOption,
    Discount, FamilyDiscount, RecordDiscount,
    SemesterConfig,
    SubjectCategory, SubjectCategoryLink,
    DisciplineGroup, DurationEntry, GroupSizeEntry,
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


# =============================================================================
# Semester Configuration Admin
# =============================================================================

@admin.register(SemesterConfig)
class SemesterConfigAdmin(admin.ModelAdmin):
    """Admin interface for SemesterConfig model.

    For years that already have KarteiRecords the boundary field is shown
    as read-only and the row is highlighted.
    """
    list_display = ['year', 'last_month_sem1', 'has_records_display', 'updated_at']
    list_filter = ['last_month_sem1']
    ordering = ['-year']
    readonly_fields = ['created_at', 'updated_at']

    fieldsets = (
        (None, {
            'fields': ('year', 'last_month_sem1'),
        }),
        ('Metadaten', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',),
        }),
    )

    @admin.display(description='Einträge vorhanden', boolean=True)
    def has_records_display(self, obj):
        """Show whether KarteiRecords exist for this year."""
        from apps.karteien.models import KarteiRecord
        return KarteiRecord.objects.filter(year=obj.year).exists()

    def get_readonly_fields(self, request, obj=None):
        """Make last_month_sem1 read-only when the year already has records."""
        readonly = list(super().get_readonly_fields(request, obj))
        if obj and obj.pk:
            from apps.karteien.models import KarteiRecord
            if KarteiRecord.objects.filter(year=obj.year).exists():
                readonly.append('last_month_sem1')
                readonly.append('year')
        return readonly


# =============================================================================
# Subject Category Admin
# =============================================================================

class SubjectCategoryLinkInline(admin.TabularInline):
    """Inline for linking subjects to a category."""
    model = SubjectCategoryLink
    extra = 1
    autocomplete_fields = ['subject']
    readonly_fields = ['year']


@admin.register(SubjectCategory)
class SubjectCategoryAdmin(admin.ModelAdmin):
    """Admin interface for SubjectCategory model."""
    list_display = ['year', 'name', 'kind', 'yearly_rate', 'monthly_rate',
                    'group_threshold', 'is_active']
    list_filter = ['year', 'kind', 'is_active']
    search_fields = ['name']
    ordering = ['-year', 'name']
    list_per_page = 50
    inlines = [SubjectCategoryLinkInline]


@admin.register(SubjectCategoryLink)
class SubjectCategoryLinkAdmin(admin.ModelAdmin):
    """Admin interface for SubjectCategoryLink (debugging)."""
    list_display = ['year', 'subject', 'category']
    list_filter = ['year', 'category__kind']
    search_fields = ['subject__name', 'category__name']
    autocomplete_fields = ['subject', 'category']
    ordering = ['-year', 'subject__name']
    readonly_fields = ['year']


# =============================================================================
# Discipline Group Admin
# =============================================================================

class DurationEntryInline(admin.TabularInline):
    """Inline for duration entries within a DisciplineGroup."""
    model = DurationEntry
    extra = 1
    readonly_fields = ['changed_at']


class GroupSizeEntryInline(admin.TabularInline):
    """Inline for manual group-size overrides within a DisciplineGroup."""
    model = GroupSizeEntry
    extra = 1
    readonly_fields = ['changed_at']


@admin.register(DisciplineGroup)
class DisciplineGroupAdmin(admin.ModelAdmin):
    """Admin interface for DisciplineGroup model."""
    list_display = ['year', 'subject', 'category', 'auto_scaling_enabled', 'is_active']
    list_filter = ['year', 'category', 'is_active', 'auto_scaling_enabled']
    search_fields = ['subject__name', 'category__name']
    ordering = ['-year', 'subject__name']
    autocomplete_fields = ['subject', 'category']
    list_per_page = 50
    inlines = [DurationEntryInline, GroupSizeEntryInline]
