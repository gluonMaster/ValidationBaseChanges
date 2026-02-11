from django.contrib import admin

from .models import (
    ContractStatusEntry,
    ContractTypeEntry,
    KarteiRecord,
)


# =============================================================================
# Inline admins for KarteiRecord
# =============================================================================

class ContractTypeEntryInline(admin.TabularInline):
    model = ContractTypeEntry
    extra = 0
    fields = ("effective_from_month", "is_monthly", "changed_by", "changed_at", "comment")
    readonly_fields = ("changed_at",)


class ContractStatusEntryInline(admin.TabularInline):
    model = ContractStatusEntry
    extra = 0
    fields = ("effective_from_month", "kind", "changed_by", "changed_at", "comment")
    readonly_fields = ("changed_at",)


# =============================================================================
# KarteiRecord Admin
# =============================================================================

@admin.register(KarteiRecord)
class KarteiRecordAdmin(admin.ModelAdmin):
    list_display = ("family_id", "parent_name", "child_name", "year", "status")
    list_filter = ("year", "status")
    search_fields = ("family_id", "parent_name", "child_name")
    inlines = [ContractTypeEntryInline, ContractStatusEntryInline]


# =============================================================================
# Standalone admins (useful for debugging / bulk inspection)
# =============================================================================

@admin.register(ContractTypeEntry)
class ContractTypeEntryAdmin(admin.ModelAdmin):
    list_display = ("record", "effective_from_month", "is_monthly", "changed_by", "changed_at")
    list_filter = ("is_monthly", "effective_from_month")
    raw_id_fields = ("record",)
    readonly_fields = ("changed_at",)


@admin.register(ContractStatusEntry)
class ContractStatusEntryAdmin(admin.ModelAdmin):
    list_display = ("record", "effective_from_month", "kind", "changed_by", "changed_at")
    list_filter = ("kind", "effective_from_month")
    raw_id_fields = ("record",)
    readonly_fields = ("changed_at",)
