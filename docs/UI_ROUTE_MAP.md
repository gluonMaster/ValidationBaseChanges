# UI Route Map - KindEltern Web

Last updated: 2026-05-27. Created by `PROMPT_180_DOCS_QA_MATRIX_UI_ROUTE_MAP_PROMPT_TEMPLATE.md`.

This is a compact route/permission map for AI agents. It reflects current code, not the PRICELIST V2 target state.

## Role Summary

| Role | Current primary landing | Main permissions | Main restrictions |
| --- | --- | --- | --- |
| ADMIN | `/karteien/` | Kartei list/detail/create/edit/delete, catalog management, family tools, approvals admin/info flows. | PRICELIST V2 risky changes still require SUPERADMIN approval. |
| OPERATOR | `/karteien/` | Kartei list/detail/create/edit normal records; some pricing editor paths (`CatalogEditorMixin`). | Blocked from editing `PENDING`/`DECLINED` in standard editor; no delete; catalog navigation is incomplete because `PROMPT_162` is not applied. |
| SUPERADMIN | `/approvals/superadmin/pending/` | Pending review/approve/decline, NeuList, read-only Kartei, family Kosten report. | No Kartei edit/create/delete. |
| USER | `/user/` | User cabinet and read-only Kartei/search/history views. | No write access. |

## Root and Auth

| URL | Name | View | Template | Roles |
| --- | --- | --- | --- | --- |
| `/accounts/login/` | `login` | Django `LoginView` | `accounts/login.html` | Anonymous/authenticated redirect flow |
| `/accounts/logout/` | `logout` | Django `LogoutView` | `accounts/logged_out.html` | Authenticated |
| `/` | `home` | `apps.accounts.views.role_based_redirect` | n/a | Authenticated, role-based redirect |

## Accounts / User Cabinet

| URL | Name | View | Template | Roles |
| --- | --- | --- | --- | --- |
| `/user/` | `accounts:user_dashboard` | `UserDashboardView` | `accounts/user_dashboard.html` | USER only |
| `/user/search/` | `accounts:user_search` | `UserKarteiSearchView` | `accounts/user_search.html` | USER only |
| `/user/record/<pkid>/` | `accounts:user_record_detail` | `UserKarteiDetailView` | `accounts/user_record_detail.html` | USER only |
| `/user/record/<pkid>/history/` | `accounts:user_record_history` | `UserRecordHistoryView` | `accounts/user_record_history.html` | USER only |

`UserRoleMixin` redirects SUPERADMIN to approvals and ADMIN/OPERATOR to `/karteien/`.

## Kartei

| URL | Name | View | Template | Roles |
| --- | --- | --- | --- | --- |
| `/karteien/` | `karteien:record_list` | `KarteiRecordListView` | `karteien/record_list.html`, `_record_list_table.html`, `_record_list_pagination.html` | ADMIN, OPERATOR, SUPERADMIN, USER |
| `/karteien/<pkid>/` | `karteien:record_detail` | `KarteiRecordDetailView` | `karteien/record_detail.html` | ADMIN, OPERATOR, SUPERADMIN, USER |
| `/karteien/create/` | `karteien:record_create` | `KarteiRecordCreateView` | `karteien/record_form.html` | ADMIN, OPERATOR |
| `/karteien/<pkid>/edit/` | `karteien:record_update` | `KarteiRecordUpdateView` | `karteien/record_form.html` | ADMIN, OPERATOR for normal records; ADMIN only for `PENDING`/`DECLINED` |
| `/karteien/<pkid>/delete/` | `karteien:record_delete` | `KarteiRecordDeleteView` | `karteien/record_confirm_delete.html` | ADMIN only |
| `/karteien/<pkid>/months-override/` | `karteien:months_override` | `MonthsOverrideView` | `karteien/months_override.html` | ADMIN only |
| `/karteien/<pkid>/apply-price/preview/` | `karteien:record_apply_price_preview` | `apply_price_preview` | JSON/redirect flow | ADMIN, OPERATOR |
| `/karteien/<pkid>/apply-price/` | `karteien:record_apply_price` | `ApplyCategoryPriceView` | Redirect flow | ADMIN, OPERATOR |
| `/karteien/<pkid>/contract-type/change/` | `karteien:record_contract_type_change` | `ContractTypeChangeView` | Redirect flow | ADMIN, OPERATOR |
| `/karteien/<pkid>/contract-status/change/` | `karteien:record_contract_status_change` | `ContractStatusChangeView` | Redirect flow | ADMIN, OPERATOR |
| `/karteien/<pkid>/quick-set-subject-ref/` | `karteien:record_quick_set_subject_ref` | `QuickSetSubjectRefView` | Redirect flow | ADMIN, OPERATOR |

Permission implementation:

- `KarteiViewerMixin`: ADMIN, OPERATOR, SUPERADMIN, USER can view list/detail.
- `KarteiEditorMixin`: ADMIN, OPERATOR can enter edit/create paths.
- `KarteiRecordUpdateView.dispatch`: OPERATOR is blocked for `PENDING` and `DECLINED`; ADMIN can edit those via standard editor.
- Delete and months override override `test_func` to ADMIN only.

## Kartei APIs and Family Tools

| URL | Name | View | Output/template | Roles |
| --- | --- | --- | --- | --- |
| `/karteien/live-search/` | `karteien:live_search_api` | `live_search_api` | JSON | ADMIN, OPERATOR, SUPERADMIN, USER |
| `/karteien/<pkid>/month-breakdown/` | `karteien:month_breakdown_api` | `month_breakdown_api` | JSON | ADMIN, OPERATOR, SUPERADMIN, USER |
| `/karteien/<pkid>/billing-preview/` | `karteien:billing_preview_api` | `billing_preview_api` | JSON | `user.can_edit_kartei` (ADMIN/OPERATOR in current role model; code comments still say Admin only) |
| `/karteien/autocomplete/parents/` | `karteien:autocomplete_parents` | `autocomplete_parents_api` | JSON | ADMIN, OPERATOR |
| `/karteien/autocomplete/children/` | `karteien:autocomplete_children` | `autocomplete_children_api` | JSON | ADMIN, OPERATOR |
| `/karteien/prefill-from-record/` | `karteien:prefill_from_record` | `prefill_from_record_api` | JSON | ADMIN, OPERATOR |
| `/karteien/subject-dependents/` | `karteien:subject_dependents` | `subject_dependents_api` | JSON | ADMIN only |
| `/karteien/family/` | `karteien:family_dashboard` | `FamilyDashboardView` | `karteien/family_dashboard.html` | ADMIN only |
| `/karteien/family/new/` | `karteien:new_family_wizard` | `NewFamilyWizardView` | `karteien/new_family_wizard.html` | ADMIN only |
| `/karteien/family/discounts/create/` and edit/delete | family discount routes | `FamilyDiscount*View` | catalog/karteien discount templates | ADMIN only |
| `/karteien/family/apply-discounts/` | `karteien:apply_discounts` | `ApplyDiscountsView` | Redirect flow | ADMIN only |
| `/karteien/family-search/` | `karteien:family_search_api` | `family_search_api` | JSON | SUPERADMIN only |
| `/karteien/family-kosten/` | `karteien:family_kosten_report` | `FamilyKostenReportView` | `karteien/family_kosten_report.html` | SUPERADMIN only |
| `/karteien/family-kosten-fragment/` | `karteien:family_kosten_fragment` | `FamilyKostenFragmentView` | `karteien/_family_kosten_report_fragment.html` | SUPERADMIN only |

The same `apps.karteien.urls` is also included under `/api/karteien/` with namespace `karteien_api`; be explicit about namespace and prefix in tests.

## Approvals

| URL | Name | View | Template | Roles |
| --- | --- | --- | --- | --- |
| `/approvals/declined/` | `approvals:declined_overview` | `DeclinedOverviewView` | `approvals/declined_overview.html` | `AdminEditorMixin` currently uses `user.can_edit_kartei` |
| `/approvals/declined/<pk>/` | `approvals:declined_detail` | `DeclinedDetailView` | `approvals/declined_detail.html` | `AdminEditorMixin` |
| `/approvals/declined/<pk>/edit/` | `approvals:declined_edit` | `DeclinedChangeEditView` | `approvals/declined_edit.html` | `AdminEditorMixin` |
| `/approvals/declined/<pk>/apply-fix/` | `approvals:apply_fix` | `ApplyDeclinedFixView` | Redirect flow | `AdminEditorMixin` |
| `/approvals/declined/apply-all/` | `approvals:apply_all_fixes` | `ApplyAllDeclinedFixesView` | Redirect flow | `AdminEditorMixin` |
| `/approvals/pending/` | `approvals:pending_list` | `PendingChangesListView` | `approvals/pending_list.html` | `AdminEditorMixin` |
| `/approvals/pending/<pk>/` | `approvals:pending_detail` | `PendingDetailView` | `approvals/pending_detail.html` | `AdminEditorMixin` |
| `/approvals/pending/<pk>/edit/` | `approvals:pending_edit` | `PendingChangeEditView` | `approvals/pending_edit.html` | `AdminEditorMixin` |
| `/approvals/superadmin/pending/` | `approvals:superadmin_pending_overview` | `SuperadminPendingOverviewView` | `approvals/superadmin/pending_overview.html` | SUPERADMIN only |
| `/approvals/superadmin/pending/<pk>/` | `approvals:superadmin_war_ist` | `SuperadminWarIstView` | `approvals/superadmin/war_ist.html` | SUPERADMIN only |
| `/approvals/superadmin/approve-all/` | `approvals:superadmin_approve_all` | `SuperadminApproveAllView` | Redirect flow | SUPERADMIN only |
| `/approvals/superadmin/decline-all/` | `approvals:superadmin_decline_all` | `SuperadminDeclineAllView` | Redirect flow | SUPERADMIN only |
| `/approvals/superadmin/neu/` | `approvals:superadmin_neu_list` | `SuperadminNeuListView` | `approvals/superadmin/neu_list.html` | SUPERADMIN only |
| `/approvals/superadmin/mark-seen/` | `approvals:superadmin_mark_seen` | `SuperadminMarkSeenView` | Redirect flow | SUPERADMIN only |
| `/approvals/superadmin/history/<pkid>/` | `approvals:superadmin_record_history` | `SuperadminRecordHistoryView` | `approvals/superadmin/record_history.html` | SUPERADMIN only |

Route risk: `AdminEditorMixin` documentation says Admin, but `test_func` uses `user.can_edit_kartei`, which likely includes OPERATOR. Treat this as a permission audit item before adding UI links or tests that assume Admin-only approvals admin routes.

## Catalog

Most catalog routes use `CatalogAdminMixin` (ADMIN only). Pricing workflow exceptions use `CatalogEditorMixin` (ADMIN and OPERATOR).

| URL group | Representative names | Views/templates | Roles |
| --- | --- | --- | --- |
| `/catalog/` | `catalog:index` | `CatalogIndexView`, `catalog/index.html` | ADMIN only |
| `/catalog/teachers/...` | `teacher_list`, `teacher_create`, `teacher_edit` | `Teacher*View`, `catalog/teachers_list.html`, `catalog/teacher_form.html` | ADMIN only |
| `/catalog/subjects/...` | `subject_list`, `subject_create`, `subject_edit` | `Subject*View`, `catalog/subjects_list.html`, `catalog/subject_form.html` | ADMIN only |
| `/catalog/assignments/...` | `assignment_list`, `assignment_create`, `assignment_delete`, `copy_year` | assignment/copy templates | ADMIN only |
| `/catalog/prices/...` | `price_list`, `price_create`, `price_edit`, `price_delete`, `copy_prices_year` | price/copy templates | ADMIN only |
| `/catalog/discounts/...` | `discount_*`, `family_discount_*`, `record_discount_*` | discount templates | ADMIN only |
| `/catalog/familyid-reservations/...` | reservation routes | `FamilyIdReservation*View`, `catalog/familyid_reservations.html` | ADMIN only |
| `/catalog/sync-from-legacy/` | `sync_from_legacy` | `SyncFromLegacyView`, `catalog/sync_from_legacy.html` | ADMIN only |
| `/catalog/categories/...` | category copy/list/create/edit/delete/subjects | category templates | ADMIN only |
| `/catalog/groups/<year>/` | `group_list` | `DisciplineGroupListView`, `catalog/group_list.html` | ADMIN only |
| `/catalog/groups/<year>/<pk>/` | `group_detail` | `DisciplineGroupDetailView`, `catalog/group_detail.html` | ADMIN, OPERATOR |
| group duration/size/toggle/prepare legacy | `group_duration_add`, `group_size_add`, `group_toggle_scaling`, `group_prepare_legacy` | redirect/action views | ADMIN only |
| group bulk apply | `group_apply_preview`, `group_apply_bulk` | `BulkApplyPreviewView`, `BulkApplyCategoryPriceView`, `catalog/group_bulk_apply_preview.html` | ADMIN, OPERATOR |
| `/catalog/api/groups/<year>/<pk>/size/` | `group_size_api` | `GroupSizeApiView` | ADMIN only |

Important current mismatch:

- `CatalogAdminMixin`: ADMIN only.
- `CatalogEditorMixin`: ADMIN and OPERATOR, intended for pricing workflows where OPERATOR may create `PENDING` changes.
- `PROMPT_162` is not applied. Because catalog index and group list are ADMIN-only, OPERATOR may have backend access to group detail/bulk apply but incomplete UI navigation to reach it.

## Top-level APIs

| URL | View/module | Notes |
| --- | --- | --- |
| `/api/accounts/` | includes `apps.accounts.urls` | Same user routes under API prefix; not a separate DRF contract. |
| `/api/karteien/` | includes `apps.karteien.urls` | Same karteien routes under API prefix. |
| `/api/approvals/` | includes `apps.approvals.urls` | Same approvals routes under API prefix. |
| `/api/history/` | `apps.history.urls` | History API endpoints. |
| `/api/notifications/` | `apps.notifications.urls` | Notification polling/dropdown endpoints. |
| `/api/legacy-import/` | `apps.legacy_import.urls` | Legacy import endpoints. |
| `/api/reporting/` | `apps.reporting.urls` | Reporting endpoints. |
| `/api/catalog/subjects/` | `SubjectsApiView` | Dynamic form data. |
| `/api/catalog/teachers/` | `TeachersApiView` | Dynamic form data. |
| `/api/catalog/prices/` | `PricesApiView` | Dynamic form data. |

## QA hotspots

- Standard Kartei editor and approvals snapshot editors are separate workflows; do not assume one test covers both.
- UI button visibility must be verified against server mixins, especially OPERATOR vs ADMIN.
- PRICELIST V2 preview/apply/approve paths are known to be split before prompts 166-178.
- Minimal browser smoke tests exist in `backend/tests/browser/`, but they cover only fixture-backed non-destructive route/permission checks. Manual browser QA evidence remains required for visual/navigation/AJAX/responsive changes and flows outside that smoke scope.
