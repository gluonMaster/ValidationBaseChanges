# QA Matrix - KindEltern Web

Last updated: 2026-05-27. Created by `PROMPT_180_DOCS_QA_MATRIX_UI_ROUTE_MAP_PROMPT_TEMPLATE.md`; browser QA protocol added by `PROMPT_183_BROWSER_UI_QA_PROTOCOL_AND_PLAYWRIGHT_FEASIBILITY.md`; automated browser smoke baseline added by `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md`.

Use this as the compact test/manual-QA index for future prompts. `PROMPT_179` added a minimal backend pytest harness (`pytest.ini`, `backend/tests/`). `PROMPT_185` added minimal `pytest-playwright` smoke tests in `backend/tests/browser/`; use `docs/BROWSER_QA.md` for mandatory manual browser verification when UI prompts touch templates, navigation, JavaScript, redirects, role-specific UI, or flows outside the automated smoke scope.

| Area | Scenario | Roles | URL / entry point | Expected behavior | Current status | Automated coverage | Manual QA | Known defects / xfail |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Auth / routing | Login and role redirect | USER, ADMIN, OPERATOR, SUPERADMIN | `/accounts/login/`, `/` | USER -> `/user/`; SUPERADMIN -> `/approvals/superadmin/pending/`; ADMIN/OPERATOR -> `/karteien/`. | Implemented in `role_based_redirect`. | Browser smoke: `test_login_and_root_redirect_per_role`. | Login as each role and verify navbar/landing details when navigation UI changes. | None tracked. |
| Kartei list | List, filters, live search | ADMIN, OPERATOR, SUPERADMIN, USER | `/karteien/`, `/karteien/live-search/` | All four roles can view records; ADMIN/OPERATOR get edit-oriented UI; SUPERADMIN/USER are read-only. | Implemented. | Browser smoke covers ADMIN list entry point only; live-search/filter behavior remains manual. | Check list loads, filters update, live search returns only allowed view data. | Browser smoke does not cover filters/live-search/visual drift. |
| Kartei detail | Read-only detail view | ADMIN, OPERATOR, SUPERADMIN, USER | `/karteien/<pkid>/` | All viewer roles can open detail; action buttons must match server permissions. | Implemented. | Not covered yet. | Verify detail for normal, pending, declined, legacy billing records. | UI/server drift must be watched manually. |
| Kartei create/edit | Create and edit normal record | ADMIN, OPERATOR | `/karteien/create/`, `/karteien/<pkid>/edit/` | ADMIN/OPERATOR can create/edit normal records; operator restrictions still apply. SUPERADMIN/USER cannot edit. | Implemented. | Browser smoke: `test_admin_sees_kartei_create_and_edit_entry_points`; backend access tests cover OPERATOR pending/declined blocking. | Create/edit one normal record as ADMIN and OPERATOR when form behavior changes; verify SUPERADMIN/USER redirect. | Browser smoke checks ADMIN entry points only; OPERATOR normal edit remains manual unless touched. |
| Pending/declined edit | Admin vs Operator edit for `PENDING` / `DECLINED` | ADMIN, OPERATOR | `/karteien/<pkid>/edit/` | ADMIN can use standard editor for `PENDING`/`DECLINED`; OPERATOR is blocked and redirected to detail. | Implemented. | Backend: `test_admin_can_open_standard_editor_for_pending_record`, `test_admin_can_open_standard_editor_for_declined_record`, `test_operator_is_blocked_from_editing_pending_or_declined_records`; browser smoke: `test_operator_standard_editor_redirects_for_pending_and_declined_records`. | Confirm messages and UI buttons for pending/declined records. | Approvals app `AdminEditorMixin` name/docs say Admin, but implementation uses `user.can_edit_kartei`; verify OPERATOR exposure before relying on docs. |
| Approvals decision | Approve / decline single pending change | SUPERADMIN | `/approvals/superadmin/pending/`, `/approvals/superadmin/pending/<pk>/` | SUPERADMIN reviews War/Ist and applies approve/decline; live record should reflect decision according to current workflow. | Implemented, but PRICELIST V2 target semantics not stabilized. | Browser smoke opens pending overview and War/Ist detail only: `test_superadmin_opens_pending_overview_and_war_ist_detail`. | Create risky change, approve it, decline another, verify status, history, notification cleanup. | Target frozen payload / no-prewrite behavior not applied; destructive approve/decline is intentionally not automated. |
| Approvals bulk decision | Approve all / decline all | SUPERADMIN | `/approvals/superadmin/approve-all/`, `/approvals/superadmin/decline-all/` | Bulk decision applies to all pending changes available to SUPERADMIN. | Implemented. | Not covered yet. | Run on a controlled test set only; verify counts and statuses. | Same frozen-payload risk as single decision. |
| Declined resubmit | Review declined and resubmit fix | ADMIN, possibly OPERATOR by current mixin | `/approvals/declined/`, `/approvals/declined/<pk>/`, `/approvals/declined/<pk>/edit/`, `/approvals/declined/<pk>/apply-fix/` | Admin reviews declined snapshot, edits/applies fix, record returns to `PENDING`. | Implemented legacy/current flow. | Not covered yet. | Decline a pending change, edit/apply fix, confirm new pending appears for SUPERADMIN. | Permission ambiguity: `AdminEditorMixin` currently allows `user.can_edit_kartei`, likely OPERATOR too. |
| Pending info/edit | Admin pending list and pending snapshot edit | ADMIN, possibly OPERATOR by current mixin | `/approvals/pending/`, `/approvals/pending/<pk>/`, `/approvals/pending/<pk>/edit/` | Admin can inspect pending changes and update pending snapshot/comment before SUPERADMIN decision. | Implemented. | Snapshot metadata preservation covered by `backend/tests/test_pending_snapshot_metadata.py`. | Update a pending comment/snapshot, verify SUPERADMIN sees updated data and optimistic guard works. | Permission ambiguity: same `AdminEditorMixin` issue. |
| NeuList per-year | New records list tracks seen records per year | SUPERADMIN | `/approvals/superadmin/neu/?year=YYYY`, `/approvals/superadmin/mark-seen/` | New records are calculated per `(year, id)` and per-year `last_seen_by_year`. | Implemented. | Covered by `backend/tests/test_neulist.py`. | Import/create records in two years, mark seen in one year, verify the other year remains independent. | None tracked. |
| Billing preview | Live billing preview in record form | ADMIN, OPERATOR | `/karteien/<pkid>/billing-preview/` | Preview should match current form inputs and explain months/base/discount effects. | Implemented, but not target-stable for PRICELIST V2. | `test_target_billing_preview_and_category_apply_preview_use_same_pipeline` (`xfail`). | Change price/months/discount fields and compare preview vs saved result. | Preview/apply parity is a known PRICELIST V2 defect. |
| Category apply | Apply suggested category price to one record | ADMIN, OPERATOR | `/karteien/<pkid>/apply-price/preview/`, `/karteien/<pkid>/apply-price/` | Preview and apply should route billing changes through approval when risky; operator month restrictions apply. | Implemented current flow. | `test_single_category_apply_creates_pending_proposal_with_current_metadata`; target no-prewrite `xfail` in `test_target_category_apply_decline_does_not_need_live_prewrite_or_rollback_marker`. | Preview/apply for ADMIN and OPERATOR; verify pending/live changes and history. | Known split-brain: apply path may prewrite billing/context before target stabilization. |
| Bulk apply group flow | Apply category/group price to group records | ADMIN, OPERATOR for editor paths | `/catalog/groups/<year>/<pk>/apply-preview/`, `/catalog/groups/<year>/<pk>/apply-bulk/` | ADMIN/OPERATOR can use editor bulk apply paths; operator restrictions apply; affected records should become pending when required. | Backend editor paths allow ADMIN/OPERATOR. UI navigation is incomplete for OPERATOR. | `test_bulk_category_apply_creates_pending_proposals_for_group_records`; `test_bulk_category_apply_operator_restrictions_block_past_months`; target no-prewrite `xfail` in `test_target_bulk_apply_does_not_prewrite_live_base_amounts`. | Open group detail, preview bulk apply, apply to a small group, verify pending counts. | `PROMPT_162` not applied: OPERATOR group navigation/catalog entry remains admin-only (`xfail`). |
| Contract changes | Contract type/status changes | ADMIN, OPERATOR | `/karteien/<pkid>/contract-type/change/`, `/karteien/<pkid>/contract-status/change/` | Changes create pending approval entries and must not silently corrupt approved billing state. | Implemented current flow. | `test_contract_status_active_paused_terminated_month_values`; group size status coverage in `test_group_size_counts_active_and_paused_but_excludes_terminated`. | Change type/status for active, paused, terminated cases; verify pending, history, billing zeroing. | Target single source of truth for contract timeline not fully stabilized. |
| Nachhilfe | NH pricing and discounts | ADMIN, OPERATOR | Record form, billing preview, category/group/catalog setup | Nachhilfe uses hourly/UE semantics; discounts must not apply. | Implemented partially/currently documented. | `test_nachhilfe_month_values_do_not_apply_discounts`; dual schema target `xfail` in `test_target_schema_allows_dual_group_and_individual_nachhilfe_categories`. | Create/edit Nachhilfe record; verify hours, price, no discounts, preview/save parity. | `PROMPT_160` not applied: schema/bootstrap decision around dual Nachhilfe category remains open. |
| LEGACY -> AUTO | Recalculate legacy months into AUTO | ADMIN, OPERATOR with restrictions | `/karteien/<pkid>/edit/` (`Monate neu berechnen`) | Legacy imported month values can be converted/recalculated through explicit action. | Implemented current flow. | `test_explicit_legacy_to_auto_conversion_updates_only_touched_months`. | Use a LEGACY record, trigger recalculation, verify base/month values and pending behavior. | Target no-prewrite frozen payload not applied. |
| Decline rollback | Declining pending billing/context proposal | SUPERADMIN | `/approvals/superadmin/pending/<pk>/` | Current workflow may need rollback for prewritten live fields; target v2 should avoid rollback by keeping live approved-state until approval. | Known target defect. | `test_current_decline_rolls_back_old_base_amounts_marker`; target no-prewrite xfails in `backend/tests/test_known_defects_xfail.py` and `test_target_category_apply_decline_does_not_need_live_prewrite_or_rollback_marker`. | Create billing/context pending change, decline, verify live record fields are restored or were never changed. | `xfail`: no-prewrite frozen payload semantics not applied. |
| Preview/apply parity | Preview result equals applied/saved result | ADMIN, OPERATOR | Billing preview, category apply, bulk apply, approve | Preview, form save, apply, and approve should use one billing pipeline. | Known PRICELIST V2 target defect. | `test_target_billing_preview_and_category_apply_preview_use_same_pipeline` (`xfail`). | For a fixed fixture, compare preview JSON, pending snapshot, approved record months/base amounts. | Known split-brain across preview/apply/approval paths. |

## Browser QA status

After `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md`, the scenarios below are either partially covered by fixture-backed browser smoke or still manual/browser-required for UI prompts. Use the role and route details from `docs/BROWSER_QA.md`.

| Scenario group | Browser QA status | Required roles | Required evidence |
| --- | --- | --- | --- |
| Auth / role redirects | Manual browser required when login, redirect, navbar, or access flow changes. | ADMIN, OPERATOR, SUPERADMIN, USER | Root redirect result, landing page screenshot/observation, navbar visibility. |
| Kartei list/detail/form | Manual browser required for template, filter, button, live-search, form, or billing preview changes. | ADMIN, OPERATOR, SUPERADMIN, USER as applicable | `/karteien/`, affected detail/edit/create route, role-specific buttons, console/network check for AJAX. |
| Pending/declined admin flows | Manual browser required when approvals admin templates, pending/declined edit, or notification links change. | ADMIN and OPERATOR when permission ambiguity is relevant | Pending/declined list/detail/edit route, visible actions, direct URL permission result. |
| Superadmin approvals | Manual browser required for War/Ist, pending overview, approve/decline, NeuList, or notification changes. | SUPERADMIN | Pending overview, War/Ist decision area, NeuList, history link if touched. |
| Catalog group pricing | Manual browser required for group list/detail/bulk apply templates or route access changes. | ADMIN, OPERATOR | Group detail and bulk preview; note expected OPERATOR navigation gap until `PROMPT_162` is applied. |
| User read-only cabinet | Partially covered by browser smoke for dashboard/root and record detail read-only route; manual browser still required when search/history UI changes. | USER | Dashboard/search/detail/history route and absence of write actions. |
| Automated Playwright | Present as minimal smoke baseline in `backend/tests/browser/`. | ADMIN, OPERATOR, SUPERADMIN, USER | Run `python -m playwright install chromium` once, then `pytest -m browser backend/tests/browser --browser chromium --reuse-db`. Do not run in parallel with backend pytest against the same PostgreSQL test DB. |

## Current automated test index

- `backend/tests/test_core_invariants.py`: core model invariants such as `RecordStatus.NORMAL == ""` and `(year, id)` domain key.
- `backend/tests/test_neulist.py`: per-year NeuList state.
- `backend/tests/test_record_edit_access.py`: ADMIN vs OPERATOR edit behavior for `PENDING` / `DECLINED`.
- `backend/tests/test_pending_snapshot_metadata.py`: reserved metadata keys in `PendingChange.snapshot`.
- `backend/tests/test_known_defects_xfail.py`: current known target defects (`PROMPT_162`, no-prewrite frozen payload).
- `backend/tests/test_pricelist_v2_regression_cases.py`: PRICELIST V2 regression/xfail coverage added by `PROMPT_182`.
- `backend/tests/browser/test_role_smoke.py`: minimal non-destructive Playwright smoke for role redirects, ADMIN Kartei entry points, OPERATOR pending/declined edit redirects, SUPERADMIN pending overview + War/Ist detail, USER read-only cabinet/detail route.

Automated browser test command:

```bash
python -m playwright install chromium
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

## PROMPT_182 PRICELIST V2 Regression Index

| Scenario | Automated coverage | Status |
| --- | --- | --- |
| `base_amounts` in `AUTO` is historical base, not silent cache | `test_auto_base_amounts_are_historical_base_not_silent_price_cache` | Passing |
| Single-record category apply creates proposed `base_amounts`, proposed `month_*`, `PendingChange`, and `_old_base_amounts` metadata | `test_single_category_apply_creates_pending_proposal_with_current_metadata` | Passing, documents current compatibility path |
| Bulk apply creates pending proposals for group records | `test_bulk_category_apply_creates_pending_proposals_for_group_records` | Passing, documents current compatibility path |
| Bulk apply does not bypass Operator past-month restrictions | `test_bulk_category_apply_operator_restrictions_block_past_months` | Passing |
| Bulk apply target no-live-prewrite behavior | `test_target_bulk_apply_does_not_prewrite_live_base_amounts` | `xfail`: target snapshot v2 behavior, not implemented before `PROMPT_166...178` |
| Contract status `ACTIVE` / `PAUSED` / `TERMINATED` affects month values | `test_contract_status_active_paused_terminated_month_values` | Passing |
| Group size counts `ACTIVE + PAUSED` and excludes `TERMINATED` | `test_group_size_counts_active_and_paused_but_excludes_terminated` | Passing |
| `Nachhilfe` discounts are skipped | `test_nachhilfe_month_values_do_not_apply_discounts` | Passing |
| Dual `Nachhilfe` GROUP/INDIVIDUAL category schema target | `test_target_schema_allows_dual_group_and_individual_nachhilfe_categories` | `xfail`: `PROMPT_160` not applied |
| Explicit `LEGACY -> AUTO` conversion | `test_explicit_legacy_to_auto_conversion_updates_only_touched_months` | Passing |
| Decline rollback via current `_old_base_amounts` compatibility marker | `test_current_decline_rolls_back_old_base_amounts_marker` | Passing |
| Target category apply decline without live prewrite or rollback marker | `test_target_category_apply_decline_does_not_need_live_prewrite_or_rollback_marker` | `xfail`: target snapshot v2 behavior, not implemented before `PROMPT_166...178` |
| Preview/apply parity uses one billing pipeline | `test_target_billing_preview_and_category_apply_preview_use_same_pipeline` | `xfail`: known PRICELIST V2 split-brain before stabilization |

## Verification commands

For docs-only changes, link consistency and file review are sufficient. For any behavior/code prompt, run:

```bash
python backend/manage.py check
python backend/manage.py makemigrations --check --dry-run
pytest -m "not browser" --reuse-db
python -m playwright install chromium
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```
