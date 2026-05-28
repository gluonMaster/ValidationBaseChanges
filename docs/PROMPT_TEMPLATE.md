# Prompt Template - KindEltern Web

Use this template for future `PROMPT_XXX_*.md` files. Keep prompts atomic: one prompt = one bounded change.

```markdown
# PROMPT XXX - Short Title

## Status and Order

- Status: planned | applied | superseded
- Must run before/after:
- Related prompts:
- This prompt changes: code | docs | tests | data/migrations | QA only

## Context Files

Read first:

- `CLAUDE.md`
- `docs/Status.md`
- `docs/ProjectMap.md`
- `docs/QA_MATRIX.md`
- `docs/UI_ROUTE_MAP.md`
- `docs/BROWSER_QA.md`
- `docs/Spec.md`
- `docs/Decisions.md`

Task-specific files:

- `path/to/file.py`
- `path/to/template.html`
- `path/to/test.py`

## Problem

Describe the current user-visible or engineering problem.

Be explicit about current behavior vs target behavior. If this is a known defect, name the defect and link the source document/test.

## Exact Task

1. Make the smallest code/docs/test changes needed.
2. Preserve existing behavior outside this scope.
3. Update or add tests/manual QA evidence.
4. Update canonical docs if behavior/status changes.

## Non-goals

- Do not refactor unrelated modules.
- Do not change Django models/migrations unless explicitly required.
- Do not change PRICELIST V2 stabilization semantics outside this prompt.
- Do not apply unrelated prompts.

## Invariants

- `KarteiRecord.pkid` is Django PK.
- `KarteiRecord.id` is legacy Access/Excel ID.
- Domain key is `(year, id)`.
- `RecordStatus.NORMAL == ""`; normal checks use `not record.status`.
- UI text/messages are German.
- Code comments are English.
- Prompt files are Russian unless the task explicitly uses an English docs template.
- For behavior prompts, update `docs/Status.md` and add automated test coverage or explicit manual QA.

Add task-specific invariants here:

- ...

## Expected Code Changes

- `backend/apps/...`: expected change.
- `backend/templates/...`: expected UI/template change.
- `backend/tests/...`: expected tests.

If docs-only, say explicitly: "No Django code changes."

## Tests / Manual QA

Automated tests to add/update:

- `backend/tests/test_...py::test_...`

Manual QA checklist:

- Role:
- URL:
- Setup:
- Steps:
- Expected:

## Browser/UI verification

Required if this prompt changes templates, navigation, redirects, role-specific UI, JavaScript/AJAX, or user-visible workflow. Use `docs/BROWSER_QA.md`.

Browser flows/screenshots needed:

- Role:
- URL / flow:
- Viewport: desktop | narrow | both
- Screenshot/observation to capture:
- Console/network checks:

Real browser verification is mandatory when:

- UI/navigation behavior changes across roles;
- JavaScript/AJAX behavior is changed;
- permissions are represented by buttons/links that can drift from server checks;
- a known UI regression is fixed.

Use Playwright MCP / Chrome DevTools MCP only if the agent runtime actually provides it. Repository `pytest-playwright` smoke tests do not prove MCP availability. If MCP/browser verification is skipped, state the exact blocker and the highest-risk manual checks still required.

Known defects / xfail:

- `pytest.mark.xfail(..., reason="...")` when target behavior is intentionally not implemented yet.

## Docs Update

Update when relevant:

- `docs/Status.md`
- `docs/QA_MATRIX.md`
- `docs/UI_ROUTE_MAP.md`
- `docs/ProjectMap.md`
- `PROMPTS_OVERVIEW.md`
- App README if local behavior changed.

## Risks / Rollback

Risks:

- Data risk:
- Permission risk:
- UI/navigation risk:
- Migration risk:

Rollback:

- Files to revert:
- Migration rollback, if any:
- Data repair, if any:

## Verification Commands

For code changes:

```bash
python backend/manage.py check
python backend/manage.py makemigrations --check --dry-run
pytest -m "not browser" --reuse-db
```

For docs-only changes:

```bash
git diff -- docs/QA_MATRIX.md docs/UI_ROUTE_MAP.md docs/PROMPT_TEMPLATE.md docs/Status.md docs/ProjectMap.md CLAUDE.md PROMPTS_OVERVIEW.md
```

Add focused commands here:

```bash
pytest -m "not browser" backend/tests/test_...py --reuse-db
# If browser/UI route smoke is relevant:
python -m playwright install chromium
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

## Expected Result

At the end, report:

- files changed;
- tests/manual QA performed;
- known remaining risks;
- whether docs/status were updated.
```
