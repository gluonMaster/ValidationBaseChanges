# Browser/UI QA - KindEltern Web

Last updated: 2026-05-27. Created by `PROMPT_183_BROWSER_UI_QA_PROTOCOL_AND_PLAYWRIGHT_FEASIBILITY.md`; automated smoke baseline added by `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md`.

This protocol is mandatory context for prompts that change UI, navigation, templates, role permissions, AJAX behavior, or user-visible workflow. It supplements `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md`, backend pytest tests, and the minimal Playwright smoke layer.

## Purpose

Browser QA catches problems that `manage.py check`, migration checks, and Django test-client tests cannot see:

- broken navigation and missing entry points;
- role-specific button drift between templates and server permissions;
- Bootstrap dropdown/modal/collapse/offcanvas failures;
- AJAX/live-search/billing-preview behavior and browser console errors;
- visual regressions in dense Kartei, approvals, and catalog workflows.

For any prompt that edits templates, view context used by templates, role mixins, redirects, JavaScript, or user-visible messages, the final answer must include either browser QA evidence or a clear reason why it was not run.

## Local Run Prerequisites

Primary local run path is Docker:

```bash
cp .env.example .env
docker compose build web
docker compose up -d
docker compose ps
docker compose logs web --tail=50
```

Expected app URL: `http://localhost:8000/`.

## Local Browser QA Credentials

For MCP/manual browser checks against the local dev database, use the ignored local file:

```text
.codex-local/browser-qa-credentials.md
```

This file is intentionally outside git. If it exists, agents may read it to log in as `ADMIN`, `OPERATOR`, `SUPERADMIN`, and `USER`. Do not print passwords in final answers and do not copy this file into committed docs or screenshots.

Expected login selectors:

- username: `#id_username`
- password: `#id_password`
- submit: `button[type="submit"]`

If the credentials file is missing or a required role is blank, browser QA for that role is incomplete. Ask the user for the missing credential or report the blocker.

Required local data for meaningful browser QA:

- users for all roles: `ADMIN`, `OPERATOR`, `SUPERADMIN`, `USER`;
- at least one normal `KarteiRecord`;
- at least one `PENDING` record/change;
- at least one `DECLINED` record/change when declined flows are touched;
- at least one catalog group with eligible records when group pricing/bulk apply is touched.

Users can be created through Django Admin as described in `RUNNING_LOCAL_QUICKSTART.md`. If the database is empty and the touched flow needs real records, the browser QA result is incomplete and must say so.

## Automated Smoke Tests

`PROMPT_185` added fixture-backed automated browser smoke tests under `backend/tests/browser/`.

Install dependencies and the Chromium browser binary:

```bash
pip install -r backend/requirements.txt
python -m playwright install chromium
```

Run the automated smoke baseline:

```bash
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

Covered by the smoke baseline:

- login/root redirect for `ADMIN`, `OPERATOR`, `SUPERADMIN`, `USER`;
- ADMIN Kartei list/create/edit entry points for a normal fixture record;
- OPERATOR direct standard-editor redirects for `PENDING` and `DECLINED` fixture records;
- SUPERADMIN pending overview and War/Ist detail page open;
- USER cabinet and read-only record detail route.

Not covered by the smoke baseline:

- destructive approve/decline/bulk apply actions;
- catalog group pricing/bulk apply fixtures;
- visual layout assertions, browser console review, AJAX/live-search interaction, responsive viewports.

The `--browser` CLI option is provided by `pytest-playwright`, so install `backend/requirements.txt` before using the smoke command. If the Chromium browser binary is missing, the browser fixture skips with an explicit reason; install it with `python -m playwright install chromium`.

Run backend and browser pytest sessions sequentially, not in parallel, when both use the same PostgreSQL service/test database. The browser tests use `django live_server` and transactional fixtures; an interrupted or concurrent run can leave a stale `test_<POSTGRES_DB>` database, commonly `test_kindeltern`. Prefer `--reuse-db` for repeated local runs and use `--create-db` after migration changes.

If a run fails with `database "test_kindeltern" already exists`, stop other pytest/live_server processes and clean the stale test DB. With Docker Compose defaults, run:

```bash
docker compose exec db sh -lc 'TEST_DB="test_${POSTGRES_DB:-kindeltern}"; psql -U "$POSTGRES_USER" -d postgres -c "SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname = '"'"'${TEST_DB}'"'"';"; dropdb -U "$POSTGRES_USER" --if-exists "$TEST_DB"'
```

## Mandatory Smoke Routes

Run these route checks before or after any broad UI/navigation prompt. Use the current year shown by the app unless the prompt is explicitly year-specific.

| Role | Routes / flow | Required observations |
| --- | --- | --- |
| Anonymous | `/accounts/login/` | Login page renders; no authenticated nav leaks. |
| ADMIN | `/` -> `/karteien/`; `/karteien/`; `/karteien/create/`; one normal `/karteien/<pkid>/edit/`; `/approvals/pending/`; `/approvals/declined/`; `/catalog/` | Navbar shows Kartei, Pending, Declined, Katalog; create/edit buttons exist where expected; filters and record table render; no console errors. |
| OPERATOR | `/` -> `/karteien/`; `/karteien/`; one normal `/karteien/<pkid>/edit/`; one pending/declined `/karteien/<pkid>/edit/`; if group flow touched, direct `/catalog/groups/<year>/<group_pk>/` and `/catalog/groups/<year>/<group_pk>/apply-preview/` | Normal edit is available; pending/declined standard edit is blocked/redirected; operator catalog navigation gap from `PROMPT_162` is still expected until fixed; direct editor group paths must match server permissions. |
| SUPERADMIN | `/` -> `/approvals/superadmin/pending/`; `/approvals/superadmin/pending/`; one `/approvals/superadmin/pending/<pending_pk>/`; `/approvals/superadmin/neu/`; `/karteien/` | Superadmin dropdown works; War/Ist page renders decision controls; NeuList and read-only Kartei work; no edit/create buttons are exposed in Kartei. |
| USER | `/` -> `/user/`; `/user/search/`; `/user/record/<pkid>/`; `/user/record/<pkid>/history/`; `/karteien/` | User cabinet renders; record views are read-only; Kartei list/detail do not expose write actions. |

## Flow-Specific Checks

Use `docs/QA_MATRIX.md` to select the focused scenario. Minimum expectations:

- Kartei list/template changes: desktop screenshot of `/karteien/`, active filters, pagination/table area, and role-specific action buttons for ADMIN and read-only role.
- Kartei form changes: screenshot of create/edit header, affected form section, validation/alert area, save/cancel controls, and browser console after interacting with dynamic fields.
- Billing preview/category apply changes: capture the preview state before apply, the submitted/confirmation state, pending status, and any warning badges.
- Approvals changes: capture pending list, War/Ist detail, approve/decline decision area, declined overview/detail when relevant.
- Catalog group changes: capture group list/detail, monthly summary, bulk preview table, and apply action visibility for ADMIN and OPERATOR when applicable.
- Navigation/permissions changes: check both the visible UI and direct URL access for every affected role.
- JavaScript changes: check browser console and failed network requests after exercising the changed control.

For broad layout/nav prompts, check both desktop and a narrow viewport. For small text-only template changes, one representative viewport is enough unless the text affects buttons, table columns, or nav.

## What To Record In Final Answer

Report browser QA in a compact evidence block:

- tool used: manual browser, Playwright MCP, Chrome DevTools MCP, or automated Playwright (`pytest-playwright`);
- base URL and run mode: for example `http://localhost:8000/` via Docker;
- roles checked;
- routes/flows checked;
- screenshots captured or visual observations made;
- browser console/network issues, if any;
- skipped checks and exact reason, if any.

If browser QA was not run for a UI prompt, the final answer must say why and list the highest-risk manual checks still needed.

## MCP status and manual browser tools

This repository confirms only repository-level browser QA assets: `pytest-playwright` automated smoke tests and this manual browser QA protocol. It does **not** contain `mcpServers` configuration and does not prove that Playwright MCP or Chrome DevTools MCP is installed for a specific agent runtime.

Treat these as separate layers:

1. `pytest-playwright` automated smoke tests in `backend/tests/browser/`;
2. manual browser QA using any real browser or MCP-capable browser tool;
3. actual Playwright MCP / Chrome DevTools MCP installation/configuration in the agent runtime.

If MCP is unavailable, do not claim MCP evidence. Use the automated smoke command and/or a manual browser check, and state the blocker plus the remaining highest-risk checks.

## Manual MCP Check vs Automated Playwright

Manual browser/MCP checks are acceptable when:

- the prompt is docs-only;
- the UI change is one-off, visual, or navigation-oriented;
- the flow requires ad hoc human judgment or data that is not fixture-backed;
- the changed behavior is already covered by backend tests and browser QA is only visual confirmation.

Automated Playwright tests should be added or updated when:

- a permission/security rule is changed;
- a repeated smoke route must stay stable across future prompts;
- JavaScript/AJAX behavior is central to the feature;
- a known UI regression is fixed and should not return;
- preview/apply/approval parity is being stabilized;
- the prompt changes a flow that future prompts will refactor again.

Do not claim broad automated UI coverage from the smoke baseline alone. Claim only the concrete routes/roles that were run.

## Playwright Baseline Decision

Current state after `PROMPT_179...PROMPT_185`:

- Stable local app run exists via Docker Compose and `http://localhost:8000/`.
- Backend pytest harness exists (`pytest`, `pytest-django`, `pytest.ini`, `backend/tests/`).
- Browser automation uses `pytest-playwright` with `pytest-django live_server`; no Node/package.json path was added.
- Browser smoke fixture data is isolated in pytest DB fixtures: role users, normal record, pending change, declined record.
- The baseline is intentionally non-destructive. Approve/decline/bulk apply remain manual or future isolated automation.

Decision: use `pytest -m browser backend/tests/browser --browser chromium --reuse-db` as a minimal route/permission regression gate, and continue using this document for manual evidence outside the smoke scope. This decision is not an MCP installation claim.
