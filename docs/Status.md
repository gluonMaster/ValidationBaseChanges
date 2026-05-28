# Status - KindEltern Web

Дата фиксации: 2026-05-27.

Этот документ — оперативный статус проекта для нового AI-чата. Если он расходится со старыми разделами `ARCHITECTURE.md`, `DOMAIN_MODEL.md`, `docs/Spec.md` или `docs/Decisions.md`, считать этот файл более актуальным.

## 1. Текущее состояние

- Проект — Django + PostgreSQL CRM для замены Excel/VBA/Access workflow.
- Основные apps: `accounts`, `karteien`, `catalog`, `approvals`, `history`, `notifications`, `reporting`, `legacy_import`.
- Базовые сценарии картотеки, approvals, history, catalog, discounts, import и AUTO-начислений реализованы.
- `KarteiRecord.pkid` — Django PK; `KarteiRecord.id` — Access/Excel ID; доменный ключ: `(year, id)`.
- `RecordStatus.NORMAL` хранится как пустая строка `""`.
- NeuList работает per-year через `SuperadminState.last_seen_by_year`.
- Admin сейчас может редактировать `PENDING`/`DECLINED` через стандартный editor; Operator в этих статусах блокируется.
- SAFE-изменения в текущей реализации могут применяться напрямую к live `KarteiRecord`; RISKY-изменения идут через `PendingChange`.

## 2. Prompt-state

Считать примененными как development/code-chain:

- `PROMPT_01...PROMPT_157`, кроме `PROMPT_131`;
- `PROMPT_159`;
- `PROMPT_161`.

Считать не примененными как development/code-chain:

- `PROMPT_131` — optional;
- `PROMPT_158` — fix для `ContractStatusChangeView ACTIVE`;
- `PROMPT_160` — schema/bootstrap решение вокруг `Nachhilfe`;
- `PROMPT_162` — operator-навигация к группам; **не применен**, код все еще в старом состоянии;
- `PROMPT_179` — применен: backend pytest/pytest-django test harness baseline;
- `PROMPT_180` — применен: docs QA matrix, UI route map, prompt template;
- `PROMPT_181` — применен: app-level approvals README и `KarteiRecordUpdateView` comments/docstrings cleanup, без изменения behavior;
- `PROMPT_182` — применен: PRICELIST V2 regression/xfail cases before stabilization refactor;
- `PROMPT_183` — применен: browser/UI QA protocol and Playwright feasibility, docs-only;
- `PROMPT_184` — применен: GPT Pro context archive refreshed после documentation/test baseline;
- `PROMPT_185` — применен: automated browser smoke tests baseline via `pytest-playwright` + `pytest-django live_server`;
- `PROMPT_166...PROMPT_178` — implementation plan для `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`; на дату фиксации не считать примененными к Django-коду.

Отдельно от development chain:

- `PROMPT_163...PROMPT_165` были использованы как GPT Pro analysis/review/spec prompts.
- Они не связаны непосредственно с цепочкой `PROMPT_162`, `PROMPT_166...PROMPT_178`, не являются Django code-change prompts и могут быть удалены как временный исторический контекст.
- `PROMPT_00_PRICELIST_SESSION_HANDOFF.md` и `PROMPT_00_PRICELIST_V2_PROMPTS_MAP.md` — исторические документы по фазе 135-151.x; не использовать их как текущую apply-очередь или доказательство статуса реализации.

Следующий новый prompt: `PROMPT_186_*.md`, если не появятся новые prompt-файлы.

## 3. Canon для прайслиста v2

- `PRICELIST_V2_STATE_AUDIT.md` — current broken state.
- `PRICELIST_V2_STABILIZATION_TZ_REVISED.md` — target state / master-spec.
- `PRICELIST_LOGIC.md` — текущая логика и статус prompt 152-162.
- `PROMPT_162` — отдельный не примененный UI/navigation fix перед stabilization chain.
- `PROMPT_179...PROMPT_185` — pre-stabilization prompts, которые должны идти до `PROMPT_166...PROMPT_178`:
  - `179` — applied: test harness baseline;
  - `180` — applied: QA matrix / UI route map / prompt template;
  - `181` — applied: app README + code-comment cleanup без behavior changes;
  - `182` — applied: PRICELIST V2 regression/xfail cases before refactor;
  - `183` — applied: browser/UI QA protocol and Playwright feasibility, docs-only;
  - `184` — applied: refresh GPT Pro context archive;
  - `185` — applied: minimal Playwright browser smoke automation.
- `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md` — применен: добавлены `pytest-playwright`, `backend/tests/browser/`, role/route smoke fixtures и команды запуска browser smoke.
- `PROMPT_166...PROMPT_178` — implementation plan/status для stabilization v2, а не уже примененная реализация. `PROMPT_166` не модифицировать; применять только после pre-stabilization prompts.

## 3.1 Рекомендуемый порядок дальнейшей работы

1. `PROMPT_179` применен — backend test harness создан.
2. `PROMPT_180` применен — созданы `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md`, `docs/PROMPT_TEMPLATE.md`.
3. `PROMPT_181` применен — обновлены `backend/apps/approvals/README.md` и устаревшие comments/docstrings без изменения behavior.
4. `PROMPT_182` применен — зафиксированы PRICELIST V2 regression/xfail cases.
5. `PROMPT_183` применен — зафиксированы browser/UI QA protocol и Playwright feasibility.
6. `PROMPT_184` применен — GPT Pro context archive пересобран с `backend/apps/familyid_reservations/`.
7. `PROMPT_185` применен — добавлен минимальный automated browser smoke baseline.
8. После этого отдельно решить/применить `PROMPT_162`.
9. Только после этого переходить к `PROMPT_166...PROMPT_178`.

## 4. Известные дефекты прайслиста v2

- Нет единого billing pipeline: preview, form save, category apply, approve/decline могут идти разными путями.
- `billing_preview_api()` все еще использует legacy-path и не полностью совпадает с category apply.
- `ApplyCategoryPriceView`, `BulkApplyCategoryPriceView`, `DisciplineGroupPrepareLegacyView` и часть form-save логики prewrite'ят billing/context state в live record до approval.
- `PendingChange.snapshot` в текущем коде в основном flat/tracked-fields snapshot с отдельными metadata keys, а не полноценный v2 frozen payload.
- Current approvals/risk model все еще опирается на legacy `TRACKED_FIELDS`; это не покрывает PRICELIST V2 non-tracked billing/context semantics.
- `base_amounts` и `month_*` могут временно относиться к разным расчетам.
- Legacy contract fields и `ContractTypeEntry` / `ContractStatusEntry` еще конкурируют как источники истины.
- `SubjectCategory` имеет unique constraint `(year, name)`, поэтому dual `Nachhilfe` для GROUP/INDIVIDUAL требует отдельного schema/bootstrap решения.
- Operator backend-доступ к bulk apply частично открыт, но prompt 162 для навигации/UI еще не применен.

## 5. Target stabilization v2

Цель `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`:

- live record остается в approved-state, пока change находится в `PENDING`;
- billing/context changes сохраняются во frozen snapshot v2 payload;
- `APPROVED` применяет reviewed payload без повторного расчета proposal;
- `DECLINED` не требует rollback live billing/context state для новых v2 snapshots;
- preview / form save / apply / approve / decline используют единый billing pipeline;
- `base_amounts` остается historical base history и не перезаписывается silently.

## 6. Testing / QA status

После применения `PROMPT_179` появилась минимальная backend test harness база:

- зависимости: `pytest`, `pytest-django`;
- конфигурация: `pytest.ini` с `DJANGO_SETTINGS_MODULE=config.settings`;
- тесты и builders: `backend/tests/`;
- быстрый запуск из корня репозитория: `pytest -m "not browser" --reuse-db`; browser tests excluded by default through `pytest.ini` and must be requested explicitly with `-m browser`.

Первые sanity/regression tests покрывают:

- `RecordStatus.NORMAL == ""`;
- доменный ключ `KarteiRecord` как `(year, id)` при surrogate PK `pkid`;
- NeuList per-year через `SuperadminState.last_seen_by_year`;
- доступ Admin к стандартному editor для `PENDING`/`DECLINED`;
- блокировку Operator на edit для `PENDING`/`DECLINED`;
- сохранение reserved metadata keys в `PendingChange.snapshot`.

После применения `PROMPT_185` появился минимальный automated browser smoke layer:

- зависимость: `pytest-playwright` (без `package.json`);
- тесты: `backend/tests/browser/`;
- запуск: `pytest -m browser backend/tests/browser --browser chromium --reuse-db`;
- установка browser binary: `python -m playwright install chromium`;
- `--browser` CLI option требует установленный `pytest-playwright` из `backend/requirements.txt`; если Chromium binary отсутствует, browser fixture пропускает tests с явной причиной.
- Запускать backend pytest и browser pytest последовательно, не параллельно, потому что оба используют PostgreSQL test DB (`test_<POSTGRES_DB>`, обычно `test_kindeltern`).

Известные target-state проверки для `PROMPT_162` и no-prewrite frozen payload semantics зафиксированы как `xfail`.

После применения `PROMPT_180` добавлены docs QA/UI baseline:

- `docs/QA_MATRIX.md` — compact matrix of critical scenarios, roles, entry points, expected behavior, automated coverage, manual QA, known defects/xfail.
- `docs/UI_ROUTE_MAP.md` — compact route/view/template/permission map, including Admin/Operator/Superadmin/User access and the `CatalogAdminMixin` vs `CatalogEditorMixin` split.
- `docs/PROMPT_TEMPLATE.md` — required template for future prompt files.

После применения `PROMPT_181` обновлены app-level README/comments около approvals/Kartei edit flow:

- `backend/apps/approvals/README.md` теперь явно указывает canonical docs, per-year NeuList через `last_seen_by_year`, legacy роль `last_seen_id`, current flat snapshot + reserved metadata keys и target snapshot v2.
- `backend/apps/karteien/views.py` обновлен только в docstrings/comments около `KarteiRecordUpdateView`; executable code и UI-тексты не менялись.

После применения `PROMPT_182` добавлен `backend/tests/test_pricelist_v2_regression_cases.py`:

- passing coverage: `AUTO` historical `base_amounts`, single/bulk category apply current proposals, Operator bulk restrictions, contract status/month values, group size `ACTIVE + PAUSED`, `Nachhilfe` no-discount rule, explicit `LEGACY -> AUTO`, current `_old_base_amounts` decline rollback.
- `xfail` target gaps: dual `Nachhilfe` schema (`PROMPT_160`), category/bulk no-live-prewrite snapshot v2 semantics, billing preview/apply parity split-brain.

После применения `PROMPT_183` добавлен browser/UI QA protocol:

- `docs/BROWSER_QA.md` — обязательный manual browser QA protocol для UI/navigation/template/role/AJAX prompts.
- `docs/PROMPT_TEMPLATE.md` теперь содержит раздел `Browser/UI verification`.
- `docs/QA_MATRIX.md` содержит секцию `Browser QA status` с manual/browser-required сценариями.
- Feasibility decision: Playwright automation возможна через `pytest-playwright` + `pytest-django live_server`; реализация перенесена в `PROMPT_185`.
- `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md` после этого применен как отдельный minimal browser smoke layer.

После применения `PROMPT_184` обновлен GPT Pro context archive:

- archive name for the current post-185 audit: `GPT_PRO_POST_179_185_REPO_OPTIMIZATION_CONTEXT.zip`;
- дата refresh: 2026-05-27;
- manifest: `GPT_PRO_POST_179_185_REPO_OPTIMIZATION_CONTEXT_MANIFEST.md`;
- archive остается selected context, не полным repo snapshot;
- явно включены `backend/apps/familyid_reservations/`, новые docs после `PROMPT_179...183`, `backend/tests/`, `pytest.ini`, selected Django apps/templates/config и prompt-chain context.

После применения `PROMPT_185` добавлен automated browser smoke baseline:

- `backend/tests/browser/conftest.py` — live-server browser fixtures for role users, normal/pending/declined records and guarded Playwright page launch.
- `backend/tests/browser/test_role_smoke.py` — non-destructive smoke coverage for login/root redirects, ADMIN Kartei create/edit entry points, OPERATOR pending/declined edit redirects, SUPERADMIN pending overview + War/Ist detail, USER read-only cabinet/detail route.
- Browser smoke intentionally does not automate approve/decline/bulk apply decisions and does not cover catalog group pricing fixtures.

Для UI/permission/behavior prompts сначала проверять `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md` и `docs/BROWSER_QA.md`; новые prompt-файлы создавать по `docs/PROMPT_TEMPLATE.md`.

## 7. Sanity-check

После code-change prompt:

```bash
python backend/manage.py check
python backend/manage.py makemigrations --check --dry-run
pytest -m "not browser" --reuse-db
python -m playwright install chromium
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

Опционально для шаблонов:

```bash
python backend/manage.py shell -c "from pathlib import Path; from django.template.loader import get_template; root=(Path('templates') if Path('templates').exists() else Path('backend/templates')); [get_template(p.relative_to(root).as_posix()) for p in root.rglob('*.html')]; print('OK')"
```

Важно: первые две команды проверяют базовую консистентность Django и миграций. `pytest -m "not browser" --reuse-db` запускает backend regression/sanity suite без browser tests. Browser smoke tests используют отдельный `pytest-playwright` + `live_server` режим и проверяют только fixture-backed non-destructive маршруты; для UI/AJAX/visual изменений manual evidence по `docs/BROWSER_QA.md` все еще нужен, если сценарий не покрыт browser smoke. Не запускать backend и browser pytest параллельно против одной PostgreSQL test DB; при stale `test_kindeltern` использовать recovery commands из quickstart/browser QA docs.

## 8. Обязательное правило для будущих prompt-файлов

Каждый prompt, который меняет поведение приложения, должен:

- обновить `docs/Status.md`;
- добавить автоматический тест или явный manual QA checklist;
- явно указать, если тесты невозможны/отложены и почему.

Для browser/UI изменений manual QA checklist и evidence по `docs/BROWSER_QA.md` остаются обязательными для сценариев вне fixture-backed smoke baseline или когда нужны visual/AJAX/responsive/console checks.
