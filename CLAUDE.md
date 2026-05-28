# CLAUDE.md — Agent Start / источник правды для AI-агента

Дата актуализации: 2026-05-27.

Этот файл — короткая стартовая инструкция для нового AI-чата. Если другие документы расходятся с этим файлом, сначала ориентироваться на `CLAUDE.md`, затем на `docs/Status.md`.

## 1. Роль и пайплайн работы

Claude Code обычно работает как **тимлид/архитектор**:

1. анализирует задачу и код;
2. обсуждает решение без преждевременных правок;
3. создает атомарные `PROMPT_XXX_*.md`;
4. ревьюит результат применения промптов и при необходимости готовит fix-prompts.

Правила:

- Не менять код напрямую, если пользователь явно не попросил.
- Markdown-документы и prompt-файлы можно создавать/редактировать по задаче.
- Промпты создавать в корне репозитория: `PROMPT_XXX_DESCRIPTION.md`.
- Следующий номер промпта = `max(существующих PROMPT_*.md) + 1`; на дату фиксации следующий номер: **186**.
- 1 prompt = 1 атомарная правка.

## 2. Язык

- Общение с пользователем: русский.
- UI-тексты и сообщения приложения: немецкий.
- Комментарии в коде: английский.
- Prompt-файлы: русский.

## 3. Что читать первым

Минимальный порядок чтения для нового агента:

1. `CLAUDE.md` — этот стартовый файл.
2. `docs/Status.md` — текущий статус, known defects, prompt-state.
3. `docs/ProjectMap.md` — карта файлов и быстрый поиск нужных модулей.
4. `docs/QA_MATRIX.md` — compact QA scenario matrix, automated coverage, manual QA, known defects/xfail.
5. `docs/UI_ROUTE_MAP.md` — compact route/view/template/role map for UI and permission work.
6. `docs/BROWSER_QA.md` — browser/manual QA protocol and automated smoke scope.
7. `docs/PROMPT_TEMPLATE.md` — required template for future prompt files.
8. Для прайслиста v2:
   - `PRICELIST_V2_STATE_AUDIT.md` — **current broken state**.
   - `PRICELIST_V2_STABILIZATION_TZ_REVISED.md` — **target state / master-spec**.
   - `PRICELIST_LOGIC.md` — текущая логика и статус prompt 152-162.
9. Затем точечно: `ARCHITECTURE.md`, `DOMAIN_MODEL.md`, `docs/Spec.md`, `docs/Decisions.md`.

Важно: `ARCHITECTURE.md`, `DOMAIN_MODEL.md`, `docs/Spec.md` и `docs/Decisions.md` содержат исторические слои. В них добавлены разделы `current implementation / known defects / target stabilization v2`; эти разделы имеют приоритет над старым текстом ниже.

Также учитывать: локальные `backend/apps/*/README.md` — вторичные app-level заметки. Если они расходятся с `CLAUDE.md`, `docs/Status.md` или `docs/ProjectMap.md`, считать их устаревшими. Пример: current NeuList работает per-year через `SuperadminState.last_seen_by_year`; описания через один глобальный `last_seen_id` — legacy/fallback/diagnostic.

## 4. Текущий prompt-state

Считать примененными как development/code-chain:

- `PROMPT_01...PROMPT_157`, кроме `PROMPT_131`;
- `PROMPT_159`;
- `PROMPT_161`.

Считать не примененными как development/code-chain:

- `PROMPT_131` — optional;
- `PROMPT_158`;
- `PROMPT_160`;
- `PROMPT_162` — **не применен**, код все еще в старом состоянии по operator group navigation;
- `PROMPT_179` — применен: backend pytest/pytest-django test harness baseline;
- `PROMPT_180` — применен: docs QA matrix, UI route map, prompt template;
- `PROMPT_181` — применен: app-level approvals README и `KarteiRecordUpdateView` comments/docstrings cleanup;
- `PROMPT_182` — применен: PRICELIST V2 regression/xfail cases before stabilization refactor;
- `PROMPT_183` — применен: browser/UI QA protocol and Playwright feasibility, docs-only;
- `PROMPT_184` — применен: GPT Pro context archive refreshed после documentation/test baseline;
- `PROMPT_185` — применен: automated browser smoke tests baseline via `pytest-playwright` + `pytest-django live_server`;
- `PROMPT_166...PROMPT_178` — implementation plan для `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`; на дату фиксации не считать примененными к коду.

Отдельно:

- `PROMPT_163...PROMPT_165` были использованы как GPT Pro analysis/review/spec prompts.
- Они не относятся непосредственно к development chain `162, 166-178`, не являются Django code-change prompts и не доказывают примененность каких-либо изменений в коде.
- Эти файлы можно считать историческим/временным контекстом и удалить позже без изменения статуса разработки.

Исторические pricelist handoff/map файлы:

- `PROMPT_00_PRICELIST_SESSION_HANDOFF.md` — handoff от 2026-02-10 для старта цепочки 135+; не является текущим статусом.
- `PROMPT_00_PRICELIST_V2_PROMPTS_MAP.md` — historical decomposition map для 135-151.x; не является текущей очередью применения.
- Текущий статус prompt-chain брать из этого файла, `docs/Status.md` и `PROMPTS_OVERVIEW.md`.

## 5. Проект и ключевые инварианты

Проект: Django + PostgreSQL CRM для замены Excel/VBA/Access workflow.

Ключевые инварианты:

- `KarteiRecord.pkid` = Django PK.
- `KarteiRecord.id` = Access/Excel ID.
- Доменный ключ записи: `(year, id)`.
- `RecordStatus.NORMAL` хранится как пустая строка `""`; проверка normal: `not record.status`.
- `base_amounts` в `months_mode=AUTO` — историческая база по месяцам, а не кэш для тихого пересчета.
- `month_*` — сохраненная approved truth на записи.

Текущее поведение approvals:

- SAFE-изменения могут применяться напрямую к live `KarteiRecord`.
- RISKY-изменения создают/обновляют `PendingChange`.
- Admin сейчас может редактировать `PENDING`/`DECLINED` через стандартный record editor; Operator в этих статусах блокируется.
- `PendingChange.snapshot` сейчас в основном flat snapshot по legacy `TRACKED_FIELDS` плюс reserved metadata keys (`_old_base_amounts`, `_pending_contract_type_entry`, `_pending_contract_status_entry`); это еще не полноценный v2 frozen payload.
- Risk classification сейчас все еще начинается от legacy `TRACKED_FIELDS`; этого недостаточно для PRICELIST V2 billing/context changes, которые меняют финансовый смысл через non-tracked поля.
- Это current implementation, не целевое состояние stabilization v2.

Целевое состояние для прайслиста v2:

- Пока change находится в `PENDING`, live record должен оставаться в approved-state.
- Billing/context изменения должны попадать во frozen snapshot payload и применяться только после approval.
- Preview / form save / apply / approve / decline должны использовать единый billing pipeline.

## 6. Дисциплина будущих промптов

### Текущее состояние тестов / QA

После применения `PROMPT_179` в репозитории есть минимальная backend test harness база: `pytest`, `pytest-django`, `pytest.ini`, `backend/tests/`.

После применения `PROMPT_182` добавлен `backend/tests/test_pricelist_v2_regression_cases.py` с passing coverage для текущих PRICELIST V2 behaviors и `xfail` target gaps: dual `Nachhilfe` schema, no-live-prewrite snapshot v2 semantics, preview/apply parity split-brain.

После применения `PROMPT_183` добавлен `docs/BROWSER_QA.md`, обновлены `docs/QA_MATRIX.md` и `docs/PROMPT_TEMPLATE.md`.

После применения `PROMPT_184` и последующего post-185 refresh текущий selected GPT Pro context archive для этого аудита: `GPT_PRO_POST_179_185_REPO_OPTIMIZATION_CONTEXT.zip`. Архив включает `backend/apps/familyid_reservations/`, новые QA/browser docs, `backend/tests/`, `pytest.ini`, selected Django code/templates/config и prompt-chain context; это не полный snapshot репозитория.

После применения `PROMPT_185` добавлен минимальный automated browser smoke baseline: `pytest-playwright` в `backend/requirements.txt`, marker `browser` в `pytest.ini`, fixtures/tests в `backend/tests/browser/`.

Текущий быстрый backend gate из корня (browser tests исключены по умолчанию через `pytest.ini`):

```bash
pytest -m "not browser" --reuse-db
```

Текущий запуск browser smoke:

```bash
python -m playwright install chromium
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

Не запускать backend pytest и browser pytest параллельно против одной PostgreSQL test DB. Если был оборванный запуск или ошибка `database "test_kindeltern" already exists`, остановить лишние pytest/live_server процессы и выполнить cleanup/recovery из `RUNNING_LOCAL_QUICKSTART.md` / `docs/BROWSER_QA.md`.

Browser smoke покрывает только fixture-backed non-destructive route/permission checks: role redirects, ADMIN Kartei create/edit entry points, OPERATOR pending/declined edit redirects, SUPERADMIN pending overview + War/Ist detail, USER read-only cabinet/detail. Для UI/navigation/template/role/AJAX prompts вне этого scope использовать manual protocol из `docs/BROWSER_QA.md`.

MCP status: репозиторий подтверждает `pytest-playwright` smoke tests и manual browser QA protocol, но не содержит `mcpServers` config и не доказывает наличие Playwright MCP / Chrome DevTools MCP в runtime агента. Если MCP/browser tool недоступен, явно назвать blocker и выполнить доступный manual/automated substitute.

Для manual MCP/browser QA в локальной dev-базе сначала проверить наличие `.codex-local/browser-qa-credentials.md`. Это ignored local file с логинами/паролями ролей `ADMIN`, `OPERATOR`, `SUPERADMIN`, `USER` и base URL. Если файл есть, использовать его для входа в UI, но не печатать пароли в ответах и не коммитить файл. Если файла нет или роль не заполнена, запросить credentials у пользователя или явно отметить blocker. Login form selectors: `#id_username`, `#id_password`, submit button `button[type="submit"]`.

For UI/permission/behavior prompts, check `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md`, and `docs/BROWSER_QA.md` before editing. Future prompt files should use `docs/PROMPT_TEMPLATE.md`.

`manage.py check` и `makemigrations --check --dry-run` — это sanity-check, а не функциональные/регрессионные тесты.

Каждый prompt, который меняет поведение приложения, обязан включать:

1. обновление `docs/Status.md`;
2. либо автоматический тест, либо явный manual QA checklist;
3. sanity-check:

```bash
python backend/manage.py check
python backend/manage.py makemigrations --check --dry-run
pytest -m "not browser" --reuse-db
```

Если тесты добавить нельзя, prompt должен прямо объяснить почему и дать ручной сценарий проверки.
