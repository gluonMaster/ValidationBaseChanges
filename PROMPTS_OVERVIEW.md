# Prompts Overview - KindEltern Web

Этот файл дает краткое и последовательное представление о пайплайне разработки через `PROMPT_XX_*.md` (1 промпт = 1 атомарная правка).

Примечание по контексту: не копипастить весь файл в новый чат. Обычно достаточно открыть нужный этап/номер и работать точечно.

Дата актуализации: 2026-05-27.

Текущий источник правды по статусу промптов: `CLAUDE.md` + `docs/Status.md`. Этот файл дает обзор pipeline, но не должен использоваться в отрыве от этих двух файлов.

Считать примененными как development/code-chain:

- `PROMPT_01...PROMPT_157`, кроме `PROMPT_131`;
- `PROMPT_159`;
- `PROMPT_161`.

Считать не примененными как development/code-chain:

- `PROMPT_131` — optional;
- `PROMPT_158`;
- `PROMPT_160`;
- `PROMPT_162` — **не применен**, код все еще в старом состоянии по operator group navigation;
- `PROMPT_179` — applied pre-stabilization test harness baseline;
- `PROMPT_180` — applied pre-stabilization docs baseline: QA matrix, UI route map, prompt template;
- `PROMPT_181` — applied pre-stabilization app README/comment cleanup, no behavior changes;
- `PROMPT_182` — applied pre-stabilization PRICELIST V2 regression/xfail cases;
- `PROMPT_183` — applied pre-stabilization browser/UI QA protocol and Playwright feasibility, docs-only;
- `PROMPT_184` — applied pre-stabilization GPT Pro context archive refresh;
- `PROMPT_185` — applied pre-stabilization automated browser smoke tests baseline;
- `PROMPT_166...PROMPT_178` — implementation plan для `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`, на дату актуализации не считать примененным к Django-коду.

Отдельно: `PROMPT_163...PROMPT_165` были использованы как GPT Pro analysis/review/spec prompts. Они не являются частью development chain `162, 166-178`, не являются Django code-change prompts и могут быть удалены позже как временный исторический контекст.

Примечание по файлам: архив промптов лежит в `Alt-PROMPTS-ValidationBaseChange/` (в основном `PROMPT_01...PROMPT_130`). Исторически некоторые промпты могут оставаться в `Alt-PROMPTS-ValidationBaseChange/pending-PROMPTS/` даже после применения, поэтому ориентироваться нужно на статусы в этом файле (`Applied`/`Pending`). В корне репозитория остаётся handoff-файл `PROMPT_00_NEXT_CHAT_HANDOFF.md`, а также сюда же создаются новые промпты для последующих code-change правок.

## Общие правила пайплайна

- После каждого промпта: короткий sanity-check и фиксация результата.
- Архитектуру не перепридумывать: менять точечно и минимально.
- UI-тексты/сообщения: немецкий. Комментарии в коде: английский.
- Следующий номер для новых code-change промптов: выбирать как `max(PROMPT_XX)+1`.
- Каждый prompt, который меняет поведение приложения, обязан обновить `docs/Status.md` и добавить автоматический тест или явный manual QA checklist.

Текущее QA-состояние (2026-05-27): после применения `PROMPT_179` в репозитории есть минимальная backend test harness база (`pytest`, `pytest-django`, `pytest.ini`, `backend/tests/`). После применения `PROMPT_180` есть `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md`, `docs/PROMPT_TEMPLATE.md`; будущие prompts должны использовать эти документы для QA/route/permission контекста. После применения `PROMPT_181` обновлены только app-level README/comments вокруг approvals/Kartei edit flow; behavior, executable code и UI-тексты не менялись. После применения `PROMPT_182` добавлен `backend/tests/test_pricelist_v2_regression_cases.py` с passing regression cases и explicit `xfail` target gaps для PRICELIST V2. После применения `PROMPT_183` добавлен `docs/BROWSER_QA.md`, обновлены `docs/QA_MATRIX.md` и `docs/PROMPT_TEMPLATE.md`. После применения `PROMPT_184` и post-185 refresh текущий selected GPT Pro context archive для аудита — `GPT_PRO_POST_179_185_REPO_OPTIMIZATION_CONTEXT.zip` с `backend/apps/familyid_reservations/`, QA/browser docs, backend tests и browser smoke baseline. После применения `PROMPT_185` добавлены `pytest-playwright`, `backend/tests/browser/` и минимальный non-destructive browser smoke baseline через `pytest-django live_server`; `manage.py check` и `makemigrations --check --dry-run` остаются sanity-check, а browser smoke покрывает только явно перечисленные role/route сценарии.

Рекомендуемые команды проверки:

- `python backend/manage.py check`
- `python backend/manage.py makemigrations --check --dry-run`
- `pytest -m "not browser" --reuse-db`
- `python -m playwright install chromium`
- `pytest -m browser backend/tests/browser --browser chromium --reuse-db`
- Не запускать backend и browser pytest параллельно против одной PostgreSQL test DB; при stale `test_kindeltern` использовать cleanup из `RUNNING_LOCAL_QUICKSTART.md`.
- (опционально) компиляция шаблонов:
  - `python backend/manage.py shell -c "from pathlib import Path; from django.template.loader import get_template; root=(Path('templates') if Path('templates').exists() else Path('backend/templates')); [get_template(p.relative_to(root).as_posix()) for p in root.rglob('*.html')]; print('OK')"`

## Applied (132-134)

- `PROMPT_132_CATALOG_SYNC_FROM_LEGACY_SUBJECTS_TEACHERS_ASSIGNMENTS.md` - `/catalog/`: инструмент для добавления недостающих Fächer/Lehrer/Zuweisungen из легаси-записей года.
- `PROMPT_133_FIX_CATALOG_SYNC_ASSIGNMENTS_PREVIEW_SHOW_SUBJECT_NAMES.md` - FIX: в preview синка Zuweisungen показывать человекочитаемые названия предметов (не `casefold`).
- `PROMPT_134_RECORD_EDIT_AUTO_MODE_DO_NOT_OVERWRITE_PRICE_HISTORY.md` - FIX: в AUTO режиме не перезатирать `base_amounts`/месяцы при сохранении без биллинговых изменений (сохранять “историю цен”).

## Applied (135-142.3) — Pricelist v2 (итерация 1)

- `PROMPT_135_SEMESTER_CONFIG_MODEL.md` - SemesterConfig (граница семестра по году) + использование в billing.
- `PROMPT_136_SUBJECT_CATEGORY_MODELS.md` - модели категорий дисциплин: SubjectCategory, SubjectCategoryLink.
- `PROMPT_137_CONTRACT_ENTRY_MODELS.md` - модели помесячной истории контракта: ContractTypeEntry, ContractStatusEntry (+ helpers).
- `PROMPT_138.1_CATEGORIES_DEFAULT_SERVICE.md` - сервис: дефолтные категории года (ensure_default_categories).
- `PROMPT_138.2_CATEGORIES_FORM.md` - форма категорий SubjectCategoryForm.
- `PROMPT_138.3_CATEGORIES_VIEWS_URLS.md` - CRUD views+urls категорий.
- `PROMPT_138.4_CATEGORIES_TEMPLATES_AND_INDEX.md` - templates категорий + ссылка на `/catalog/`.
- `PROMPT_139_DISCIPLINE_GROUP_MODELS.md` - модели групп: DisciplineGroup, DurationEntry, GroupSizeEntry (+ helpers).
- `PROMPT_140.1_CATEGORY_SUBJECTS_VIEWS_URLS.md` - views+urls привязки/отвязки Subject ↔ Category.
- `PROMPT_140.2_CATEGORY_SUBJECTS_TEMPLATES_AND_LINKS.md` - template управления привязками + ссылки из списка категорий.
- `PROMPT_141.1_GROUP_FORMS.md` - формы: DurationEntryForm, GroupSizeEntryForm.
- `PROMPT_141.2_GROUP_VIEWS_URLS.md` - views+urls групп (list/detail, add entries, toggle scaling).
- `PROMPT_141.3_GROUP_TEMPLATES_AND_NAV.md` - templates групп + навигация.
- `PROMPT_142.1_GROUP_SIZE_SERVICE.md` - сервис авто-размера группы.
- `PROMPT_142.2_GROUP_SIZE_INTEGRATE_GROUP_DETAIL.md` - интеграция auto/manual size в group_detail.
- `PROMPT_142.3_GROUP_SIZE_API_ENDPOINT.md` - (опционально) API endpoint размера группы.

## Applied (143.1-151.5) — Pricelist v2 (итерация 2)

- `PROMPT_143.1_PRICING_MODULE.md` - pricing engine (suggested prices, CATEGORY vs PRICE_OPTION).
- `PROMPT_143.2_PRICING_INTEGRATE_GROUP_DETAIL.md` - suggested price на странице группы.
- `PROMPT_144.1_BILLING_CONTRACT_STATUS_ZEROING.md` - billing: PAUSED/TERMINATED → 0.00.
- `PROMPT_144.2_MONTHS_OVERRIDE_ENFORCE_CONTRACT_STATUS.md` - months-override: PAUSED/TERMINATED read-only = 0.00.
- `PROMPT_145.1...145.4` - category price apply service/view/UI/preview.
- `PROMPT_146.1...146.2` - bulk apply на группе: preview + apply.
- `PROMPT_147.1...147.4` - pricing контекст, badges/mismatch, readonly price fields, breakdown.
- `PROMPT_148.1...148.3` - contract timeline UI/change views/approval metadata.
- `PROMPT_149.1...149.3` - warnings service/UI и quick-set subject refs.
- `PROMPT_150.1...150.3` - копирование категорий между годами.
- `PROMPT_151.1...151.5` - динамические semester ranges без хардкода 1-6/7-12.

## Applied / Not Applied (152-162) — fixes after pricelist v2 iteration 2

| Промпт | Статус |
| --- | --- |
| `PROMPT_152` | Applied |
| `PROMPT_153` | Applied |
| `PROMPT_154` | Applied |
| `PROMPT_155` | Applied |
| `PROMPT_156` | Applied |
| `PROMPT_157` | Applied |
| `PROMPT_158` | Not applied |
| `PROMPT_159` | Applied |
| `PROMPT_160` | Not applied |
| `PROMPT_161` | Applied |
| `PROMPT_162` | Not applied |

Детали по этим статусам: `PRICELIST_LOGIC.md`.

## Historical / temporary (163-165) — GPT Pro review/spec phase, not development chain

- `PROMPT_163_PRICELIST_V2_STABILIZATION_FOR_GPT_PRO.md` - инженерный аудит и стратегия стабилизации.
- `PROMPT_164_PRICELIST_V2_SPEC_REVIEW_FOR_GPT_PRO.md` - review промежуточного spec.
- `PROMPT_165_FINAL_REVISED_TZ_REVIEW_FOR_GPT_PRO.md` - финальное review revised TZ.

Эти промпты были использованы для делегирования анализа в GPT Pro. Они не должны использоваться как маркер состояния Django-кода и не входят в implementation chain. Если они будут удалены, статус разработки не меняется.

## Planned / Not Applied (166-178) — stabilization implementation plan

`PROMPT_166...PROMPT_178` — последовательный implementation plan для `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`. На дату 2026-05-27 эти prompts не считать примененными к Django-коду.

Канон для этой фазы:

- `PRICELIST_V2_STATE_AUDIT.md` — current broken state.
- `PRICELIST_V2_STABILIZATION_TZ_REVISED.md` — target state / master-spec.
- `PROMPT_166...PROMPT_178` — planned implementation steps.

## Applied / Not Applied (179-185) — pre-stabilization prompts before 166

Эти prompt-файлы созданы после аудита документации и должны быть применены **до** `PROMPT_166`. Нумерация 179-185 отражает порядок создания файлов, а не то, что они должны идти после 166-178 по смыслу реализации.

Рекомендуемый порядок:

1. `PROMPT_179_PRE_STABILIZATION_TEST_HARNESS_BASELINE.md` — Applied: минимальная pytest/pytest-django база и первые regression/sanity tests.
2. `PROMPT_180_DOCS_QA_MATRIX_UI_ROUTE_MAP_PROMPT_TEMPLATE.md` — Applied: `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md`, `docs/PROMPT_TEMPLATE.md`.
3. `PROMPT_181_APPROVALS_README_AND_KARTEI_EDIT_COMMENT_CLEANUP.md` — Applied: app README и устаревшие comments/docstrings без изменения behavior.
4. `PROMPT_182_PRICELIST_V2_REGRESSION_CASES_BEFORE_REFACTOR.md` — Applied: regression/xfail cases для PRICELIST V2 до refactor.
5. `PROMPT_183_BROWSER_UI_QA_PROTOCOL_AND_PLAYWRIGHT_FEASIBILITY.md` — Applied: browser QA protocol и Playwright feasibility, docs-only.
6. `PROMPT_184_REFRESH_GPT_PRO_CONTEXT_ARCHIVE.md` — Applied: обновлен GPT Pro context archive с `backend/apps/familyid_reservations/`, QA/browser docs и backend test baseline.
7. `PROMPT_185_AUTOMATED_BROWSER_SMOKE_TESTS_BASELINE.md` — Applied: минимальный Playwright browser smoke baseline через `pytest-playwright` + `live_server`.

После 179-185 отдельно решить `PROMPT_162`, затем переходить к `PROMPT_166...PROMPT_178`.

`PROMPT_185` не заменяет обязательный manual browser QA protocol полностью: он покрывает только fixture-backed non-destructive smoke routes, а не destructive approvals, catalog group pricing, AJAX/visual/responsive checks.

## Этап A. База проекта и доменная модель (01-06)

- `PROMPT_01_PROJECT_SETUP.md` - каркас Django-проекта в `backend/`, базовые приложения, Docker Compose (web + Postgres).
- `PROMPT_02_KARTEIEN_MODELS.md` - модель `KarteiRecord` и базовые сущности `karteien` по `DOMAIN_MODEL.md`, миграции.
- `PROMPT_03_APPROVALS_FLOW.md` - модели/сервисы approvals: `PendingChange`, `DeclinedChange`, risk-классификация (без UI).
- `PROMPT_04_HISTORY_AND_REPORTING.md` - `HistoryEvent` + совместимость с `history_raw` (legacy текст), базовое отображение истории.
- `PROMPT_05_NOTIFICATIONS.md` - `Notification` + генерация событий и минимальный UI-индикатор.
- `PROMPT_06_DOCKER_DEPLOY.md` - завершение Docker-настроек и документации по запуску/деплою.

## Этап B. Импорт из Access и базовый UI по ролям (07-10)

- `PROMPT_07_LEGACY_IMPORT.md` - импорт `.accdb` (Access) в Postgres через Django модели + management-команда.
- `PROMPT_07A_LEGACY_IMPORT_PATCH.md` - точечные исправления/дополнения к импорту (patch поверх 07).
- `PROMPT_08_ADMIN_UI_AND_RULES.md` - UI Admin/Operator: list/detail/create/edit `KarteiRecord`, правила валидации, ограничения.
- `PROMPT_08A_ADMIN_UI_PATCH.md` - точечные фиксы UI Admin/Operator (patch поверх 08).
- `PROMPT_09_SUPERADMIN_UI_AND_DECISIONS.md` - UI Superadmin: работа с pending/declined, применение решений.
- `PROMPT_10_USER_HISTORY_UI.md` - read-only UI пользователя: поиск/просмотр записей и истории.

## Этап C. Стабилизация импорта и workflow (11-15)

- `PROMPT_11_FINAL_FIXES_AND_MIGRATIONS.md` - фиксы миграций/urls/entrypoint/deps для стабильного запуска.
- `PROMPT_12_MULTIYEAR_IMPORT_FIX_PRIMARY_KEY.md` - surrogate PK `pkid` для `KarteiRecord` + доменный ключ `(year, id)`.
- `PROMPT_13_FIX_EDIT_BUTTON_FOR_NORMAL_STATUS.md` - `NORMAL` хранится как пустая строка: фиксы проверок/кнопок.
- `PROMPT_14_FIX_DECLINED_EDIT_WORKFLOW_VARIANT_A.md` - DECLINED variant A: редактируется snapshot в `DeclinedChange`, затем resubmit в pending.
- `PROMPT_15_FIX_FAMILYID_PARENT_UNIQUE_EXCLUDE_PKID.md` - корректное exclude в проверках уникальности после ввода `pkid`.

## Этап D. Catalog + прайс + скидки + AUTO начисления (16-23)

- `PROMPT_16_CATALOG_MODELS_TEACHERS_SUBJECTS_ASSIGNMENTS.md` - модели справочников Teacher/Subject и назначения по годам.
- `PROMPT_17_CATALOG_ADMIN_UI_AND_COPY_YEAR.md` - UI справочников + копирование настроек "год -> год".
- `PROMPT_18_PRICE_LIST_MODEL_AND_ADMIN_UI.md` - прайс-лист (год+дисциплина), спец-логика Ind/VSpE\_/NH, UI и копирование.
- `PROMPT_19_DISCOUNTS_MODELS_AND_ADMIN_UI.md` - модели скидок и назначения (семья/запись) с привязкой к месяцам.
- `PROMPT_20_INTEGRATE_CATALOG_SELECTION_IN_KARTEI_FORMS.md` - интеграция выбора справочников/прайса в формы karteien.
- `PROMPT_21_AUTO_MONTH_CALC_DISCOUNTS_LOCKING_AND_ROUNDING.md` - автоначисления: семестры, Ind/NH (hours), скидки, "с месяца", запрет ручного ввода, округление вверх.
- `PROMPT_22_DISCOUNT_WORKFLOW_FAMILY_DASHBOARD_AND_APPLY.md` - Family Dashboard + массовое применение семейных скидок через approvals.
- `PROMPT_23_MONTHLY_CHARGE_BREAKDOWN_UI.md` - UI детализации начислений по месяцам (Admin/User), включая Ind/NH.

## Этап E. UX улучшения и точечные post-fix (24-29)

- `PROMPT_24_KARTEI_LIST_LIVE_SEARCH_AJAX.md` - live-search на `/karteien/` без перезагрузки.
- `PROMPT_25_CREATE_RECORD_AUTOCOMPLETE_AND_PREFILL.md` - `/karteien/create/`: autocomplete Eltern/Kind + prefill + fallback по прошлым годам.
- `PROMPT_26_FIX_MONTHS_OVERRIDE_TEMPLATE_ATTR_FILTER.md` - фиксы компиляции `months_override.html` (убран несуществующий filter).
- `PROMPT_27_FIX_ZERO_CLAMP_CONFIRMATION_FLOW_CREATE_AND_EDIT.md` - подтверждение clamp (negative -> 0) на create/edit.
- `PROMPT_28_FIX_PERCENT_DISCOUNT_DISPLAY_AS_PERCENT.md` - процентные скидки хранятся долей (0.25), отображаются как 25%.
- `PROMPT_29_ADJUST_ACADEMIC_HOURS_STEP_TO_0_01.md` - ввод academic hours с шагом 0.01 + нормализация до 2 знаков.

## Этап F. Скидки: ускорение работы админа (31-33)

- `PROMPT_31_FIX_KARTEI_PKID_DISPLAY_AND_LABELS.md` - устранение путаницы `pkid` vs Access `id` в UI.
- `PROMPT_32_ADD_RECORD_DISCOUNT_SHORTCUTS_ON_KARTEI_CREATE_EDIT.md` - shortcuts для назначения скидок прямо из create/edit записи.
- `PROMPT_33_RECORD_DISCOUNT_RECORD_PICKER_AUTOCOMPLETE.md` - autocomplete-подбор записи для Eintragsrabatte (вместо ручного ввода PKID).

## Этап G. Legacy маркеры контракта + patch-импорт (34-35, 38)

- `PROMPT_34_IMPORT_PATCH_CONTRACT_MARKERS_TEACHERS_AND_SEPA.md` - доимпорт legacy полей из `tblKartei` и `--patch-fields` режим (teachers, contract markers, SEPA).
- `PROMPT_35_UI_CONTRACT_TYPE_STATUS_AND_FILTERS_ALL_ROLES.md` - отображение/фильтры: месячный (O/V) и расторгнутый (KN) по ролям.
- `PROMPT_38_ADMIN_EDIT_CONTRACT_FIELDS_WITH_RAW_MARKERS.md` - редактирование админом raw-маркеров и вычисляемых флагов.

## Этап H. Доступы по ролям + мастер "Neue Familie" (36-37, 39)

- `PROMPT_36_SUPERADMIN_READONLY_KARTEIEN_LIST_DETAIL.md` - read-only доступ Superadmin к `/karteien/` (list/detail/live search) без edit/delete/create.
- `PROMPT_37_NEW_FAMILY_WIZARD_CREATE_MULTI_CHILD_WITH_DISCOUNTS.md` - мастер создания новой семьи с несколькими детьми, скидками и пересчетом AUTO.
- `PROMPT_39_ENABLE_USER_READONLY_KARTEIEN_LIST_DETAIL_LIVESEARCH.md` - read-only доступ User к `/karteien/` (list/detail/live search).

## Этап I. Доработки мастера "Neue Familie" (41-45)

- `PROMPT_41_NEUE_FAMILIE_FIX_FAMILYID_AUTOSUGGEST.md` - автопредложение следующего `FamilyID` от глобального максимума (по всем годам).
- `PROMPT_42_NEUE_FAMILIE_SEPA_MARKER_AS_CHOICE.md` - `SEPA-Marker` как выбор/опция (не ручной ввод).
- `PROMPT_43_NEUE_FAMILIE_REMOVE_CONTRACT_TERMINATED.md` - убрать опцию "Vertrag gekündigt" при создании новой семьи.
- `PROMPT_44_NEUE_FAMILIE_FILTER_TEACHER_AND_PRICE_BY_SUBJECT.md` - фильтрация `Lehrer` и `Preis` по выбранному `Fach` (оба полугодия).
- `PROMPT_45_NEUE_FAMILIE_EINTRAGRABATTE_UI_AND_DEDUP_WITH_FAMILY.md` - улучшения UI Eintragsrabatte + дедупликация с семейной скидкой.

## Этап J. Approval/notifications (46-47, 50)

- `PROMPT_46_SUPERADMIN_PENDING_NOTIFICATIONS_AND_APPROVAL_LAUNCH.md` - уведомления и точки входа к pending для Superadmin.
- `PROMPT_47_FIX_PENDING_NOTIFICATION_LINK.md` - корректный возврат по ссылке уведомления.
- `PROMPT_50_SUPERADMIN_WAR_IST_ENTRYPOINTS.md` - восстановление/добавление входов на War/Ist и навигации по pending.

## Этап K. Редактирование записи (48-49, 51-59)

- `PROMPT_48_RECORD_EDIT_SEPA_MARKER_AS_CHOICE.md` - `SEPA-Marker` как выбор/опция в edit.
- `PROMPT_49_RECORD_EDIT_FILTER_TEACHER_PRICE_FOR_LEGACY.md` - фильтрация `Lehrer/Preis` по `Fach` с поддержкой legacy.
- `PROMPT_51_RECORD_EDIT_LIVE_PREVIEW_AND_EXPLAIN_MONTHS.md` - live-preview начислений + пояснение по клику на месяц.
- `PROMPT_52_RECORD_EDIT_LEGACY_TO_AUTO_ON_FIRST_MEANINGFUL_CHANGE.md` - перевод legacy-записей в `months_mode=AUTO` при первом "смысловом" изменении.
- `PROMPT_53_RECORD_DISCOUNTS_RETURN_TO_EDIT_AND_REFRESH_PREVIEW.md` - возврат из каталога скидок в edit + auto-refresh preview.
- `PROMPT_54_DISCOUNTS_DISABLED_BY_MONTHS.md` - `Rabatte deaktiviert` с выбором месяцев (von/bis + CSV), применение в preview и сохранении.
- `PROMPT_55_CONTRACT_TYPE_AND_STATUS_EFFECTIVE_MONTH.md` - "ab Monat" для контрактных изменений.
- `PROMPT_56_COMMENT_REQUIRED_FOR_ANY_CHANGE.md` - обязательный admin-комментарий при любых изменениях.
- `PROMPT_57_RECORD_EDIT_LIVE_PREVIEW_FIX_CSRF_AND_SHOW_VALUES.md` - фиксы preview (CSRF/отображение значений).
- `PROMPT_58_RECORD_EDIT_BILLING_PREVIEW_APPLY_DISCOUNTS_AND_REASONS.md` - причины начислений в preview + применение скидок.
- `PROMPT_59_RECORD_EDIT_LEGACY_TO_AUTO_TOUCHED_MONTHS_RESPECT_APPLY_FROM.md` - пересчет только затронутых месяцев при legacy->AUTO с уважением `apply_from_month_*`.

## Этап L. Комментарии администратора в approvals/history (60-64)

- `PROMPT_60_KARTEIEN_FORMS_REMOVE_DUPLICATE_CONTRACT_FIELDS.md` - устранение дублей контрактных полей в формах.
- `PROMPT_61_CONTRACT_TERMINATION_EFFECTIVE_MONTH_PERSIST_AND_PREVIEW.md` - хранение и учет "Kündigung ab Monat" в preview/расчете.
- `PROMPT_62_APPROVALS_STORE_ADMIN_COMMENT_IN_PENDING_CHANGE.md` - сохранение admin-комментария в pending.
- `PROMPT_63_HISTORY_APPEND_ADMIN_COMMENT_ON_SAFE_SAVE.md` - запись admin-комментария в историю при safe-save.
- `PROMPT_64_HISTORY_INCLUDE_ADMIN_COMMENT_ON_APPROVE_DECLINE.md` - перенос комментария в историю при approve/decline.

## Этап M. Legacy-UX и быстрые сценарии каталогов (65-72)

- `PROMPT_65_RECORD_EDIT_FIX_MONTH_CSS_RENDERED_AS_TEXT.md` - исправить CSS, который рендерился как текст.
- `PROMPT_66_LEGACY_IGNORE_REF_LINK_ONLY_CHANGES_IN_MEANINGFUL_DETECTION.md` - не считать ref-link-only изменения "смысловыми" для legacy->AUTO.
- `PROMPT_67_CATALOG_SUBJECT_CREATE_PREFILL_AND_NEXT.md` - create Subject: prefill + `next`.
- `PROMPT_68_CATALOG_TEACHER_CREATE_PREFILL_FULL_NAME_AND_NEXT.md` - create Teacher: prefill full name + `next`.
- `PROMPT_69_CATALOG_ASSIGNMENT_CREATE_PREFILL_AND_NEXT.md` - create TeachingAssignment: prefill + `next`.
- `PROMPT_70_CATALOG_PRICE_CREATE_QUICKADD_REQUIRE_COMMENT_DUPLICATES_NEXT.md` - create PriceOption: quick-add, обязательный комментарий, обработка дублей, `next`.
- `PROMPT_71_KARTEI_PREFILL_TEACHER_REF_FROM_TEACHER_LEGACY_NAME.md` - prefill `teacher*_ref` по `teacher*_legacy_name` (если возможно).
- `PROMPT_72_RECORD_EDIT_LEGACY_BADGES_AND_QUICK_LINKS_TO_CATALOG.md` - бейджи под `Fach/Lehrer/Preis` при несоответствии каталогам + быстрые ссылки.

## Этап N. Месяцы: подсветка несовпадений/режимы/диапазон цены/preview (73-79)

- `PROMPT_73_LEGACY_MONTHS_HIGHLIGHT_MISMATCH_WITH_PRICE_RED_BORDER.md` - подсветка подозрительных начислений (кроме Stundenfächer) + бейдж `Override`.
- `PROMPT_74_RECORD_EDIT_MONTHS_READONLY_HIDE_SPINNERS_AND_MOVE_OVERRIDE_BUTTON.md` - месяцы в edit read-only по умолчанию, скрыть спиннеры, переместить кнопку Override.
- `PROMPT_75_RECORD_EDIT_LEGACY_ADD_RECALCULATE_MONTHS_BUTTON_AND_LINK_OVERRIDE.md` - добавить кнопку пересчета месяцев для legacy и ссылку на `months-override`.
- `PROMPT_76_RECORD_EDIT_PRICE_CHANGE_PANEL_SHOW_ONLY_WHEN_PRICE_CHANGED.md` - показывать панель применения изменения цены только когда цена действительно изменена.
- `PROMPT_77_RECORD_EDIT_PRICE_CHANGE_ADD_END_MONTH_RANGE.md` - добавить `bis Monat` (диапазон) для применения изменения цены.
- `PROMPT_78_RECORD_EDIT_LEGACY_ADD_MONATE_NEU_BERECHNEN_BUTTON.md` - добавить кнопку **Monate neu berechnen** (перевод legacy -> auto + пересчет).
- `PROMPT_79_BILLING_PREVIEW_AUTO_ALWAYS_CALCULATE_ALL_MONTHS.md` - preview в `AUTO` всегда пересчитывает все месяцы (не только "затронутые").

## Этап O. Pending UX, approvals edit и уведомления (80-84)

- `PROMPT_80_RECORD_EDIT_FORCE_RECALC_ALLOW_PRICE_CHANGE_WITHOUT_APPLY_FROM.md` - legacy: `Monate neu berechnen` не требует `Preis anwenden ab Monat` при смене цены (forced recalc).
- `PROMPT_81_MONTHS_OVERRIDE_PERSIST_REASON_TO_PENDING_ADMIN_COMMENT.md` - Override: причина сохраняется в `PendingChange.admin_comment` и видна Superadmin.
- `PROMPT_82_RECORD_DETAIL_PENDING_SHOW_PENDING_VALUES_WITH_BADGES_AND_TOGGLE.md` - record detail: ADMIN видит pending snapshot по умолчанию + бейджи `Wartend` + переключатель Wartend/Aktuell.
- `PROMPT_83_ADMIN_CAN_EDIT_PENDING_CHANGE_UPDATE_SNAPSHOT_AND_ADD_OPTIMISTIC_CONFLICT_GUARD.md` - ADMIN редактирует `PendingChange.snapshot` + optimistic guard по `updated_at`.
- `PROMPT_84_NOTIFICATIONS_REALTIME_PENDING_AND_FIX_APPROVED_TYPE_UI.md` - колокольчик: корректные типы + pending живёт до решения + частый polling.

## Этап O1. Фиксы по результатам ревью (85-86)

- `PROMPT_85_SUPERADMIN_WAR_IST_OPTIMISTIC_GUARD_FIX_UPDATED_AT_FORMAT.md` - War/Ist: поправить формат `pending_updated_at`, чтобы optimistic guard не был ложноположительным.
- `PROMPT_86_NOTIFICATIONS_DROPDOWN_SHOW_ONLY_UNREAD_AND_HIDE_RESOLVED_PENDING.md` - колокольчик: показывать только unread, чтобы pending реально исчезал после решения.

## Этап P. Pending/Declined UX и уведомления (87-89)

- `PROMPT_87_RECORD_EDIT_TEACHER_LEGACY_ALWAYS_SHOW_ASSIGNMENT_LINK.md` - `/edit/`: для legacy Lehrer всегда показывать `Zuweisung erstellen` при выбранном Fach (даже без teacher-candidate).
- `PROMPT_88_NOTIFICATIONS_MARK_PENDING_READ_AFTER_DECISION.md` - уведомления: после approve/decline помечать `PENDING_CREATED` как read (очистка pending у Superadmin).
- `PROMPT_89_PENDING_DECLINED_USE_STANDARD_RECORD_EDITOR_INSTEAD_OF_SNAPSHOT_ONLY.md` - PENDING/DECLINED: редактирование через стандартный редактор записи с prefill из snapshot и сохранением safe-полей.

## Этап P1. Фиксы по результатам ревью (90-91, применены)

- `PROMPT_90_DECLINED_OVERVIEW_EDIT_BUTTON_USE_STANDARD_RECORD_EDITOR.md` - declined-overview: кнопка редактирования ведёт в стандартный редактор (а не snapshot-only).
- `PROMPT_91_NOTIFICATIONS_FIX_PENDING_CLEANUP_PKID_AND_AUTO_CLEAR_STALE.md` - колокольчик SUPERADMIN: исправить очистку pending (pkid vs id) + авто-очистка «протухших» pending.

## Этап Q. Биллинг-месяцы и видимость скидок (92-99, применены)

- `PROMPT_92_BILLING_PREVIEW_LEGACY_TOUCHED_MONTHS_DEFAULTS.md` - preview для LEGACY: убрать ложное требование `Preis anwenden ab Monat`, выровнять с save-логикой `touched_months`.
- `PROMPT_93_BILLING_MONTHS_END_MONTH_PERSIST.md` - добавить сохраняемый `Endmonat (bis)` начислений для каждого семестра.
- `PROMPT_94_BILLING_MONTHS_CSV_PERSIST.md` - добавить `Monate (CSV)` для точечного выбора месяцев начислений (поверх Start/End).
- `PROMPT_95_RECORD_DETAIL_AND_EDIT_SHOW_DISCOUNTS_SUMMARY.md` - `/record_detail` и `/edit`: показывать полный список применённых скидок (Family + Eintrag).
- `PROMPT_96_KARTEI_LIST_SHOW_DISCOUNT_BADGES.md` - `/karteien/`: краткие бейджи скидок в таблице без N+1 (Exists/Count).
- `PROMPT_97_DISCOUNTS_LISTS_ADD_PARENT_NAMES.md` - Eintragsrabatte/Familienrabatte: добавить колонку `Eltern` в списки.
- `PROMPT_98_FAMILY_APPLY_DISCOUNTS_INCLUDE_LEGACY_OPTION.md` - Family: предупреждать про LEGACY + опционально применять скидки к LEGACY на базе текущих Monatswerte.
- `PROMPT_99_LEGACY_RECALCULATED_BADGE_AND_MONTH_EXPLANATION.md` - LEGACY-recalc: бейдж и объяснение по клику, что база legacy и нет привязки к цене.

## Этап R. NeuList/History/Kartei UX + create-prefill (100-112, применены)

- `PROMPT_100_PENDING_SAFE_FIELDS_PERSIST_BILLING_MONTH_SETTINGS_AND_LEGACY_RECALC_FLAG.md` - PENDING (risky-save): сохранять SAFE-поля `end_month_*`, `months_csv_*`, `legacy_base_amounts_enabled`, чтобы настройки месяцев и LEGACY-recalc маркер не терялись.
- `PROMPT_101_APPROVALS_NEULIST_FIX_NEW_RECORDS_PER_YEAR.md` - NeuList: исправить «новые записи» с учётом доменного ключа `(year, id)` (перейти на per-year last_seen).
- `PROMPT_102_HISTORY_PARSE_AND_RENDER_IN_RECORD_DETAIL_AND_USER_VIEWS.md` - Historie: показывать таймлайн (не сырой `history_raw`) и не терять comment-only события (APR/ADM), корректно парсить дату DCL.
- `PROMPT_103_KARTEI_LIST_CONTRACT_STATUS_ADD_SEPA_VARIANTS.md` - `/karteien/`: добавить `Aktiv-SEPA` и `Gekündigt-SEPA` в Vertragsstatus (и в list view, и в live-search API).
- `PROMPT_104_KARTEI_LIST_FILTER_SUBJECT_TEACHER_SEMESTER_LIVESEARCH.md` - `/karteien/`: динамические фильтры Unterricht/Lehrer по семестру (через существующий AJAX live-search).
- `PROMPT_105_KARTEI_LIST_COLLAPSIBLE_FILTER_PANEL.md` - `/karteien/`: секция Filter сворачиваемая (collapsed по умолчанию), без ломания live-search.
- `PROMPT_106_CREATE_FAMILYID_HELP_TOOLTIP_AND_FREETEXT_WARNING.md` - `/api/karteien/create/`: подсказка возле FamilyID (max + рекомендация) и предупреждение для Freitext-секции.
- `PROMPT_107_FAMILY_DASHBOARD_ADD_CREATE_RECORD_LINK_AND_BASIC_PREFILL.md` - Family Dashboard: entry-point «Neuer Datensatz» + базовый prefill полей из семьи.
- `PROMPT_108_CREATE_PREFILL_FROM_FAMILY_MULTIVALUE_DROPDOWNS.md` - Create: «умный» prefill из семьи — dropdown выбора значения при нескольких вариантах.
- `PROMPT_109_APPROVALS_NEULIST_PER_YEAR_FALLBACK_FIX_EXISTING_SUPERADMIN_STATES.md` - NeuList: исправить fallback для старых SuperadminState (не наследовать legacy last_seen_id для нового года).
- `PROMPT_110_KARTEI_CREATE_BOOTSTRAP_FALLBACK_FIX_TOOLTIP_AND_RESTORE_AUTOCOMPLETE.md` - Create: bootstrap JS fallback + защита от `bootstrap is not defined`, восстановить autocomplete Eltern/Kind и dropdown prefill, подсказка FamilyID.
- `PROMPT_111_KARTEI_RABATTE_BADGE_SHOW_CORRECT_PERCENT_OR_FIXED_EUR.md` - Rabatte: корректно показывать проценты/€ (PERCENT хранится как доля 0.25=25%).
- `PROMPT_112_KARTEI_CREATE_FAMILYID_HELP_USE_REAL_BOOTSTRAP_POPOVER_WITH_LAZY_INIT.md` - Create: FamilyID help — lazy-init popover по клику (чтобы popover работал при загрузке bootstrap после inline-скрипта).

## Этап S. Superadmin отчёты/UX Kartei + резервация FamilyID (113-130, применены)

- `PROMPT_113_SUPERADMIN_FAMILY_KOSTEN_SEARCH_API.md` - API поиска семей по фамилии родителя (AJAX) для отчёта «Kosten (Familie)».
- `PROMPT_114_SUPERADMIN_KARTEI_LIST_ADD_KOSTEN_REPORT_MODAL.md` - `/karteien/`: модальное окно для поиска семьи и запуска отчёта по FamilyID.
- `PROMPT_115_SUPERADMIN_FAMILY_KOSTEN_REPORT_VIEW.md` - Multi-year отчёт начислений по семье (свод по годам/детям/месяцам).
- `PROMPT_116_SUPERADMIN_FAMILY_KOSTEN_REPORT_BEZAHLT_DIFFERENZ_TOTAL.md` - В отчёте: поле `Bezahlt` по каждому году, расчёт `Differenz` и итог `Noch zu bezahlen`.
- `PROMPT_117_OPTIONAL_OFFCANVAS_KOSTEN_REPORT_ON_KARTEI_LIST.md` - (Опционально) offcanvas-предпросмотр отчёта с подгрузкой фрагмента без ухода со страницы.
- `PROMPT_118_SUPERADMIN_FAMILY_SEARCH_API_EXCLUDE_DECLINED_AND_EMPTY_FAMILYID.md` - Поиск семей: исключить `DECLINED` и записи без `family_id`, уникализировать по `family_id`.
- `PROMPT_119_KARTEI_LIST_YEAR_HEADER_SYNC_AND_SESSION_PERSIST.md` - `/karteien/`: заголовок «Jahr: …» синхронизируется с фильтром и запоминается в сессии.
- `PROMPT_120_KARTEI_LIST_WARN_ON_CREATE_IN_NON_CURRENT_YEAR.md` - Предупреждение при создании записи/семьи в году, отличном от текущего выбранного.
- `PROMPT_121_KARTEI_LIST_FILTER_PANEL_COMPACT_SPACING.md` - Уменьшить вертикальные отступы в секции Filter (компактнее по высоте).
- `PROMPT_122_HISTORY_APPROVE_EVENTS_INCLUDE_FIELD_DIFFS.md` - Historie: для approve/изменений сохранять и показывать diffs (исходное → новое) и комментарии.
- `PROMPT_123_CATALOG_LIVE_REFRESH_FOR_SUBJECT_TEACHER_PRICE_SELECTS.md` - Редактор/создание: динамическое обновление списков Fach/Lehrer/Preis при изменении каталога (без перезагрузки).
- `PROMPT_124_LEGACY_EDIT_SAVE_SHOULD_TRIGGER_CONVERSION_AND_NO_FALSE_WARNINGS.md` - LEGACY edit: сохранение должно инициировать пересчёт/конверсию и не выдавать ложные «нет изменений».
- `PROMPT_125_KARTEI_LIST_SYNC_RESET_FILTERS_LINK_AND_YEAR_WARNING_USE_SYSTEM_YEAR.md` - Доработки `/karteien/`: reset filters + предупреждение про год берёт системный текущий год.
- `PROMPT_126_SUPERADMIN_PENDING_INFO_MENU_REPURPOSE_ALL_YEARS_PENDING.md` - Superadmin меню: repurpose «Pending (Info)» под единый список PENDING по всем годам.
- `PROMPT_127_ADMIN_FAMILYID_RESERVATIONS_MODEL.md` - Модель резерваций FamilyID (номер + дата/время, управление).
- `PROMPT_128_FAMILYID_NEXT_NUMBER_MUST_CONSIDER_RESERVATIONS.md` - Расчёт «следующего FamilyID» учитывает активные резервации.
- `PROMPT_129_ADMIN_RESERVE_NEXT_FAMILYID_AND_LIST_VIEW.md` - UI для админа: зарезервировать следующий FamilyID, просмотр/отмена резерваций.
- `PROMPT_130_NEW_FAMILY_WIZARD_ALLOW_USING_RESERVED_FAMILYID.md` - Мастер «Neue Familie»: выбрать «использовать зарезервированный FamilyID».

## Pending (не применены)

- `PROMPT_131_OPTIONAL_KARTEI_CREATE_FORM_USE_RESERVED_FAMILYID.md` - (Опционально) добавить выбор зарезервированного FamilyID в форме «Kartei Create».

## Примечания по эксплуатации

- `KarteiRecord`:
  - `pkid` - Django PK (surrogate),
  - `id` - Access/Excel ID,
  - доменный ключ: `(year, id)`.
- Для Access импорта: `ACCESS_CONN_STRING_TEMPLATE` использует Python `.format(...)`, поэтому DRIVER должен быть экранирован как `{{...}}`.
- Patch-импорт `--patch-fields` требует примененных миграций к той же базе данных (иначе возможны ошибки `UndefinedColumn ...`).
