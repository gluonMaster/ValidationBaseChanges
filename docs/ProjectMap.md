# Project Map - KindEltern Web (для экономии контекста)

Цель этого файла: дать **короткую карту**, чтобы в новом чате не тащить в контекст большие документы целиком.
Правило чтения: сначала `CLAUDE.md`, затем `docs/Status.md`, затем этот файл как навигационная карта. Для UI/permission/behavior задач также открыть `docs/QA_MATRIX.md`, `docs/UI_ROUTE_MAP.md` и `docs/BROWSER_QA.md`; новые prompt-файлы делать по `docs/PROMPT_TEMPLATE.md`. После этого точечно открывать нужные разделы в `ARCHITECTURE.md` / `DOMAIN_MODEL.md` / `docs/Spec.md`.

## 1) Как экономить контекст (рекомендуемый режим)

- Не копипастить большие блоки архитектуры/домена в чат.
- Не начинать с app-level `backend/apps/*/README.md`: они вторичны и могут содержать historical wording; canonical override — `CLAUDE.md` + `docs/Status.md`.
- Вместо этого:
  - искать точные места через `rg -n "..." backend ...`;
  - открывать файл кусками (первые ~150 строк + нужные участки).
- В решениях не предлагать реархитектуру: только точечные фиксы/UX по фактам.

## 2) Неподвижные инварианты (частые источники багов)

- `KarteiRecord.pkid` = Django PK (surrogate).
- `KarteiRecord.id` = Access/Excel ID (не глобально уникален).
- Доменный ключ записи: `(year, id)`.
- `RecordStatus.NORMAL` хранится как пустая строка `""` → проверка `not record.status`.
- Языки:
  - UI-тексты/сообщения — немецкий,
  - комментарии в коде — английский,
  - PROMPT-файлы — русский.

## 3) Где что лежит (быстрые указатели по коду)

### Kartei (list/detail/edit)

- list + фильтры: `backend/apps/karteien/views.py` (`KarteiRecordListView`)
- live-search API: `backend/apps/karteien/api.py` (`live_search_api`)
- Kostenbericht (Familie): `backend/apps/karteien/kosten_report.py` + маршруты в `backend/apps/karteien/urls.py` (family search API, report view, fragment view)
- UI запуска отчёта: `backend/templates/karteien/record_list.html` (modal/offcanvas для SUPERADMIN)
- расширенные фильтры списка (контракты/SEPA + Unterricht/Lehrer по семестру): `backend/templates/karteien/record_list.html` (JS формирует query params)
- detail: `backend/apps/karteien/views.py` (`KarteiRecordDetailView`)
- edit/create form: `backend/apps/karteien/forms.py` (`KarteiRecordForm`)
- edit/create views: `backend/apps/karteien/views.py` (`KarteiRecordCreateView`, `KarteiRecordUpdateView`)
- шаблоны:
  - list: `backend/templates/karteien/record_list.html`, `backend/templates/karteien/_record_list_table.html`
  - detail: `backend/templates/karteien/record_detail.html`
  - form: `backend/templates/karteien/record_form.html`

### Биллинг (месяцы, скидки, режимы)

- расчёты: `backend/apps/karteien/billing.py`
  - `build_base_amounts(...)`
  - `calculate_month_values(...)`
  - `get_month_breakdown(...)`
  - LEGACY→AUTO: `detect_meaningful_changes(...)`, `recalculate_legacy_to_auto(...)`
  - AUTO: `base_amounts` — “история цен” по месяцам; при сохранении без биллинговых изменений не должен выполняться полный пересчёт (см. PROMPT_134).
- live-preview API: `backend/apps/karteien/api.py` (`billing_preview_api`)
- режимы месяцев: `MonthsMode.LEGACY/AUTO/OVERRIDE` (см. модели `KarteiRecord`)

### Approvals (pending/declined)

- модели: `backend/apps/approvals/models.py`
- логика: `backend/apps/approvals/services.py`
- NeuList per-year: `SuperadminState.last_seen_by_year` + `get_new_records(..., year)` / `update_last_seen_id(..., year)`
- Risk classification: legacy `TRACKED_FIELDS` в `karteien/models.py` + `approvals/services.py`; known defect для PRICELIST V2, потому что non-tracked billing/context fields тоже меняют финансовый смысл.
- `PendingChange.snapshot`: current flat tracked-fields snapshot + reserved metadata keys (`_old_base_amounts`, `_pending_contract_type_entry`, `_pending_contract_status_entry`), не полноценный frozen v2 payload.
- UI Superadmin: `backend/apps/approvals/views.py`, templates `backend/templates/approvals/...`

### Notifications (колокольчик)

- модели/сервисы: `backend/apps/notifications/models.py`, `backend/apps/notifications/services.py`
- API: `backend/apps/notifications/views.py` (DRF)
- фронт (polling/dropdown): `backend/templates/base.html`

### Catalog (предметы/учителя/назначения/цены/скидки)

- `backend/apps/catalog/models.py`, `backend/apps/catalog/views.py`
- templates: `backend/templates/catalog/...`
- инструмент “Aus Legacy ergänzen” (добивка справочников из legacy): `backend/apps/catalog/views.py` (`SyncFromLegacyView`) + `backend/templates/catalog/sync_from_legacy.html`
- FamilyID резервации: `backend/apps/familyid_reservations/` + UI в `backend/templates/catalog/familyid_reservations.html` (вход из `backend/templates/catalog/index.html`)
- **Прайслист v2 (категории/группы):**
  - модели: `backend/apps/catalog/models.py` (`SemesterConfig`, `SubjectCategory`, `SubjectCategoryLink`, `DisciplineGroup`, `DurationEntry`, `GroupSizeEntry`)
  - bootstrap дефолтных категорий: `backend/apps/catalog/services.py` (`ensure_default_categories`)
  - авторазмер группы: `backend/apps/catalog/group_size_service.py` (`calculate_auto_group_size`, `get_group_size_for_month`)
  - pricing suggestion: `backend/apps/catalog/pricing.py`; apply path: `backend/apps/karteien/category_price.py`; legacy preview path: `backend/apps/karteien/api.py` (`billing_preview_api`) + `backend/apps/karteien/billing.py` — это current split-brain, а не target pipeline.
  - UI:
    - категории: `backend/apps/catalog/views.py` (`SubjectCategory*`) + `backend/templates/catalog/category_*.html`
    - привязки Fach↔Kategorie: `backend/apps/catalog/views.py` (`SubjectCategoryLinksView`, `SubjectCategoryUnlinkView`) + `backend/templates/catalog/category_subjects.html`
    - группы: `backend/apps/catalog/views.py` (`DisciplineGroupListView`, `DisciplineGroupDetailView`, `DurationEntryCreateView`, `GroupSizeEntryCreateView`) + `backend/templates/catalog/group_*.html`
  - API: `backend/apps/catalog/views.py` (`GroupSizeApiView`) + URL `catalog:group_size_api`
  - Operator group navigation/UI: `PROMPT_162` не применен; list/index/navbar остаются в старом/admin-only состоянии, хотя часть detail/bulk backend paths уже использует editor-доступ после `PROMPT_161`.

### Family Dashboard (семья, семейные скидки)

- view: `backend/apps/karteien/families.py` (`FamilyDashboardView`, `ApplyDiscountsView`)
- template: `backend/templates/karteien/family_dashboard.html`
- entry-point на создание записи для семьи: `karteien:record_create` с `prefill_family_id=...`

### History (история изменений)

- парсер raw-истории: `backend/apps/history/services.py` (`parse_raw_history`, `sync_history_from_raw`)
- UI таймлайна: `backend/templates/approvals/superadmin/record_history.html`, `backend/templates/accounts/user_record_history.html` (и в `backend/templates/karteien/record_detail.html`)

### Import из Access

- `backend/apps/legacy_import/...` (management команды, маппинги)
- ключевой нюанс: `ACCESS_CONN_STRING_TEMPLATE` собирается через `.format(file_path=...)` → DRIVER в `{{...}}`

## 4) Куда смотреть по типу задачи

- “почему статус Normal не работает” → `RecordStatus.NORMAL == ""` и UI/фильтры
- “pending/declined не сходится” → `approvals/services.py` + `notifications/services.py` + `templates/base.html`
- “месяцы/скидки/preview странно себя ведут” → `karteien/billing.py` + `karteien/api.py (billing_preview_api)` + `record_form.html`
- “скидки не видны/непонятны” → `record_detail.html` + `record_form.html` + `catalog/*discount*`
- “operator не видит группы/нет навигации к группам” → `PROMPT_162` не применен; смотреть `catalog/views.py`, `templates/base.html`, `templates/catalog/index.html`, `templates/catalog/group_*`.
- “почему approval/risk пропустил billing/context change” → legacy `TRACKED_FIELDS` недостаточен; смотреть `approvals/services.py`, `karteien/models.py`, `PRICELIST_V2_STATE_AUDIT.md`.

## 5) Канонические документы (читать точечно)

- Архитектура и модули: `ARCHITECTURE.md`
- Маппинг legacy Access/Excel: `DOMAIN_MODEL.md`
- Инварианты/контракты: `docs/Spec.md`
- Решения (ADR-lite): `docs/Decisions.md`
- Ограничения процесса: `docs/Constraints.md`
- Текущее состояние: `docs/Status.md`
- QA matrix: `docs/QA_MATRIX.md`
- UI route/permission map: `docs/UI_ROUTE_MAP.md`
- Browser/manual QA protocol and smoke scope: `docs/BROWSER_QA.md`
- Future prompt template: `docs/PROMPT_TEMPLATE.md`
- Prompt-chain overview: `PROMPTS_OVERVIEW.md`
- PRICELIST V2 current broken state: `PRICELIST_V2_STATE_AUDIT.md`
- PRICELIST V2 target state: `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`
