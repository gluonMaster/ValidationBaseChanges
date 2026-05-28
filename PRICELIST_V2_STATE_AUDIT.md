# PRICELIST V2 — Аудит текущего состояния

## Цель документа

Этот файл нужен как компактный, но технически точный снимок текущего состояния прайслиста v2 и связанной логики начислений. Он предназначен для нового чата GPT PRO и должен помочь избежать двух типичных ошибок:

1. начать “чинить по месту”, не заметив конфликтующих источников истины;
2. переписать всё сразу, не сохранив важные инварианты домена.

## Что задумано по документам

По `PRICELIST_LOGIC.md`, `TZ_PRICELIST_V2.md` и `ARCHITECTURE.md` текущая целевая модель выглядит так:

- Для каждого семестра запись использует один источник базовой цены:
  - `CATEGORY`, если предмет привязан к `SubjectCategoryLink`;
  - `PRICE_OPTION`, если категориальной привязки нет.
- `base_amounts` в `AUTO` — это историческая база по месяцам, а не кэш, который можно тихо пересчитывать на каждом сохранении.
- Категориальная цена должна быть прозрачной: система показывает suggestion/diff, а изменение `base_amounts` происходит только через осознанное применение и approval flow.
- Финальные `month_01..month_12` должны получаться из `base_amounts` через единый расчёт:
  - скидки;
  - статус договора по месяцам;
  - правила округления и обнуления.
- Размер группы:
  - в effective size считаются `ACTIVE + PAUSED`;
  - `TERMINATED` не считаются;
  - billable count — только `ACTIVE`.
- Длительность занятия, ручной размер группы и тип/статус договора должны поддерживать помесячную историю.
- Границы семестров должны браться из `SemesterConfig`, а не быть захардкожены как 1..6 / 7..12.

## Что реально реализовано

### Уже есть и выглядит полезно

- `backend/apps/catalog/pricing.py`
  - умеет определять `PricingSource`;
  - умеет считать suggested price для GROUP и INDIVIDUAL;
  - учитывает `ContractTypeEntry`, `DurationEntry`, `GroupSizeEntry`, scaling.
- `backend/apps/catalog/group_size_service.py`
  - умеет считать auto group size;
  - использует `counts_in_group_size()` и исключает terminated.
- `backend/apps/catalog/warnings.py`
  - показывает диагностические предупреждения и mismatch между suggested price и `base_amounts`.
- `backend/apps/karteien/category_price.py`
  - умеет применить категориальную цену в `base_amounts`;
  - после этого пересчитывает итоговые месяцы.
- `backend/apps/karteien/models.py`
  - уже содержит `ContractTypeEntry` и `ContractStatusEntry`;
  - уже умеет вычислять effective type/status по месяцу.
- `backend/apps/catalog/views.py`
  - уже содержит UI для групп, bulk preview и bulk apply.

### Главный факт

Новая логика не стала ядром системы. Она добавлена рядом со старой логикой, а не вместо неё.

## Критические расхождения и источники багов

### 1. Нет единого ядра расчёта базы

`backend/apps/karteien/billing.py` по-прежнему является главным биллинговым движком:

- `build_base_amounts()` строит базу по старой схеме `subject_ref + price_ref`;
- категориальная логика там не участвует;
- GROUP/INDIVIDUAL category pricing живёт отдельно в `catalog/pricing.py`.

Следствие:

- preview в форме, перерасчёт `AUTO`, LEGACY -> AUTO и реальное применение category price используют разные расчётные пути;
- один и тот же record может показывать разные ответы в зависимости от точки входа.

### 2. Две параллельные модели договора

В проекте одновременно существуют:

- legacy-поля записи:
  - `is_monthly_contract`
  - `is_contract_terminated`
  - `contract_terminated_from_month`
  - `contract_type_raw`
  - `contract_status_raw`
- новые timeline-модели:
  - `ContractTypeEntry`
  - `ContractStatusEntry`

Проблема в том, что основная форма записи всё ещё работает в legacy-модели:

- `backend/apps/karteien/forms.py::_configure_contract_fields()`
- `backend/apps/karteien/forms.py::_validate_contract_effective_months()`
- `backend/apps/karteien/forms.py::_apply_contract_fields()`

Эти методы:

- инициализируют форму из boolean-полей;
- валидируют изменения относительно boolean-полей;
- сохраняют изменения обратно в boolean-поля.

При этом другие части системы уже считают effective status/type из timeline entries.

Следствие:

- UI, approvals и финальный расчёт могут смотреть на разные источники истины;
- пользователь видит “активный” или “monthly/yearly” в одном месте, а начисление считает по другому состоянию.

### 3. Live preview по сути считает старый мир

`backend/apps/karteien/api.py::billing_preview_api()`:

- строит базу через `build_base_amounts()`;
- preview termination делает через legacy-поля;
- не использует category pricing как базовый источник;
- не моделирует timeline entries как основной механизм.

Следствие:

- пользователь меняет поля в форме, видит один preview;
- category apply / save / pending-change дают другой результат.

Это один из самых разрушительных UX-багов, потому что ломает доверие к интерфейсу.

### 4. `base_amounts` и `month_*` живут несинхронно

Сейчас есть как минимум четыре разных сценария мутации:

- `recalculate_record_months()` в `billing.py`;
- `recalculate_legacy_to_auto()` в `billing.py`;
- `apply_category_price_to_record()` в `category_price.py`;
- ручные/approval-ветки во views.

Проблема не только в дублировании. Current category apply paths (`ApplyCategoryPriceView`, `BulkApplyCategoryPriceView`, `DisciplineGroupPrepareLegacyView`) могут prewrite'ить live `base_amounts` и `status` до решения Superadmin, тогда как proposed `month_*` по tracked fields находятся в `PendingChange.snapshot`. В результате запись может находиться в `PENDING` с proposal-версией `base_amounts` на live record и approved/current `month_*` на той же live record.

`DECLINED` сейчас восстанавливает `_old_base_amounts` для путей, которые положили этот metadata key в snapshot. Это полезная совместимость после `PROMPT_159`, но она path-specific и не превращает snapshot в полноценный frozen v2 payload для всех billing/context changes.

Следствие:

- pending/decline semantics остаются path-dependent;
- review может не видеть единый атомарный billing/context payload;
- future non-tracked поля способны менять финансовый смысл без полного покрытия legacy `TRACKED_FIELDS`.

### 5. `detect_meaningful_changes()` не знает о новой ценовой модели

`backend/apps/karteien/billing.py::detect_meaningful_changes()` отслеживает старые billing-зависимости:

- `price*_ref`
- `subject*_ref`
- `start/end`
- `months_csv`
- скидки
- legacy contract fields
- hours

Но функция не знает о критичных новых зависимостях:

- смена `SubjectCategoryLink`;
- смена параметров `SubjectCategory`;
- смена `DurationEntry`;
- смена `GroupSizeEntry`;
- смена `ContractTypeEntry` / `ContractStatusEntry`.

Следствие:

- запись может остаться “как будто без meaningful change”, хотя логика начисления для неё объективно изменилась;
- mismatch warning появляется, но базовый workflow продолжает жить по старым условиям.

### 6. Групповая модель incomplete даже в catalog

`backend/apps/catalog/group_size_service.py` прямо описывает себя как best-effort without pending projection.

Это конфликтует с ТЗ, где ожидалась более полная модель группового размера, включая pending-контекст и устойчивое поведение при переходных состояниях.

Следствие:

- группа в detail/bulk preview может быть рассчитана неполно;
- real-world сценарии с pending-изменениями по нескольким детям будут выглядеть нестабильно.

### 7. Не до конца применены старые prompt-фиксы

Подтверждены незакрытые/частично закрытые темы:

- `PROMPT_158_FIX_CONTRACT_STATUS_CHANGE_ACTIVE_OVERRIDE_EXISTING_ENTRY.md`
  - баг с заменой статуса в том же месяце;
  - частично видно, что в `ContractStatusChangeView` есть логика exclude same month, но это место всё равно требует финальной ревизии.
- `PROMPT_162_FIX_OPERATOR_GROUP_NAVIGATION.md`
  - `DisciplineGroupListView` всё ещё на `CatalogAdminMixin`;
  - bulk apply формально доступен в backend, но недоступен из UI.
- `PROMPT_160_FIX_DEFAULT_CATEGORIES_ADD_NACHHILFE_TO_GROUP.md`
  - здесь проблема глубже, чем просто “добавить слово в список”:
  - `SubjectCategory` имеет unique constraint на `(year, name)`;
  - значит одна и та же категория `"Nachhilfe"` не может существовать одновременно как `GROUP` и `INDIVIDUAL`.

Следствие:

- ТЗ и текущая схема данных в части default categories противоречат друг другу.

### 8. Test/QA baseline был добавлен после исходного аудита, но не закрывает v2 defects

Post-179/185 update: после этого аудита в репозитории появились `pytest`, `pytest-django`, `pytest-playwright`, `pytest.ini`, `backend/tests/` и minimal fixture-backed browser smoke в `backend/tests/browser/`. Это улучшает regression contour, но не меняет выводы этого документа о PRICELIST V2 split-brain/current broken state. Browser smoke покрывает только non-destructive role/route checks; PRICELIST V2 preview/apply/approve parity, no-live-prewrite и frozen snapshot semantics остаются target gaps/xfail до `PROMPT_166...PROMPT_178`.

Следствие:

- нельзя писать, что UI/UX или approval/billing scenarios “покрыты тестами”;
- система развивается с высоким регрессионным риском;
- каждый будущий prompt должен либо добавлять тест, либо давать manual QA checklist.

## Корневые причины

Если убрать детали, проблем всего несколько:

1. Нет одного authoritative engine для расчёта.
2. Нет одного authoritative состояния договора.
3. `base_amounts` и `month_*` обновляются несколькими независимыми путями.
4. Preview, apply, save, approve и decline не используют один и тот же pipeline.
5. Доменные инварианты из `TZ_PRICELIST_V2.md` описаны лучше, чем реально закреплены кодом.

## Что стоит считать целевой упрощённой моделью

Ниже не “идеальный rewrite”, а целевая архитектура, которая даёт шанс стабилизировать систему без потери требований.

### А. Разделить две задачи, которые сейчас смешаны

Нужны два разных, но общих pure-like слоя:

1. `BaseAmountProposalService`
   - считает, какой базовый amount должен быть по каждому месяцу;
   - умеет работать от `CATEGORY` или `PRICE_OPTION`;
   - умеет считать diff только для touched months;
   - ничего сам не сохраняет.

2. `MonthValueCalculator`
   - берёт уже подготовленные `base_amounts`;
   - применяет скидки;
   - применяет contract status/type timeline;
   - отдаёт итоговые `month_*`.

Тогда:

- form preview использует те же сервисы;
- category apply использует те же сервисы;
- LEGACY -> AUTO использует те же сервисы;
- contract status change preview использует те же сервисы.

### Б. `base_amounts` оставить исторической истиной для AUTO

Это важно не ломать.

Корректная модель:

- изменения в categories/groups/duration не должны молча переписывать уже сохранённые `base_amounts`;
- они должны менять suggested price / diff / warnings;
- фактическое обновление `base_amounts` должно происходить только через явный apply workflow.

### В. Timeline entries сделать главной моделью договора

Нужно выбрать один authoritative источник истины:

- `ContractTypeEntry`
- `ContractStatusEntry`

Legacy boolean-поля можно временно оставить как back-compat fallback, но:

- основная форма не должна на них опираться как на primary source;
- preview не должен использовать их как primary source;
- финальный расчёт должен смотреть на timeline helpers.

### Г. Approval flow должен быть атомарным

Для billing-affecting изменений есть два разумных пути.

Путь 1. Консервативный:

- пока change pending, live record не получает новых `base_amounts` и `month_*`;
- всё новое лежит только в snapshot.

Путь 2. Минимально инвазивный:

- если live record меняется до approval, то decline обязан полностью восстанавливать и `base_amounts`, и `month_*`, и metadata timeline-переходов.

Текущая реализация не делает ни первое, ни второе до конца.

## Рекомендуемый порядок стабилизации

### Фаза 1. Свести источники истины

- Зафиксировать один pipeline для preview/apply/save.
- Вынести общий расчёт proposed `base_amounts`.
- Переподключить `billing_preview_api()` на этот pipeline.

### Фаза 2. Выровнять договорную модель

- Сделать `ContractTypeEntry` / `ContractStatusEntry` первичными.
- Форму редактирования перевести на timeline helpers.
- Legacy boolean-поля оставить только как fallback/совместимость до полной миграции.

### Фаза 3. Выровнять approval semantics

- Решить, меняется ли live record до approval.
- Закрыть rollback inconsistency.
- Убедиться, что approval/decline одинаково корректно работают для:
  - category apply;
  - contract type change;
  - contract status change;
  - LEGACY -> AUTO conversion.

### Фаза 4. Закрыть UI-конфликты и маленькие structural bugs

- доступ Operator к группам;
- breadcrumbs / navigation;
- финальное решение по dual `"Nachhilfe"` category;
- прозрачные предупреждения там, где preview не может быть достоверным.

### Фаза 5. Тесты

Минимальный обязательный пакет:

- pure calculation tests;
- preview vs apply consistency tests;
- approval/decline rollback tests;
- timeline contract tests;
- group scaling tests;
- regression tests на edge cases из отдельного чеклиста.

## Что GPT PRO должен считать самыми опасными зонами

Если нужно приоритизировать, то вот порядок риска:

1. `backend/apps/karteien/forms.py`
2. `backend/apps/karteien/api.py`
3. `backend/apps/karteien/billing.py`
4. `backend/apps/approvals/services.py`
5. `backend/apps/karteien/views.py`
6. `backend/apps/karteien/category_price.py`
7. `backend/apps/catalog/pricing.py`
8. `backend/apps/catalog/group_size_service.py`

## Ключевой вывод

Проблема не в одном баге и не в трёх пропущенных prompt-файлах. Проблема в том, что прайс v2 внедрён как надстройка поверх старого биллинга, а не как замена ядра. Пока не будет одного расчётного pipeline и одного источника истины по договору, локальные фиксы будут продолжать конфликтовать друг с другом.
