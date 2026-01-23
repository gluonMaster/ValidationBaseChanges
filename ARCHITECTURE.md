# System Architecture – KindEltern Web

Этот документ описывает высокоуровневую архитектуру веб‑системы, которая заменяет набор Excel+VBA файлов (`KindElternDaten_XX_Admin.xlsm`, `KindElternDaten_XX_Suprime.xlsm`, `KindElternDaten_XX_Data.xlsm`) и Access‑баз (`KindElternDaten_XX_front.accdb`).

Примечание для нового чата: чтобы экономить контекст, начните с `docs/ProjectMap.md`, а здесь открывайте только нужные разделы (через поиск по файлу).

## 1. Overview

- Исходники legacy VBA-модулей для справки лежат в `legacy_VBA/`:

  - `legacy_VBA/*.bas` и `legacy_VBA/*.cls` — модули из `KindElternDaten_XX_Admin.xlsm`
  - `legacy_VBA/admin_forms/` — VBA-код форм из `KindElternDaten_XX_Admin.xlsm`
  - `legacy_VBA/Superadmin/` — модули из `KindElternDaten_XX_Suprime.xlsm`
  - `legacy_VBA/alt/` — модули из `KindElternDaten_XX_Data.xlsm`

- Цель системы: единый веб‑интерфейс для:
  - **Admin**: ведение картотеки семей/детей, синхронизация, работа с отклонёнными изменениями.
  - **Superadmin**: утверждение/отклонение рискованных изменений.
  - **User**: просмотр истории и статуса записей.
- Источник истины: централизованная реляционная БД (PostgreSQL), заменяющая Access‑файлы.
- Основные концепции заимствованы из существующей схемы:
  - `tblKartei` → основная таблица (`karteien_record`).
  - `pre_tblKartei` → pending изменения.
  - `decl_tblKartei` → отклонённые изменения.
  - История изменений хранится в отдельной структуре, но поддерживает текущий текстовый формат для обратной совместимости.

Технологический стек:

- Backend: Django + Django REST Framework.
- DB: PostgreSQL.
- Frontend: Django templates + HTMX / лёгкий JS (без тяжёлого SPA на первом этапе).
- Deployment: Docker + docker-compose (локально и на VPS).

## 2. Top-Level Modules

Проект разделён на несколько Django‑приложений (apps) по областям ответственности.

### 2.1 `accounts` (пользователи и роли)

**Задачи:**

- Регистрация/управление пользователями.
- Ролевое управление доступом:
  - Role `ADMIN` (аналог Admin файла).
  - Role `OPERATOR` (ограниченный Admin).
  - Role `SUPERADMIN` (аналог Suprime файла).
  - Role `USER` (только просмотр, read-only доступ).

**Ключевые сущности:**

- `User` (расширение стандартной Django модели).
- Поле `role` с выбором из `UserRole` enum.

**Интерфейсы:**

- Страницы логина/логаута.
- Простое управление пользователями через Django admin.

**User Cabinet (PROMPT_10):**

Реализован веб-интерфейс для роли User (read-only):

- **User Dashboard** (`/user/`):

  - Стартовая страница после логина для роли User.
  - Форма поиска записей.
  - Информация о правах доступа.

- **User Search** (`/user/search/`):

  - Поиск записей по FamilyID, Parent, Child, Year.
  - Таблица результатов с ссылками на детали и историю.
  - Лимит: 50 записей.

- **User Record Detail** (`/user/record/<id>/`):

  - Read-only просмотр записи KarteiRecord.
  - Отображение: основные данные, предметы, месяцы, статус.
  - Индикатор pending/declined изменений.
  - Нет кнопок редактирования.

- **User Record History** (`/user/record/<id>/history/`):
  - Просмотр истории изменений записи (аналог Geschichte).
  - Хронологический список событий HistoryEvent.
  - Для каждого события: дата/время, пользователь, изменённые поля (old → new), комментарий.

**Ключевые ограничения User role:**

- Строго read-only на серверном уровне (UserRoleMixin блокирует все POST/PUT/DELETE).
- Нет доступа к редактированию Kartei (`/karteien/`).
- Нет доступа к approvals workflow.
- При попытке доступа к Admin/Superadmin URL — редирект на `/user/`.

**Role-based redirect:**

- Функция `role_based_redirect()` на корневом URL `/`:
  - USER → `/user/` (User Dashboard)
  - ADMIN/OPERATOR → `/karteien/` (Kartei List)
  - SUPERADMIN → `/approvals/superadmin/pending/` (Pending Overview)

### 2.2 `karteien` (основная картотека)

**Происхождение из VBA:**

- `ImportData`, `ExportSyncKartei`, `ExportUtilities`, `ExportProtection`, часть форматных модулей.

**Задачи:**

- Хранение базовых данных семей/детей:
  - FamilyID, Parent, Child, Birthdate, Address, Phone, Mobile, Email.
  - Subjects, Prices, месяцы U–AF (12 месяцев).
  - Доп. предметы AK–AM.
  - Статус строки (PENDING/DECLINED/normal) и ID (аналог колонки AV).
- Операции Admin:
  - Создание/редактирование/удаление записей.
  - Мастер **Neue Familie**: создание новой семьи с несколькими детьми за один проход (`/karteien/family/new/`), с назначением скидок и пересчётом `AUTO`.
  - Синхронизация изменений в БД (внутри веб‑системы).
  - Правила по годам и ограничениям (аналог `ExportProtection.ValidateAndFixPastMonths`).
- Просмотр (read-only):
  - Superadmin и User могут открывать `/karteien/` (list/detail + AJAX live-search) без edit/delete.
- Валидации:
  - Уникальность FamilyID + Parent (аналог `Export_ValidationKartei`).
  - Ограничения по прошлым месяцам и ролям.

**Ключевые модели:**

- `KarteiRecord` (`apps/karteien/models.py`):
  - Primary key: `pkid` (surrogate Django PK).
  - Access/Excel ID: `id` (Access `ID` / Excel AV/48), доменный ключ вместе с `year`.
  - Обязательное поле `year` для разделения 2024/2025/… (вместо отдельных файлов).
  - Основные поля: `family_id`, `parent_name`, `child_name`, `birthdate`, `address`, `phone`, `mobile`, `email`.
  - Предметы и цены: `subject1`, `price1`, `subject2`, `price2`, `extra1..3`.
  - Ref‑поля справочников: `subject*_ref`, `teacher*_ref`, `price*_ref` + стартовые месяцы `start_month_1/2`.
  - Режим начислений: `months_mode` = `LEGACY`/`AUTO`/`OVERRIDE`.
  - Данные расчёта: `base_amounts` (суммы до скидок), `hours_amounts` (UE по месяцам для Stundenfächer).
  - Отключение скидок по месяцам: `discounts_disabled` + `discounts_disabled_months` (список месяцев 1-12).
  - Месячные поля: `month_1` … `month_12` (DecimalField для начислений/платежей).
  - Служебные: `sepa_marker`, `status`, `last_change_role/date/time`, `history_raw`.
  - Контракт (legacy): `teacher1_legacy_name`, `teacher2_legacy_name`, `contract_type_raw`/`is_monthly_contract`, `contract_status_raw`/`is_contract_terminated`, `contract_terminated_from_month`.
  - Поле `status` (enum: normal/pending/declined) для быстрых фильтров.
  - Уникальное ограничение: `(year, id)`.
- В дальнейшем возможно выделение отдельных сущностей (`Family`, `Child`, `Subscription`), но первый шаг максимально близок к текущей структуре.

**Catalog Integration (PROMPT_20):**

Формы создания/редактирования записей (`KarteiRecordForm`) интегрированы со справочниками:

- **Subject/Teacher/Price selection**: Вместо свободного ввода admin выбирает предмет, преподавателя и цену из справочников (`catalog.Subject`, `catalog.Teacher`, `catalog.PriceOption`).
- **Dynamic loading**: При смене предмета JavaScript динамически подгружает списки преподавателей и цен для выбранного предмета и года через API endpoints (`/api/catalog/teachers/`, `/api/catalog/prices/`).
- **Validation**: Форма валидирует, что выбранная цена принадлежит году записи и выбранному предмету; преподаватель имеет `TeachingAssignment` для предмета и года.
- **Legacy sync**: При сохранении legacy-поля `subject1/subject2/price1/price2` синхронизируются с выбранными ref-полями для обратной совместимости с импортом/экспортом.
- **Start months**: Поля `start_month_1` (1-6) и `start_month_2` (7-12) определяют, с какого месяца начинается начисление в каждом полугодии.
- **Legacy-подсказки и быстрые ссылки (PROMPT_65…72):** Если legacy-значения `Fach/Lehrer/Preis` не соответствуют каталогам, UI показывает информативные бейджи и даёт быстрые ссылки на create-формы справочников с предзаполнением + `next` возвратом обратно в редактирование записи.
- **Live-preview начислений (PROMPT_51,57,58,79):** В edit есть API-preview (`karteien:billing_preview_api`) и модалка с деталями начисления при клике на месяц.

**Сервисный слой (`apps/karteien/services.py`):**

- `KarteiSyncService` — синхронизация записей (аналог ExportSyncKartei).
- `KarteiValidationService` — валидация данных (уникальность, прошлые месяцы, SEPA).
- Константы: `MONTH_FIELD_NAMES`, `TRACKED_FIELDS`, `HISTORY_FIELD_TAGS`.

### 2.3 `approvals` (pending/declined workflow и риск‑классификация)

**Происхождение из VBA:**

- `Export_RiskClassification`, `Export_OverlayPending`, `Export_DeclinedTools`, `Export_DeclinedHelpers`,
- `valid_ImportPending`, `valid_ApproveFlow`.

**Задачи:**

- Определение, какие изменения считаются “risk” и требуют Superadmin:
  - Использует политику, аналогичную `RISK_POLICY_MODE = "TRACKED_FIELDS"`.
- Отдельное хранение:
  - Pending изменений (аналог `pre_tblKartei`).
  - Declined изменений (аналог `decl_tblKartei`).
- Workflow:
  - Admin вносит изменения → safe изменения идут напрямую в `KarteiRecord`, risky — создают/обновляют запись в pending.
  - Superadmin просматривает pending, сравнивает “War/Ist” и ставит `Approved/Declined`.
  - Approved → изменения применяются к `KarteiRecord`, pending‑запись помечается обработанной.
  - Declined → создаётся запись в “declined” и отправляется сигнал/уведомление Admin’у.

**Ключевые модели:**

- `PendingChange` (`apps/approvals/models.py`):
  - `id` — auto-generated BigAutoField (не совпадает с ID Kartei).
  - `record` — OneToOneField на `karteien.KarteiRecord` (on_delete=CASCADE).
  - `snapshot` — JSONField со снимком всех tracked-полей (формат: `{"field_name": value, ...}`).
  - `is_processed` — bool, True после обработки (approve/decline).
  - `created_at`, `updated_at` — timestamps.
- `DeclinedChange` (`apps/approvals/models.py`):
  - `id` — auto-generated BigAutoField.
  - `record` — ForeignKey на `karteien.KarteiRecord` (допускает несколько decline для одной записи).
  - `snapshot` — JSONField со снимком отклонённых значений.
  - `decline_reason` — TextField с причиной отклонения.
  - `declined_by` — ForeignKey на `accounts.User` (Superadmin).
  - `created_at`, `updated_at` — timestamps.

**Ключевые сервисы** (`apps/approvals/services.py`):

- `is_risky_change(local: KarteiRecord, original: KarteiRecord | None) -> bool`:
  - Для новых записей (original=None) → возвращает `False` (safe).
  - Для существующих записей → сравнивает tracked-поля, возвращает `True` если хоть одно изменилось.
- `classify_change(local, original) -> Literal["SAFE", "RISKY"]`:
  - Удобная обёртка над `is_risky_change`.
- `get_changed_tracked_fields(local, original) -> dict[str, tuple[old, new]]`:
  - Возвращает словарь изменённых tracked-полей для отображения "War/Ist".
- `build_snapshot(record: KarteiRecord) -> dict`:
  - Строит JSON-сериализуемый снимок tracked-полей.
- `create_or_update_pending_change(record) -> PendingChange`:
  - Создаёт или обновляет pending-запись (заглушка, полная интеграция позже).
- `create_declined_change(record, reason, declined_by) -> DeclinedChange`:
  - Создаёт decline-запись (заглушка, полная интеграция позже).
- `apply_pending_change(pending) -> KarteiRecord`:
  - Применяет approved изменения к записи (заглушка).

**Superadmin Web UI (PROMPT_09):**

Реализован веб-интерфейс для роли Superadmin:

- **Pending Overview** (`/approvals/superadmin/pending/`):

  - Список всех неподтверждённых pending-записей.
  - Фильтры: по году, FamilyID, Parent, Child.
  - Массовые действия: "Approve All", "Decline All" с общим комментарием.
  - Ссылки на War/Ist для каждой записи.

- **War/Ist View** (`/approvals/superadmin/pending/<id>/`):

  - Сравнение оригинальных данных (KarteiRecord) и pending-снимка.
  - Подсветка изменённых полей.
  - Форма решения: Approved/Declined с комментарием.
  - Ссылка на историю записи.

- **NeuList** (`/approvals/superadmin/neu/`):

  - Список новых записей **по выбранному году**: `id > last_seen_by_year[year]` (доменный ключ `(year, id)`).
  - Кнопка "Mark as Seen" обновляет `last_seen_by_year[year]`.
  - Аналог VBA valid_NeuList.

- **Record History** (`/approvals/superadmin/history/<id>/`):
  - Просмотр истории изменений одной записи.
  - Timeline с событиями (CHANGE, APPROVE, DECLINE, CREATE).
  - Аналог VBA Geschichte.bas.

**Ключевые сервисы для Superadmin** (`apps/approvals/services.py`):

- `apply_decision(pending, decision, comment, user) -> KarteiRecord`:

  - Применяет решение Superadmin к pending-записи.
  - APPROVED: применяет snapshot к записи, статус → NORMAL.
  - DECLINED: создаёт DeclinedChange, статус → DECLINED.
  - Записывает APR/DCL в history_raw.
  - Создаёт уведомления для Admin.

- `approve_all_pending(user, year) -> (count, errors)`:

  - Массовое одобрение всех pending-записей.

- `decline_all_pending(comment, user, year) -> (count, errors)`:

  - Массовое отклонение с общим комментарием.

- `get_new_records(user, year) -> list[KarteiRecord]`:

  - Возвращает записи `KarteiRecord(year=year, id__gt=last_seen_by_year[year])`.

- `update_last_seen_id(user, year, max_id=None)`:
  - Обновляет `last_seen_by_year[year]` для NeuList (и может поддерживать legacy `last_seen_id` как диагностическое поле).

**Модель SuperadminState** (`apps/approvals/models.py`):

- `user` — OneToOneField на User.
- `last_seen_by_year` — JSONField: `{ "<year>": <last_seen_id>, ... }` для NeuList (per-year).
- `last_seen_id` — legacy PositiveIntegerField (не используется для фильтрации NeuList per-year; нужен только для совместимости/диагностики).
- `last_seen_date` — DateTimeField (альтернативный трекер).

**Admin/Operator Web UI:**

Реализован веб-интерфейс для ролей Admin/Operator:

- **Kartei List View** (`/karteien/`):
  - Список записей с фильтрами по году, FamilyID, Parent, Child, статусу, типу/статусу договора (в т.ч. SEPA-варианты), а также динамическими фильтрами Unterricht/Lehrer по семестру.
  - Быстрые ссылки на PENDING/DECLINED обзоры.
- **Kartei Create/Edit** (`/karteien/create/`, `/karteien/<id>/edit/`):

  - Формы с валидациями: уникальность FamilyID+Parent, пустой FamilyID при непустом Parent.
  - Ограничения для Operator: SEPA-строки, прошлые месяцы.
  - Классификация изменений: safe → прямое обновление, risky → создание PendingChange.
  - Поле комментария (Notitzen) для risky-изменений.
  - Create: UX-подсказка по FamilyID (max + рекомендация) и предупреждение для Freitext; entry-point создания записи из Family Dashboard с prefill данных семьи.

- **Declined Overview** (`/approvals/declined/`):

  - Список всех DECLINED-записей (аналог VBA DeclinedOverview).
  - Возможность применить исправления (переместить в PENDING).
  - Кнопка "Применить все" для массовой обработки.

- **Pending List** (`/approvals/pending/`):
  - Информационный список изменений, ожидающих Superadmin.

**Ограничения по статусу записи:**

Редактирование через `/karteien/<id>/edit/` разрешено **только для записей со статусом NORMAL**:

- **NORMAL**: Admin/Operator могут редактировать запись. SAFE-изменения (не затрагивающие tracked-поля) применяются напрямую к `KarteiRecord`. RISKY-изменения создают `PendingChange` и переводят запись в статус PENDING.

- **PENDING**: Редактирование через Kartei UI заблокировано. Запись ожидает решения Superadmin через approvals-флоу (PROMPT_09). При попытке открыть форму редактирования пользователь получает сообщение об ошибке и перенаправляется на страницу деталей записи.

- **DECLINED**: Редактирование через Kartei UI заблокировано. Для исправления и повторной отправки Admin/Operator должны использовать DeclinedOverview (`/approvals/declined/`). После исправления запись переходит обратно в статус PENDING.

В UI (списки и детали записей) кнопка "Bearbeiten" отображается только для NORMAL-записей. Для PENDING показывается индикатор ожидания, для DECLINED — ссылка на DeclinedOverview.

**Валидаторы** (`apps/karteien/validators.py`):

- `validate_family_id_parent_unique()` — проверка дубликатов FamilyID+Parent.
- `validate_family_id_not_empty_with_parent()` — проверка пустого FamilyID.
- `validate_sepa_restrictions()` — SEPA-ограничения для Operator.
- `validate_past_months_restrictions()` — ограничения по прошлым месяцам.
- `apply_operator_filters()` — фильтрация запрещённых изменений.

### 2.4 `history` (история изменений и отчёты)

**Происхождение из VBA:**

- `Export_HistoryBuilder`, `Export_HistoryConverter`, `Export_HistoryParser`,
- `valid_ParseHistory`, `Geschichte`, `grossGeschichte`.

**Задачи:**

- Хранение истории изменений по каждой записи:
  - Исторические события (кто, когда, какие поля, старое/новое значения).
  - Поддержка текущего текстового формата истории (для плавной миграции), но с возможностью нормализованного хранения.
- Построение представлений “истории одной записи”:
  - Аналог листа `Geschichte` (интерактивная лента изменений).
- Построение агрегированных отчётов:
  - Аналог `grossGeschichte` / `valid_GrossGeschichteDecision` (War/Ist сравнение для множества записей).

**Ключевые модели:**

- `HistoryEvent` (`apps/history/models.py`):
  - `id` — BigAutoField (PK).
  - `record` — FK на `karteien.KarteiRecord` (on_delete=CASCADE).
  - `user` — FK на `accounts.User` (null=True, для legacy-данных).
  - `event_time` — DateTimeField (время события, парсится из истории или создаётся при изменении).
  - `event_type` — CharField с choices: `CHANGE`, `CREATE`, `APPROVE`, `DECLINE`, `IMPORT`, `RUCK`.
  - `changes` — JSONField: структура `{"field_name": {"old": ..., "new": ...}}`.
  - `comment` — TextField (комментарий из `/@...@/`).
  - `raw_history_fragment` — TextField (исходный сегмент строки истории из AZ/Value52).

**Ключевые сервисы** (`apps/history/services.py`):

- `parse_raw_history(raw: str) -> list[HistoryEventData]`:
  - Парсит сырую строку истории (AZ/Value52) в список структурированных событий.
  - Поддерживает новый формат `[RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||`
  - Поддерживает legacy формат `Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY ||`.
  - Не теряет comment-only события (например, `APR:`/`ADM:`), и извлекает дату из записей вида `DCL(...)/@...@/<DATE>`.
- `sync_history_from_raw(record: KarteiRecord) -> list[HistoryEvent]`:
  - Парсит `history_raw` записи и создаёт отсутствующие `HistoryEvent`.
  - Не изменяет исходную строку истории.

**API endpoints:**

- `GET /api/history/records/<id>/` — список событий истории для записи.
- `GET /api/history/events/<id>/` — детали одного события.
- `POST /api/history/records/<id>/sync/` — синхронизация истории из raw.

### 2.5 `notifications` (уведомления)

**Задачи:**

- Внутренние уведомления:
  - Superadmin: новые pending‑записи.
  - Admin: новые отклонённые записи, требующие правки.
- Механизм:
  - В первой версии — просто записи в таблице + бейджики/списки в UI.
  - Позже: e‑mail / Telegram / WebSocket.

**Ключевые модели:**

- `Notification` (`apps/notifications/models.py`):
  - `id` — BigAutoField (PK).
  - `recipient` — FK на `accounts.User` (получатель уведомления).
  - `type` — CharField с choices: `PENDING_CREATED`, `DECLINED_CREATED`.
  - `record` — FK на `karteien.KarteiRecord` (nullable, ссылка на связанную запись).
  - `payload` — JSONField (дополнительные данные: family_id, parent_name, decline_reason и т.п.).
  - `created_at` — DateTimeField (auto_now_add).
  - `read_at` — DateTimeField (nullable, дата прочтения).

**Типы уведомлений:**

- `PENDING_CREATED`: Создаётся при создании PendingChange, отправляется всем Superadmin.
- `DECLINED_CREATED`: Создаётся при создании DeclinedChange, отправляется всем Admin.

**Ключевые сервисы** (`apps/notifications/services.py`):

- `notify_pending_created(record, pending)`:
  - Находит всех Superadmin пользователей.
  - Создаёт уведомления типа `PENDING_CREATED`.
  - Идемпотентность: не создаёт дубликаты если непрочитанное уведомление уже существует.
- `notify_declined_created(record, declined)`:
  - Находит всех Admin пользователей.
  - Создаёт уведомления типа `DECLINED_CREATED`.
  - Идемпотентность: не создаёт дубликаты если непрочитанное уведомление уже существует.
- `get_unread_count(user)`:
  - Возвращает количество непрочитанных уведомлений пользователя.
- `get_notifications(user, unread_only, notification_type, limit)`:
  - Возвращает список уведомлений с фильтрацией.
- `mark_notification_read(notification_id, user)`:
  - Помечает уведомление как прочитанное.
- `mark_all_notifications_read(user)`:
  - Помечает все уведомления пользователя как прочитанные.

**Связь с approvals:**

- `create_or_update_pending_change()` вызывает `notify_pending_created()`.
- `create_declined_change()` вызывает `notify_declined_created()`.

**API endpoints:**

- `GET /api/notifications/` — список уведомлений текущего пользователя.
  - Query params: `unread=true`, `type=PENDING_CREATED|DECLINED_CREATED`, `limit=N`.
- `GET /api/notifications/unread-count/` — количество непрочитанных уведомлений.
- `POST /api/notifications/<id>/read/` — пометить уведомление как прочитанное.
- `POST /api/notifications/read-all/` — пометить все уведомления как прочитанные.

**UI:**

- Базовый шаблон (`templates/base.html`) содержит "колокольчик" с бейджем непрочитанных уведомлений.
- Dropdown со списком последних уведомлений и возможностью пометить как прочитанные.
- Автоматическое обновление счётчика каждые 60 секунд.

### 2.6 `reporting` (отчёты и аналитика)

**Задачи:**

- Более сложные отчёты, которые не укладываются в простые списки (dashboards для руководства).
- Использует данные из `karteien`, `history`, `approvals`.
- Обзор последних изменений (аналог `grossGeschichte`).

**API endpoints:**

- `GET /api/reporting/recent-changes/` — список последних изменений за период.
  - Параметры: `limit`, `offset`, `year`, `date_from`, `date_to`, `event_type`, `family_id`.
  - Возвращает события с контекстом записей (ID, родитель, ребёнок, дата).
- `GET /api/reporting/recent-changes-by-record/` — изменения, сгруппированные по записям.
  - Показывает записи с последними изменениями и количество событий.

### 2.7 `catalog` (справочники)

**Назначение:**

Справочные таблицы для стандартизации данных о преподавателях и предметах.
Эти таблицы **не заменяют** существующие поля `KarteiRecord.subject1/subject2`
и не влияют на импорт из Access. Это отдельные справочники для:

- Стандартизации названий предметов и имён преподавателей.
- Связи "кто вёл что в каком году".
- Будущих форм выбора с автодополнением.
- Отчётности и аналитики.

**Ключевые модели:**

- `Teacher` (`apps/catalog/models.py`):

  - `last_name` — CharField, фамилия преподавателя.
  - `first_name` — CharField, имя преподавателя.
  - `is_active` — BooleanField, активность.
  - UniqueConstraint на `(last_name, first_name)`.

- `Subject` (`apps/catalog/models.py`):

  - `name` — CharField, уникальное название предмета.
  - `is_active` — BooleanField, активность.

- `TeachingAssignment` (`apps/catalog/models.py`):

  - `year` — PositiveSmallIntegerField, учебный год.
  - `subject` — FK на `Subject`.
  - `teacher` — FK на `Teacher`.
  - `is_active` — BooleanField, активность.
  - UniqueConstraint на `(year, subject, teacher)`.
  - Индексы: `(year, subject)`, `(year, teacher)`.

- `PriceOption` (`apps/catalog/models.py`):

  - `year` — PositiveSmallIntegerField, учебный год.
  - `subject` — FK на `Subject`.
  - `amount` — DecimalField (max_digits=10, decimal_places=2), сумма в €.
  - `comment` — TextField, комментарий (почему такая цена).
  - `is_active` — BooleanField, активность.
  - UniqueConstraint на `(year, subject, amount, comment)`.
  - CheckConstraint: `amount >= 0`.
  - Индекс: `(year, subject)`.
  - **Семантика цены:** По умолчанию это цена за месяц (€/Monat).
    Для предметов с "Ind." или "VSpE\_" (индивидуальные) и "NH"/"Nachhilfe" (наххильфе)
    в названии — это цена за академический час (€/UE).
  - Метод `get_price_unit()` возвращает единицу ("€/Monat" или "€/UE").
  - Метод `is_per_hour()` проверяет тип расчёта.

- `Discount` (`apps/catalog/models.py`):

  - `kind` — CharField (choices: PERCENT, FIXED), тип скидки.
  - `value` — DecimalField, значение скидки:
    - Для PERCENT: 0.00-0.99 (например 0.25 = 25%).
    - Для FIXED: сумма в € (вычитается после процентной скидки).
  - `description` — TextField, описание скидки.
  - `is_active` — BooleanField, активность.
  - Валидация: PERCENT 0..0.99, FIXED ≥ 0.

- `FamilyDiscount` (`apps/catalog/models.py`):

  - `year` — PositiveSmallIntegerField, учебный год.
  - `family_id` — CharField, идентификатор семьи (как в KarteiRecord).
  - `discount` — FK на `Discount`.
  - `start_month`, `end_month` — месяцы действия (1-12).
  - `months` — JSONField, опционально список конкретных месяцев [1,2,3].
  - `created_at`, `updated_at` — timestamps.
  - Применяется ко всем записям семьи в данном году.

- `RecordDiscount` (`apps/catalog/models.py`):
  - `record` — FK на `karteien.KarteiRecord`.
  - `discount` — FK на `Discount`.
  - `start_month`, `end_month`, `months` — как в FamilyDiscount.
  - `created_at`, `updated_at` — timestamps.
  - Применяется к конкретной записи.

**Скидки и расчёт начислений (реализовано):**

Скидки (`FamilyDiscount`, `RecordDiscount`) участвуют в расчёте `AUTO` начислений в `apps/karteien/billing.py`:

- Базовые суммы берутся из `PriceOption` (с учётом `start_month_*` и режима Stundenfächer).
- Затем применяются скидки (сначала процентные, затем фиксированные).
- Поддерживается отключение скидок на всю запись или на конкретные месяцы (`discounts_disabled`, `discounts_disabled_months`).

**Django Admin:**

Все модели зарегистрированы в Django Admin для быстрого наполнения справочников.

**Web UI (Admin):**

- `/catalog/` — главная страница каталога со статистикой.
- `/catalog/teachers/` — управление преподавателями.
- `/catalog/subjects/` — управление предметами.
- `/catalog/assignments/` — управление назначениями (преподаватель-предмет-год).
- `/catalog/assignments/copy-year/` — копирование назначений между годами.
- `/catalog/prices/` — управление прайс-листом по годам и предметам.
- `/catalog/prices/copy-year/` — копирование прайса между годами.
- `/catalog/discounts/` - управление справочником скидок (процентные/фиксированные).
- `/catalog/family-discounts/` - назначение скидок на семью+год.
- `/catalog/record-discounts/` - назначение скидок на конкретные записи.

Create-формы справочников поддерживают предзаполнение через query-params и `next` (для сценариев быстрого добавления из `/karteien/<pkid>/edit/`).

### 2.8 `legacy_import` (миграция из Access)

**Происхождение из VBA:**

- `AccessCreation`, `DumpBase`, `BaseBackupRestore`, импорт Kartei из Access.

**Задачи:**

- Подключение к Access‑файлам (`KindElternDaten_XX_front.accdb`) через pyodbc/ACE ODBC.
- Импорт данных из трёх таблиц:
  - `tblKartei` → `karteien.KarteiRecord`
  - `pre_tblKartei` → `approvals.PendingChange` + `status=PENDING`
  - `decl_tblKartei` → `approvals.DeclinedChange` + `status=DECLINED`
- Маппинг полей `Value1..Value52` → Django-модели (см. DOMAIN_MODEL.md).
- Сохранение `Value52` → `KarteiRecord.history_raw`.
- Фильтрация маркерных строк `"Zahlung"` (семейные разделители не импортируются).
- Анализ и нормализация FamilyID:
  - Обнаружение случаев "одна семья — разные FamilyID".
  - Обнаружение случаев "один FamilyID — разные семьи".
  - Генерация отчётов по проблемам.
  - Опциональная авто-нормализация (`--familyid-policy=auto-merge`).
- Режим `--dry-run` для анализа без записи в БД.
- Опциональная синхронизация истории (`--sync-history`).

**Ключ записи:**

- `(year, id)` — уникальный ключ в `KarteiRecord`.
- `id` берётся из Access поля `ID` (Excel: AV/48).
- `year` передаётся как параметр команды, определяет к какому году относится `.accdb`.
- Один и тот же ID может существовать в разных годах.

**Ключевые модули:**

- `access_client.py` — подключение к Access, чтение таблиц.

  - `open_access_connection(file_path)` — context manager для соединения.
  - `load_tbl_kartei(conn, year)` → Iterable[RowDict].
  - `load_pre_tbl_kartei(conn, year)` → Iterable[RowDict].
  - `load_decl_tbl_kartei(conn, year)` → Iterable[RowDict].

- `services.py` — бизнес-логика импорта.
  - `ImportStats` — dataclass со статистикой импорта.
  - `is_marker_row(row)` — определение маркерных строк.
  - `import_tbl_kartei(rows, year)` → ImportStats.
  - `import_pre_tbl_kartei(rows, year)` → ImportStats.
  - `import_decl_tbl_kartei(rows, year)` → ImportStats.
  - `analyze_familyid_issues(rows, existing, year)` → issues, merge_map.
  - `sync_history_for_records(year, record_ids)` — синхронизация истории.

**Management command:**

```bash
python manage.py import_access_year --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    [--patch-fields] \
    [--dry-run] \
    [--familyid-policy=report|auto-merge] \
    [--sync-history] \
    [--skip-pending] \
    [--skip-declined] \
    [--report-dir=./reports]
```

Дополнительные флаги:

- `--skip-pending` — пропустить импорт `pre_tblKartei` (pending-изменения).
- `--skip-declined` — пропустить импорт `decl_tblKartei` (declined-изменения).
- `--patch-fields` — patch‑режим: читает только `tblKartei` и обновляет только доп.поля (учителя Value11/16, контракт Value14/20 + derived bool, `sepa_marker` Value47) у уже импортированных записей, найденных по `(year, id)`.
- `--report-dir` — каталог для сохранения CSV/JSON-отчётов (по умолчанию — текущий каталог).

**Границы анализа FamilyID:**

Анализ FamilyID в текущей реализации выполняется **в рамках одного года**, задаваемого
параметром `--year`. Для межгодового анализа FamilyID (сравнение семей между годами)
предполагается использовать данные, уже импортированные в PostgreSQL, через отчёты
или аналитические запросы, а не напрямую Access-файлы разных лет.

**Отчёты:**

- `import_stats_YEAR.json` — полная статистика импорта.
- `import_familyid_issues_YEAR.csv` — обнаруженные проблемы FamilyID.
- `import_familyid_merge_YEAR.csv` — маппинг авто-нормализации.

**Конфигурация:**

- `ACCESS_BASE_DIR` — базовый каталог с `.accdb` файлами (env или settings).
- `ACCESS_CONN_STRING_TEMPLATE` — шаблон строки подключения ODBC.

## 3. Key Workflows

### 3.1 Admin workflow (edit & sync)

1. Admin авторизуется и открывает список `KarteiRecord` за нужный год.
2. Вносит изменения в записи.
3. Сохранение:
   - Система сравнивает изменённые поля (аналог `FindChangedIDs` + `HasTrackedFieldChanges`).
   - Safe изменения → напрямую обновляют `KarteiRecord` + создают историю.
   - Risky изменения → создают/обновляют `PendingChange` и помечают исходную запись как PENDING.
4. При создании `PendingChange` система создаёт `Notification` для всех Superadmin'ов.

Особенности edit-сценария начислений:

- В `AUTO` режиме начисления рассчитываются из текущих цен/скидок (включая Stundenfächer с `hours_amounts`) и сохраняются в `month_*`.
- В `LEGACY` режиме месячные поля отображаются как read-only (клик открывает пояснение). Для приведения legacy-значений к текущим правилам доступна кнопка **Monate neu berechnen** (конвертирует запись в `AUTO` и пересчитывает месяцы).
- Для ручной правки начислений используется отдельный экран `months-override` (режим `OVERRIDE`), а в UI рядом с месяцами показывается бейдж `Override`.

### 3.2 Superadmin workflow (approve/decline)

1. Superadmin видит количество pending‑записей (бейдж/уведомление).
2. Открывает список pending:
   - Для каждой записи видит War/Ist сравнение (original vs pending), аналог листа `grossGeschichte`.
3. Ставит решение `Approved` или `Declined` + комментарий при отклонении.
4. При синхронизации решений:
   - `Approved` → изменения применяются к `KarteiRecord` + создаётся `HistoryEvent`, pending помечается как обработанный.
   - `Declined` → создаётся `DeclinedChange`, исходная запись помечается как DECLINED, создаётся `Notification` для Admin.
5. Admin получает уведомление и исправляет данные, после чего цикл может повториться.

### 3.3 User workflow (view history)

1. User логинится с правами просмотра.
2. Может искать записи по FamilyID/Parent/Child и смотреть детальную историю (лента событий).

## 4. Code Organization & Conventions

### 4.1 Backend Structure

Проект Django расположен в каталоге `backend/`:

```
backend/
├── manage.py
├── Dockerfile
├── requirements.txt
├── .gitignore
├── config/                    # Django project configuration
│   ├── __init__.py
│   ├── settings.py
│   ├── urls.py
│   ├── wsgi.py
│   └── asgi.py
├── apps/                      # Django applications
│   ├── __init__.py
│   ├── accounts/              # Users & roles
│   ├── karteien/              # Main registry
│   ├── approvals/             # Pending/declined workflow
│   ├── history/               # Change tracking
│   ├── notifications/         # User notifications
│   ├── legacy_import/         # Access migration
│   └── reporting/             # Reports & analytics
├── static/                    # Static files (CSS, JS, images)
└── templates/                 # Django templates
```

Docker-окружение для локальной разработки:

- `docker-compose.yml` — в корне репозитория
- `.env.example` — шаблон переменных окружения

### 4.2 App Structure

Каждый app хранится в `backend/apps/<name>/` со следующей структурой:

- `apps.py` — конфигурация приложения (name="apps.<appname>").
- `models.py` — только определения моделей.
- `services.py` / `use_cases.py` — бизнес‑логика (например, классификация риска, применение решения Superadmin).
- `api/` или `views.py` — HTTP‑слой (DRF views / serializers или Django views).
- `urls.py` — URL‑маршруты.
- `admin.py` — регистрация моделей в Django admin.
- `README.md` — краткое описание задач app, ключевых моделей и сервисов.

- Общие правила:
  - Не создавать большие монолитные файлы; по мере роста логики дробить `services` на отдельные модули.
  - Вся “настоящая” бизнес‑логика (аналог VBA) живёт в сервисах/доменных классах, а не в HTTP‑слое.
  - Каждое добавление новых сущностей/логики сопровождается обновлением:
    - `ARCHITECTURE.md` (если меняется общая архитектура),
    - `DOMAIN_MODEL.md` (структура данных),
    - `apps/<name>/README.md` (контекст конкретнего приложения).

Этот документ должен использоваться как основной контекст для агентных задач: перед началом работы Copilot Agent считывает его (и связанные README), вместо сканирования всего дерева исходников.
