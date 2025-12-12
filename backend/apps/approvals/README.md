# Approvals App

## Назначение

Приложение `approvals` реализует workflow утверждения/отклонения изменений и риск-классификацию.

## Основные задачи

- **Риск-классификация** — определение, какие изменения требуют одобрения Superadmin:
  - Политика `TRACKED_FIELDS` — изменения tracked-полей считаются рискованными.
- **Pending-изменения** — хранение изменений, ожидающих решения (аналог `pre_tblKartei`).
- **Declined-изменения** — хранение отклонённых записей (аналог `decl_tblKartei`).
- **Workflow утверждения**:
  - Admin вносит изменения → risky идут в pending.
  - Superadmin утверждает/отклоняет.
  - Approved → применяются к KarteiRecord.
  - Declined → создаётся запись declined + уведомление Admin.

## Ключевые модели

### `PendingChange`

Pending-изменение, ожидающее одобрения Superadmin.

| Поле           | Тип           | Описание                                           |
| -------------- | ------------- | -------------------------------------------------- |
| `id`           | BigAutoField  | Авто-генерируемый PK (не ID Kartei).               |
| `record`       | OneToOneField | Ссылка на `KarteiRecord` (один pending на запись). |
| `snapshot`     | JSONField     | Снимок tracked-полей: `{"field": value, ...}`.     |
| `is_processed` | BooleanField  | True после approve/decline.                        |
| `created_at`   | DateTimeField | Время создания.                                    |
| `updated_at`   | DateTimeField | Время последнего обновления.                       |

### `DeclinedChange`

Отклонённое изменение.

| Поле             | Тип              | Описание                                           |
| ---------------- | ---------------- | -------------------------------------------------- |
| `id`             | BigAutoField     | Авто-генерируемый PK.                              |
| `record`         | ForeignKey       | Ссылка на `KarteiRecord` (допускает многократные). |
| `snapshot`       | JSONField        | Снимок отклонённых значений.                       |
| `decline_reason` | TextField        | Причина отклонения от Superadmin.                  |
| `declined_by`    | ForeignKey(User) | Superadmin, отклонивший изменение.                 |
| `created_at`     | DateTimeField    | Время отклонения.                                  |
| `updated_at`     | DateTimeField    | Время последнего обновления.                       |

## Ключевые сервисы (`services.py`)

### Риск-классификация

- `is_risky_change(local, original) -> bool` — проверяет, изменились ли tracked-поля.
- `classify_change(local, original) -> Literal["SAFE", "RISKY"]` — обёртка.
- `get_changed_tracked_fields(local, original) -> dict` — список изменённых полей.

### Управление pending/declined

- `build_snapshot(record) -> dict` — строит JSON-снимок tracked-полей.
- `create_or_update_pending_change(record) -> PendingChange` — создаёт/обновляет pending.
- `create_declined_change(record, reason, declined_by) -> DeclinedChange` — создаёт declined.
- `apply_pending_change(pending) -> KarteiRecord` — применяет approved изменения.

### Superadmin Decision Workflow (PROMPT_09)

- `apply_decision(pending, decision, comment, user) -> KarteiRecord`:

  - Главный entry-point для обработки решений Superadmin.
  - `decision="APPROVED"`: применяет snapshot к записи, статус → NORMAL, пишет APR в историю.
  - `decision="DECLINED"`: создаёт DeclinedChange, статус → DECLINED, пишет DCL в историю.
  - Создаёт уведомления для Admin.

- `approve_all_pending(user, year) -> (count, errors)`:

  - Массовое одобрение всех pending-записей.

- `decline_all_pending(comment, user, year) -> (count, errors)`:

  - Массовое отклонение с общим комментарием.

- `clear_all_decisions(year) -> count`:
  - Сброс pending-записей (для тестирования).

### NeuList (Новые записи)

- `get_or_create_superadmin_state(user) -> SuperadminState`:

  - Получает или создаёт состояние Superadmin.

- `get_new_records(user, year) -> list[KarteiRecord]`:

  - Записи с ID > last_seen_id.

- `get_new_records_count(user, year) -> int`:

  - Количество новых записей.

- `update_last_seen_id(user, max_id=None)`:
  - Обновляет last_seen_id.

## Модели (PROMPT_09)

### `SuperadminState`

Состояние Superadmin для NeuList и других функций.

| Поле             | Тип                  | Описание                                |
| ---------------- | -------------------- | --------------------------------------- |
| `user`           | OneToOneField(User)  | Ссылка на Superadmin пользователя.      |
| `last_seen_id`   | PositiveIntegerField | Последний просмотренный ID для NeuList. |
| `last_seen_date` | DateTimeField        | Альтернативный трекер (дата).           |
| `updated_at`     | DateTimeField        | Время последнего обновления.            |

## URL-эндпоинты

### Admin UI

| URL                                   | View                      | Описание                   |
| ------------------------------------- | ------------------------- | -------------------------- |
| `/approvals/declined/`                | DeclinedOverviewView      | Список declined-записей.   |
| `/approvals/declined/<pk>/`           | DeclinedDetailView        | Детали declined.           |
| `/approvals/declined/<pk>/apply-fix/` | ApplyDeclinedFixView      | Применить исправление.     |
| `/approvals/declined/apply-all/`      | ApplyAllDeclinedFixesView | Применить все исправления. |
| `/approvals/pending/`                 | PendingChangesListView    | Список pending (info).     |
| `/approvals/pending/<pk>/`            | PendingDetailView         | Детали pending.            |

### Superadmin UI (PROMPT_09)

| URL                                   | View                          | Описание                    |
| ------------------------------------- | ----------------------------- | --------------------------- |
| `/approvals/superadmin/pending/`      | SuperadminPendingOverviewView | Pending overview для SA.    |
| `/approvals/superadmin/pending/<pk>/` | SuperadminWarIstView          | War/Ist comparison.         |
| `/approvals/superadmin/approve-all/`  | SuperadminApproveAllView      | Массовое одобрение.         |
| `/approvals/superadmin/decline-all/`  | SuperadminDeclineAllView      | Массовое отклонение.        |
| `/approvals/superadmin/neu/`          | SuperadminNeuListView         | NeuList (новые записи).     |
| `/approvals/superadmin/mark-seen/`    | SuperadminMarkSeenView        | Отметить как просмотренные. |
| `/approvals/superadmin/history/<pk>/` | SuperadminRecordHistoryView   | История записи.             |

## Связь с VBA-модулями

Переносит логику из:

- `Export_RiskClassification` — `IsRiskyChange`, `HasTrackedFieldChanges`
- `Export_OverlayPending` — работа с pending-записями
- `Export_DeclinedTools`, `Export_DeclinedHelpers` — работа с declined
- `valid_ImportPending`, `valid_ApproveFlow` — workflow approve/decline
