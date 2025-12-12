# Notifications App

## Назначение

Приложение `notifications` отвечает за систему внутренних уведомлений пользователей.

## Основные задачи

- **Уведомления для Superadmin** — новые pending-записи, требующие рассмотрения.
- **Уведомления для Admin** — отклонённые записи, требующие исправления.
- **Управление статусом** — отметка прочитанных уведомлений.

## Типы уведомлений

| Тип                | Описание                     | Получатели     |
| ------------------ | ---------------------------- | -------------- |
| `PENDING_CREATED`  | Создана новая pending-запись | Все Superadmin |
| `DECLINED_CREATED` | Запись отклонена             | Все Admin      |

## Ключевые модели

### Notification

- `id` — BigAutoField (PK)
- `recipient` — FK на accounts.User (получатель)
- `type` — CharField (PENDING_CREATED, DECLINED_CREATED)
- `record` — FK на karteien.KarteiRecord (nullable)
- `payload` — JSONField (дополнительные данные)
- `created_at` — DateTimeField (auto_now_add)
- `read_at` — DateTimeField (nullable, дата прочтения)

## Сервисы (services.py)

### Создание уведомлений

- `notify_pending_created(record, pending)` — уведомляет всех Superadmin о новой pending-записи
- `notify_declined_created(record, declined)` — уведомляет всех Admin об отклонённой записи

### Чтение уведомлений

- `get_unread_count(user)` — количество непрочитанных уведомлений
- `get_notifications(user, unread_only, notification_type, limit)` — список уведомлений с фильтрацией
- `mark_notification_read(notification_id, user)` — пометить как прочитанное
- `mark_all_notifications_read(user)` — пометить все как прочитанные

## API Endpoints

| Метод | URL                                | Описание                                        |
| ----- | ---------------------------------- | ----------------------------------------------- |
| GET   | `/api/notifications/`              | Список уведомлений (query: unread, type, limit) |
| GET   | `/api/notifications/unread-count/` | Количество непрочитанных                        |
| POST  | `/api/notifications/<id>/read/`    | Пометить как прочитанное                        |
| POST  | `/api/notifications/read-all/`     | Пометить все как прочитанные                    |

## UI

Базовый шаблон `templates/base.html` включает:

- "Колокольчик" в навигации с бейджем непрочитанных
- Dropdown со списком последних уведомлений
- Автоматическое обновление каждые 60 секунд

## Интеграция с approvals

Уведомления создаются автоматически:

- `create_or_update_pending_change()` → `notify_pending_created()`
- `create_declined_change()` → `notify_declined_created()`

## Идемпотентность

Сервисы не создают дубликаты: если непрочитанное уведомление для того же record + recipient + type уже существует, новое не создаётся.

## Будущие улучшения

- E-mail уведомления
- Telegram интеграция
- WebSocket для real-time обновлений
- Уведомление конкретному автору изменения (вместо всех Admin)
