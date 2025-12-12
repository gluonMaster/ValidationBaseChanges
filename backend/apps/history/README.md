# History App

## Назначение

Приложение `history` отвечает за хранение и отображение истории изменений записей картотеки.

## Основные задачи

- **Хранение истории изменений** — кто, когда, какие поля изменены, старые/новые значения.
- **Поддержка текстового формата** — совместимость с текущим форматом `Value52` для миграции.
- **Представление истории записи** — интерактивная лента изменений (аналог листа `Geschichte`).
- **Агрегированные отчёты** — War/Ist сравнение для множества записей (аналог `grossGeschichte`).

## Ключевые сущности

### HistoryEvent

Нормализованное событие изменения:

| Поле                   | Тип               | Описание                                              |
| ---------------------- | ----------------- | ----------------------------------------------------- |
| `id`                   | BigAutoField      | Первичный ключ                                        |
| `record`               | FK → KarteiRecord | Связь с записью картотеки                             |
| `user`                 | FK → User (null)  | Пользователь, сделавший изменение                     |
| `event_time`           | DateTimeField     | Время события                                         |
| `event_type`           | CharField         | Тип: CHANGE, CREATE, APPROVE, DECLINE, IMPORT, RUCK   |
| `changes`              | JSONField         | Структура: `{"field_name": {"old": ..., "new": ...}}` |
| `comment`              | TextField         | Комментарий к изменению                               |
| `raw_history_fragment` | TextField         | Исходный текст из AZ/Value52                          |

### EventType

Типы событий:

- `CHANGE` — обычное изменение tracked-полей
- `CREATE` — создание новой записи
- `APPROVE` — одобрение pending-изменения Superadmin'ом
- `DECLINE` — отклонение изменения (тег DCL в истории)
- `IMPORT` — импорт из legacy-данных
- `RUCK` — ретроактивное изменение (префикс RUCK: в истории)

## Сервисы (services.py)

### parse_raw_history(raw: str) → list[HistoryEventData]

Парсит сырую строку истории (AZ/Value52) в список структурированных событий.

Поддерживает оба формата:

- **Новый**: `[RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||`
- **Legacy**: `Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY ||`

### sync_history_from_raw(record: KarteiRecord) → list[HistoryEvent]

Синхронизирует историю из поля `history_raw` записи в нормализованные `HistoryEvent`.
Создаёт только отсутствующие события, не изменяет исходную строку.

## API Endpoints

| Метод | URL                               | Описание                     |
| ----- | --------------------------------- | ---------------------------- |
| GET   | `/api/history/records/<id>/`      | История событий для записи   |
| GET   | `/api/history/events/<id>/`       | Детали одного события        |
| POST  | `/api/history/records/<id>/sync/` | Синхронизация истории из raw |

### Параметры запроса (records/<id>/)

- `limit` — макс. количество событий (default: 100)
- `offset` — смещение для пагинации
- `event_type` — фильтр по типу события

## Связь с VBA-модулями

Переносит логику из:

- `Export_HistoryBuilder` — константы тегов, формат строки
- `Export_HistoryParser` — парсинг обоих форматов
- `Export_HistoryConverter` — преобразование форматов
- `Geschichte`, `grossGeschichte` — UI отображение истории
