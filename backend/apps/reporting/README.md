# Reporting App

## Назначение

Приложение `reporting` отвечает за формирование сложных отчётов и аналитики.

## Основные задачи

- **Комплексные отчёты** — отчёты, не укладывающиеся в простые списки.
- **Dashboards** — панели управления для руководства.
- **Аналитика** — агрегированные данные по картотеке, истории, approvals.
- **Последние изменения** — обзор недавних изменений (аналог `grossGeschichte`).

## API Endpoints

| Метод | URL                                        | Описание                              |
| ----- | ------------------------------------------ | ------------------------------------- |
| GET   | `/api/reporting/recent-changes/`           | Список последних изменений            |
| GET   | `/api/reporting/recent-changes-by-record/` | Изменения, сгруппированные по записям |

### GET /api/reporting/recent-changes/

Возвращает список последних событий истории со всех записей.

**Параметры запроса:**

| Параметр     | Тип | Default | Описание                                   |
| ------------ | --- | ------- | ------------------------------------------ |
| `limit`      | int | 50      | Макс. количество событий (max: 500)        |
| `offset`     | int | 0       | Смещение для пагинации                     |
| `year`       | int | —       | Фильтр по году записи                      |
| `date_from`  | str | —       | Фильтр по дате начала (YYYY-MM-DD)         |
| `date_to`    | str | —       | Фильтр по дате окончания (YYYY-MM-DD)      |
| `event_type` | str | —       | Фильтр по типу события                     |
| `family_id`  | str | —       | Фильтр по Family ID (частичное совпадение) |

**Ответ:**

```json
{
  "items": [
    {
      "event_id": 123,
      "event_time": "2025-12-01T14:30:00Z",
      "event_type": "CHANGE",
      "event_type_display": "Field Change",
      "changes": { "month_1": { "old": "100", "new": "200" } },
      "comment": "Monthly update",
      "record_id": 456,
      "year": 2025,
      "family_id": "FAM001",
      "parent_name": "Müller, Hans",
      "child_name": "Müller, Anna",
      "user_name": "admin"
    }
  ],
  "total_count": 150,
  "limit": 50,
  "offset": 0,
  "filters": { "year": 2025 }
}
```

### GET /api/reporting/recent-changes-by-record/

Возвращает записи с последними изменениями, сгруппированные по записям.

**Параметры запроса:**

| Параметр    | Тип | Default | Описание                            |
| ----------- | --- | ------- | ----------------------------------- |
| `limit`     | int | 50      | Макс. количество записей (max: 200) |
| `year`      | int | —       | Фильтр по году                      |
| `date_from` | str | —       | Фильтр по дате начала               |
| `date_to`   | str | —       | Фильтр по дате окончания            |

## Источники данных

Использует данные из:

- `karteien` — записи картотеки
- `history` — события изменений (HistoryEvent)
- `approvals` — pending/declined записи

## Связь с VBA-модулями

Аналог функционала:

- `grossGeschichte` — обзор изменений по множеству записей
- `valid_GrossGeschichteDecision` — War/Ist сравнение для Superadmin
