# Karteien App

## Назначение

Приложение `karteien` — центральный модуль системы, отвечающий за хранение и управление основной картотекой семей и детей.

## Основные задачи

- **Хранение базовых данных** — FamilyID, Parent, Child, Birthdate, Address, контактные данные.
- **Предметы и цены** — Subject1/2, Price1/2, дополнительные предметы (Extra1-3).
- **Месячные поля** — 12 месяцев для отслеживания платежей/посещений.
- **CRUD-операции** — создание, редактирование, удаление записей Admin'ом.
- **Валидации** — проверка уникальности FamilyID+Parent, ограничения по прошлым месяцам.

## Ключевые модели

### KarteiRecord

Основная модель записи картотеки, соответствующая:

- Excel: лист `Kartei` (колонки A–AZ)
- Access: таблица `tblKartei` (поля ID, Value1–Value52)

**Ключевые поля:**

| Поле Django        | Excel Col | Access Field | Описание                          |
| ------------------ | --------- | ------------ | --------------------------------- |
| `id`               | AV (48)   | ID           | Уникальный ID записи (PK)         |
| `year`             | —         | —            | Год (заменяет отдельные файлы)    |
| `family_id`        | A (1)     | Value1       | Идентификатор семьи               |
| `parent_name`      | B (2)     | Value2       | Имя родителя                      |
| `child_name`       | D (4)     | Value4       | Имя ребёнка                       |
| `birthdate`        | E (5)     | Value5       | Дата рождения                     |
| `address`          | F (6)     | Value6       | Адрес                             |
| `phone`            | G (7)     | Value7       | Телефон                           |
| `mobile`           | H (8)     | Value8       | Мобильный                         |
| `email`            | I (9)     | Value9       | Email                             |
| `subject1`         | J (10)    | Value10      | Предмет 1                         |
| `price1`           | M (13)    | Value13      | Цена за предмет 1                 |
| `subject2`         | O (15)    | Value15      | Предмет 2                         |
| `price2`           | R (18)    | Value18      | Цена за предмет 2                 |
| `extra1..3`        | AK–AM     | Value37–39   | Доп. предметы                     |
| `month_1..12`      | U–AF      | Value21–32   | Месячные поля (январь–декабрь)    |
| `sepa_marker`      | AU (47)   | Value47      | Маркер SEPA                       |
| `status`           | BA (53)   | —            | Статус: '', 'PENDING', 'DECLINED' |
| `last_change_role` | AW (49)   | Value49      | Роль последнего редактора         |
| `last_change_date` | AX (50)   | Value50      | Дата последнего изменения         |
| `last_change_time` | AY (51)   | Value51      | Время последнего изменения        |
| `history_raw`      | AZ (52)   | Value52      | История изменений (текст)         |

**Константы (models.py):**

- `MONTH_FIELD_NAMES` — tuple имён месячных полей
- `TRACKED_FIELDS` — поля, участвующие в истории и риск-классификации
- `HISTORY_FIELD_TAGS` — теги для парсинга/построения истории

## Сервисный слой (services.py)

- `KarteiSyncService` — синхронизация записей (аналог ExportSyncKartei.bas)
- `KarteiValidationService` — валидация данных (уникальность, прошлые месяцы, SEPA)
- Вспомогательные функции: `get_records_by_year`, `get_pending_records`, `search_records`

## Связь с VBA-модулями

Переносит логику из:

- `ImportData`, `ExportSyncKartei`, `ExportUtilities`, `ExportProtection`
- `Export_ValidationKartei`, `Export_RiskClassification`
- Лист `Kartei` из Excel-файлов Admin
