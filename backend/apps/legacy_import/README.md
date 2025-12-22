# Legacy Import App

## Назначение

Приложение `legacy_import` — модуль для миграции данных из существующих
Access-баз (`KindElternDaten_XX_front.accdb`) в Django/Postgres.

## Основные задачи

- **Чтение Access-файлов** через pyodbc/ACE ODBC
- **Импорт данных** в Django-моделях:
  - `tblKartei` → `karteien.KarteiRecord`
  - `pre_tblKartei` → `approvals.PendingChange` + `status=PENDING`
  - `decl_tblKartei` → `approvals.DeclinedChange` + `status=DECLINED`
  - `Value52` → `KarteiRecord.history_raw` → (опционально) `history.HistoryEvent`
- **Фильтрация маркерных строк** — строки "Zahlung" (семейные разделители) не импортируются
- **Анализ FamilyID** — обнаружение и отчёты по несогласованным FamilyID
- **Patch-режим** — дополнение уже импортированных записей новыми полями
- **Management command** — `python manage.py import_access_year`

## Структура модуля

```
legacy_import/
├── __init__.py
├── access_client.py       # Подключение к Access и чтение таблиц
├── services.py            # Бизнес-логика импорта
├── models.py              # (заглушка, модели не требуются)
├── admin.py
├── apps.py
├── views.py
├── urls.py
├── README.md
└── management/
    └── commands/
        └── import_access_year.py  # Management command
```

## Использование

### Management Command

```bash
# Полный импорт
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb

# Dry run (анализ без записи)
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --dry-run

# С авто-нормализацией FamilyID
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --familyid-policy=auto-merge

# С синхронизацией истории
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --sync-history

# Patch-режим: дополнить существующие записи данными учителей и контракта
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --patch-fields

# Patch-режим с dry-run (только статистика, без записи)
python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --patch-fields \
    --dry-run
```

### Параметры команды

| Параметр            | Обязательный | Описание                                                                         |
| ------------------- | ------------ | -------------------------------------------------------------------------------- |
| `--year`            | Да           | Год для импортируемых записей (2024, 2025, ...)                                  |
| `--access-file`     | Да           | Имя файла или полный путь к .accdb                                               |
| `--dry-run`         | Нет          | Только анализ, без записи в БД                                                   |
| `--patch-fields`    | Нет          | Patch-режим: дополнить существующие записи (учитель, тип/статус контракта, SEPA) |
| `--familyid-policy` | Нет          | `report` (по умолчанию) или `auto-merge`                                         |
| `--sync-history`    | Нет          | Синхронизировать history_raw → HistoryEvent                                      |
| `--skip-pending`    | Нет          | Пропустить импорт pre_tblKartei                                                  |
| `--skip-declined`   | Нет          | Пропустить импорт decl_tblKartei                                                 |
| `--report-dir`      | Нет          | Директория для отчётов (по умолчанию `.`)                                        |

## Patch-режим (`--patch-fields`)

### Предусловия (важно)

Перед запуском убедитесь, что миграции применены к этой базе данных, иначе будет ошибка вида:
`UndefinedColumn ... teacher1_legacy_name does not exist`.

- если Django работает в Docker: `docker compose exec web python manage.py migrate`
- если импорт запускаете локальным Python: `python manage.py migrate` (при том же `DATABASE_URL`)

Patch-режим позволяет **дополнить** уже импортированные записи недостающими
полями **без полного реимпорта**.

### Зачем нужен

- Добавить данные о учителях (`Value11`, `Value16`) к существующим записям
- Добавить маркеры типа контракта (`Value14` → `contract_type_raw`, `is_monthly_contract`)
- Добавить статус контракта (`Value20` → `contract_status_raw`, `is_contract_terminated`)
- Обновить `sepa_marker` (`Value47`)

### Что делает

1. Читает **только** `tblKartei` (игнорирует `pre_tblKartei`/`decl_tblKartei`)
2. Для каждой строки ищет существующую запись по `(year, id)`
3. Обновляет **только** patch-поля, не затрагивая остальные данные
4. Поддерживает `--dry-run` для предварительной проверки

### Обновляемые поля

| Access поле | Django поле              | Описание                                   |
| ----------- | ------------------------ | ------------------------------------------ |
| Value11     | `teacher1_legacy_name`   | Учитель 1-го полугодия (текст)             |
| Value16     | `teacher2_legacy_name`   | Учитель 2-го полугодия (текст)             |
| Value14     | `contract_type_raw`      | Тип контракта (сырой текст)                |
| —           | `is_monthly_contract`    | Вычисляется: содержит ли `O/V` (bool)      |
| Value20     | `contract_status_raw`    | Статус контракта (сырой текст)             |
| —           | `is_contract_terminated` | Вычисляется: содержит ли токен `KN` (bool) |
| Value47     | `sepa_marker`            | SEPA маркер                                |

### Правила вычисления флагов

- **`is_monthly_contract`**: `True` если `contract_type_raw` содержит подстроку `O/V`
  (case-insensitive, в любом месте текста)
- **`is_contract_terminated`**: `True` если `contract_status_raw` содержит токен `KN`
  как отдельное слово (с пробелом или на границе строки)

### Отчёт

После выполнения создаётся файл `patch_stats_YEAR.json`:

```json
{
  "timestamp": "2025-01-15T10:30:00",
  "year": 2025,
  "dry_run": false,
  "mode": "patch-fields",
  "stats": {
    "total_rows": 1500,
    "records_found": 1480,
    "records_updated": 1480,
    "records_not_found": 15,
    "marker_skipped": 5,
    "parse_errors": 0
  }
}
```

## Конфигурация

### Переменные окружения

```bash
# Шаблон строки подключения ODBC (опционально)
ACCESS_CONN_STRING_TEMPLATE="DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};"

# Базовый каталог с .accdb файлами
ACCESS_BASE_DIR=/opt/kindeltern_data
```

### settings.py

```python
# Можно также задать в settings.py
ACCESS_BASE_DIR = "/opt/kindeltern_data"
ACCESS_CONN_STRING_TEMPLATE = "DRIVER={{...}};DBQ={file_path};"
```

## Ключ записи: (year, ID)

- **ID** — числовой идентификатор из Access (`tblKartei.ID`, Excel: AV/48)
- **year** — год, передаваемый как параметр команды
- Уникальное ограничение: `(year, id)` в `KarteiRecord`
- Один и тот же ID может существовать в разных годах
- `conflicts_year_id` — количество обнаруженных дублей `(year, id)` в данных
  `tblKartei` в рамках одного запуска импорта (такие строки пропускаются)

## Маркерные строки "Zahlung"

Маркерные строки используются в Excel для визуального разделения семей:

- `Value4` (Child) = " Zahlung" (с пробелом)
- Остальные поля пусты
- Строка залита зелёным

Эти строки **не импортируются** в Django. Количество пропущенных маркеров
отражается в статистике `marker_skipped`.

## Анализ FamilyID

### Обнаруживаемые проблемы

1. **SAME_FAMILY_DIFFERENT_FAMILYIDS** — одна семья (по `parent_name` + `email`)
   имеет разные FamilyID
2. **SAME_FAMILYID_DIFFERENT_FAMILIES** — один FamilyID используется разными
   семьями

**Границы анализа:** FamilyID-анализ в management-команде работает **per-year** —
в рамках одного года, заданного параметром `--year`.

### Политики обработки

- `report` (по умолчанию) — только отчёт, без изменений
- `auto-merge` — для случая #1 автоматически выбирается канонический FamilyID

### Отчёты

- `import_familyid_issues_YEAR.csv` — обнаруженные проблемы
- `import_familyid_merge_YEAR.csv` — маппинг `old_family_id → new_family_id`
- `import_stats_YEAR.json` — полная статистика импорта

## История (history_raw)

- `Value52` из Access сохраняется в `KarteiRecord.history_raw` без изменений
- При `--sync-history` вызывается `history.sync_history_from_raw()` для
  создания `HistoryEvent` из сырых данных
- Существующие `HistoryEvent` не перезаписываются (односторонняя синхронизация)

## Связь с VBA-модулями

Переносит логику из:

- `AccessCreation.bas` — структура таблиц Access
- `DumpBase.bas` — выгрузка данных
- `BaseBackupRestore.bas` — бэкап/восстановление
- `ImportData.bas` — импорт в Excel

## Примечание

Это модуль миграции. После полного перехода на веб-систему используется
только для:

- Первоначального импорта исторических данных
- Импорта данных за новый год (если Excel+Access ещё используется параллельно)
