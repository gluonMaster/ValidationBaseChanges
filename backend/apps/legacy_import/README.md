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
```

### Параметры команды

| Параметр            | Обязательный | Описание                                        |
| ------------------- | ------------ | ----------------------------------------------- |
| `--year`            | Да           | Год для импортируемых записей (2024, 2025, ...) |
| `--access-file`     | Да           | Имя файла или полный путь к .accdb              |
| `--dry-run`         | Нет          | Только анализ, без записи в БД                  |
| `--familyid-policy` | Нет          | `report` (по умолчанию) или `auto-merge`        |
| `--sync-history`    | Нет          | Синхронизировать history_raw → HistoryEvent     |
| `--skip-pending`    | Нет          | Пропустить импорт pre_tblKartei                 |
| `--skip-declined`   | Нет          | Пропустить импорт decl_tblKartei                |
| `--report-dir`      | Нет          | Директория для отчётов (по умолчанию `.`)       |

## Конфигурация

### Переменные окружения

```bash
# Шаблон строки подключения ODBC (опционально)
ACCESS_CONN_STRING_TEMPLATE="DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={file_path};"

# Базовый каталог с .accdb файлами
ACCESS_BASE_DIR=/opt/kindeltern_data
```

### settings.py

```python
# Можно также задать в settings.py
ACCESS_BASE_DIR = "/opt/kindeltern_data"
ACCESS_CONN_STRING_TEMPLATE = "DRIVER={...};DBQ={file_path};"
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
