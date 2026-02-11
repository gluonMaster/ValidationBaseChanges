# DOMAIN_MODEL – KindEltern Data Model

Примечание по исходникам VBA: все выгруженные модули Excel/VBA (для справки и сопоставления логики) находятся в `legacy_VBA/`:

- `legacy_VBA/*.bas` и `legacy_VBA/*.cls` — модули из `KindElternDaten_XX_Admin.xlsm`
- `legacy_VBA/admin_forms/` — VBA-код форм из `KindElternDaten_XX_Admin.xlsm`
- `legacy_VBA/Superadmin/` — модули из `KindElternDaten_XX_Suprime.xlsm`
- `legacy_VBA/alt/` — модули из `KindElternDaten_XX_Data.xlsm`

Этот документ описывает **фактическую** модель данных системы в текущей реализации на Excel+Access. Он служит источником истины при проектировании Django‑моделей и миграции.

Примечание для нового чата: чтобы экономить контекст, начните с `docs/ProjectMap.md`, а здесь открывайте только нужные разделы (через поиск по файлу).

## 0. Дополнения web-реализации (не в legacy)

В Django-версии проекта есть сущности, которые **не имеют прямых таблиц/колонок** в Excel/Access, но нужны для реализации ТЗ (например, прайслист v2 и вспомогательные справочники). Ключевые примеры:

- `catalog.SemesterConfig` — граница семестра по календарному году (влияет на разбиение "1/2 полугодие" в биллинге).
- `catalog.SubjectCategory` + `catalog.SubjectCategoryLink` — категории предметов на год (GROUP/INDIVIDUAL) и привязка `Subject`↔`Category`.
- `catalog.DisciplineGroup` + `DurationEntry` + `GroupSizeEntry` — группы дисциплин по предмету/году и их месячная история (длительность/размер).

Эти модели описаны в `ARCHITECTURE.md` (модуль `catalog`) и используются UI `/catalog/...`.

Основные сущности:

- Excel‑лист `Kartei` (файл Admin: `KindElternDaten_XX_Admin.xlsm`).
- Access‑база `KindElternDaten_XX_front.accdb` с таблицами:
  - `tblKartei` – основная таблица (source of truth).
  - `pre_tblKartei` – pending (ожидающие одобрения) изменения.
  - `decl_tblKartei` – declined (отклонённые) изменения.

Все сопоставления делаются по **ID** (числовое поле в Access и колонка `Kartei!AV`).

---

## 1. Kartei Sheet Layout (Excel)

Лист `Kartei` – главный пользовательский интерфейс Admin/Operator.

Ниже структура по колонкам (индексы Excel/Access начинаются с 1).

### 1.1 Основные идентификаторы и статус

| Col | Index | Name (условное)  | Тип    | Хранение в Access       | Описание / правила                                                                                                                      |
| --- | ----- | ---------------- | ------ | ----------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| A   | 1     | `FamilyID`       | string | `Value1`                | Идентификатор семьи (группирует детей/записи). Не уникален по таблице, но используется логически.                                       |
| AV  | 48    | `ID`             | int    | **отдельное поле `ID`** | Уникальный числовой ID записи. Ключ для `tblKartei/pre/decl`. Совпадает с `tblKartei.ID`.                                               |
| BA  | 53    | `Status`         | string | Excel‑только            | Статус строки на листе: `""` (обычная), `"PENDING"`, `"DECLINED"`. Выводится на основе наличия записи в `pre_tblKartei/decl_tblKartei`. |
| AT  | 46    | `StatusCopyPrev` | string | `Value46`               | Используется для копии статуса BA (Workbook_BeforeClose копирует BA→AT). Нужен для диагностики/сравнений между сессиями.                |

Примечания по ID:

- `ID` уникален в пределах одной базы (одного года).
- `Kartei!AV` всегда должен совпадать с `ID` во всех трёх таблицах (`tblKartei`, `pre_tblKartei`, `decl_tblKartei`) для данной записи.

### 1.2 Базовая информация о семье/ребенке

| Col | Idx | Name        | Тип    | Access `ValueN` | Описание                                                            |
| --- | --- | ----------- | ------ | --------------- | ------------------------------------------------------------------- |
| A   | 1   | `FamilyID`  | string | `Value1`        | Идентификатор семьи.                                                |
| B   | 2   | `Parent`    | string | `Value2`        | Имя родителя (Eltern). Используется при поиске дубликатов FamilyID. |
| D   | 4   | `Child`     | string | `Value4`        | Имя ребенка (Kind).                                                 |
| E   | 5   | `Birthdate` | date   | `Value5`        | Дата рождения.                                                      |
| F   | 6   | `Address`   | string | `Value6`        | Адрес.                                                              |
| G   | 7   | `Phone`     | string | `Value7`        | Телефон (как текст, сохранение ведущих нулей).                      |
| H   | 8   | `Mobile`    | string | `Value8`        | Мобильный (как текст, с ведущими нулями).                           |
| I   | 9   | `Email`     | string | `Value9`        | Email.                                                              |

Колонка C (3) используется визуально (цвет) и в истории, но её бизнес‑смысл **в этой модели не фиксируем**, считаем “extra/visual”.

### 1.3 Предметы и цены

| Col | Idx | Name       | Тип    | Access `ValueN` | Описание            |
| --- | --- | ---------- | ------ | --------------- | ------------------- |
| J   | 10  | `Subject1` | string | `Value10`       | Основной предмет 1. |
| M   | 13  | `Price1`   | number | `Value13`       | Цена за предмет 1.  |
| O   | 15  | `Subject2` | string | `Value15`       | Предмет 2.          |
| R   | 18  | `Price2`   | number | `Value18`       | Цена за предмет 2.  |
| AK  | 37  | `Extra1`   | string | `Value37`       | Доп. предмет 1.     |
| AL  | 38  | `Extra2`   | string | `Value38`       | Доп. предмет 2.     |
| AM  | 39  | `Extra3`   | string | `Value39`       | Доп. предмет 3.     |

Остальные промежуточные колонки (K, L, N, P, Q, S, T, AG–AJ, AN–AS) используются для вспомогательных данных/форматирования и **не входят** в набор “tracked‑полей” истории.

### 1.4 Месячные поля (U–AF)

Колонки U–AF (21–32) — 12 месяцев (обычно январь–декабрь). Они участвуют в:

- Правилах блокировки по прошлым месяцам (`ExportProtection.ValidateAndFixPastMonths`).
- Истории (tracked‑поля).
- Риск‑классификации.

| Col Range | Indices | Name           | Тип    | Access `ValueN`     | Описание                                        |
| --------- | ------- | -------------- | ------ | ------------------- | ----------------------------------------------- |
| U–AF      | 21–32   | `Months[1-12]` | number | `Value21`–`Value32` | Значения по месяцам (платежи/посещения и т.п.). |

### 1.5 Технические служебные колонки

| Col | Idx | Name             | Тип    | Access `ValueN` | Описание                                                                            |
| --- | --- | ---------------- | ------ | --------------- | ----------------------------------------------------------------------------------- |
| AU  | 47  | `SepaMarker`     | string | `Value47`       | Маркер SEPA (`"SEPA"`). Для строк с SEPA действуют ограничения для Operator.        |
| AW  | 49  | `LastChangeRole` | string | `Value49`       | Роль пользователя, сделавшего последнюю запись (`"Admin"`/`"Operator"`).            |
| AX  | 50  | `LastChangeDate` | date   | `Value50`       | Дата последнего изменения (дата синхронизации).                                     |
| AY  | 51  | `LastChangeTime` | time   | `Value51`       | Время последнего изменения.                                                         |
| AZ  | 52  | `History`        | string | `Value52`       | История изменений в текстовом формате (новый формат через `Export_HistoryBuilder`). |

Информация о цветах:

- Цвет фона ячеек A–AY (1–51) хранится в Access полях `InteriorColor1..InteriorColor51`.
- Цвет шрифта для C (3) и R (18) хранится в `FontColor3` и `FontColor18`.

---

## 2. Access Tables

Во всех годовых базах (`KindElternDaten_24_front.accdb`, `KindElternDaten_25_front.accdb` и т.п.) используются три основные таблицы одинаковой структуры.

### 2.1 Общая структура (tblKartei, pre_tblKartei, decl_tblKartei)

**tblKartei** — основная таблица (источник истины).  
**pre_tblKartei** — pending изменения (ожидают решения Superadmin).  
**decl_tblKartei** — отклонённые записи.

Структура таблиц (согласно `AccessCreation.CreateTableKartei` и аналогичным процедурам):

- `ID` – `Long`, Primary Key (не AutoNumber; значение задаётся из Excel ID/AV).
- Для каждого `col` от 1 до 51:
  - `Value<col>` – `Text(255)`, допускает пустые строки.
  - `InteriorColor<col>` – `Long`, цвет фона.
- `FontColor3` – `Long`, цвет шрифта для колонки C (3).
- `FontColor18` – `Long`, цвет шрифта для колонки R (18).
- `Value52` – `Memo`, история (соответствует AZ/52 на листе).

**Важно:**

- Для колонки AV (48) на листе:
  - В Access есть поля `Value48` и `ID`.
  - В актуальной логике:
    - `ID` используется как единственный источник истины по идентификатору.
    - При импорте в Excel колонка AV заполняется из поля `ID`, поле `Value48` игнорируется.
    - При записи обратно в Access:
      - `ID` берётся из колонки AV.
      - Остальные `ValueN` (кроме 48) заполняются по колонкам 1–51.
- Таблицы `pre_tblKartei` и `decl_tblKartei` создаются/поддерживаются кодом:
  - `Export_RiskClassification.EnsurePreTableExists` (для `pre_tblKartei`).
  - `Export_RiskClassification.CreateDeclTable` и функции в `valid_ApproveFlow` (для `decl_tblKartei`).

### 2.2 Семантика таблиц

- `tblKartei`:

  - Всегда содержит **текущее утверждённое состояние** записей.
  - При safe‑изменениях (без участия Superadmin) записи обновляются напрямую.
  - При approve Superadmin’ом pending‑изменения переносятся из `pre_tblKartei` в `tblKartei` (update/insert).

- `pre_tblKartei`:

  - Содержит одно “pending состояние” записи для каждого ID, где изменение классифицировано как risky.
  - При повторных risky‑изменениях по этому же ID запись в `pre_tblKartei` обновляется.
  - При approve → запись удаляется из `pre_tblKartei`, изменения переносятся в `tblKartei`.
  - При decline → запись удаляется из `pre_tblKartei` и переносится в `decl_tblKartei` (с записью DCL в историю).

- `decl_tblKartei`:
  - Содержит отклонённые записи (после решения Superadmin).
  - Admin может исправлять такие записи напрямую в Excel (Kartei) или через лист `DeclinedOverview`.
  - После исправления и синхронизации запись переносится обратно в `pre_tblKartei`, и статус меняется с DECLINED на PENDING.

---

## 3. Roles & Permissions (Data-Level View)

Роли:

- `Admin`:
  - Может менять любые поля записи (FamilyID, Parent, months, SEPA и т.д.).
  - Может создавать новые записи (новые ID).
  - Видит и может править PENDING/DECLINED записи (с учётом логики sync).
- `Operator`:
  - Ограничен в изменении прошлых месяцев (U–AF) и SEPA‑строк:
    - **VBA реализация:** `ExportProtection.ValidateAndFixPastMonths` и `ValidateOperatorSepaRestrictions` (ExportSyncKartei).
    - **Django реализация:** `apps/karteien/validators.py` — серверная валидация:
      - `validate_sepa_restrictions()` — запрет редактирования SEPA‑строк для Operator.
      - `validate_past_months_restrictions()` — запрет изменения прошлых месяцев.
      - `validate_operator_changes()` — комплексная проверка всех ограничений.
  - Не может изменять строки с `SepaMarker="SEPA"` (изменения откатываются).
- `Superadmin`:
  - Не редактирует Kartei напрямую.
  - Работает с pending/declined через свои листы (`Kartei`, `GrossGeschichte`, `Geschichte`) в Superadmin‑файле, но все изменения фиксируются в Access через `pre_tblKartei/decl_tblKartei/tblKartei`.
- `User`:
  - **Только read-only доступ** (без возможности изменения данных).
  - Может искать записи по FamilyID, Parent, Child, Year.
  - Может просматривать детали записи (все поля, статус).
  - Может просматривать историю изменений записи (аналог Geschichte).
  - **Запрещено**: редактирование, создание, удаление записей.
  - **Запрещено**: доступ к approvals workflow (pending/declined).
  - Серверная защита: `UserRoleMixin` блокирует все POST/PUT/DELETE операции.

### 3.1 Server-Side Validation (Django Web UI)

В веб‑интерфейсе Admin/Operator ограничения реализованы на **серверной стороне**:

| Валидатор                           | Описание                                         | Роль     |
| ----------------------------------- | ------------------------------------------------ | -------- |
| `validate_sepa_restrictions`        | Запрет редактирования строк с `sepa_marker`      | Operator |
| `validate_past_months_restrictions` | Запрет изменения прошлых месяцев                 | Operator |
| `validate_familyid_duplicate`       | Проверка дубликатов `FamilyID + parent_name`     | Все роли |
| `validate_familyid_nonempty_parent` | Запрет пустого `family_id` при непустом `parent` | Все роли |

Формы (`KarteiRecordForm`) автоматически отключают поля для Operator:

- Все месячные поля (`month_01`–`month_12`) делаются `disabled` если месяц прошёл.
- Для SEPA‑записей все поля (кроме `status`) становятся `disabled`.

Статусы по данным:

- **Normal**:
  - Строка присутствует только в `tblKartei`.
  - `Kartei!BA` пусто, колонка A не подсвечена.
  - **В веб-интерфейсе**: Admin/Operator могут редактировать запись через `/karteien/<id>/edit/`. SAFE-изменения применяются напрямую.
- **PENDING**:
  - По ID существует запись в `pre_tblKartei`.
  - `Kartei!BA = "PENDING"`, колонка A залита голубым (кроме случаев со special value `"Zahlung"` в D).
  - **В веб-интерфейсе**: Редактирование через Kartei UI заблокировано. Изменения возможны только через approvals-флоу (решение Superadmin).
- **DECLINED**:
  - По ID существует запись в `decl_tblKartei`.
  - `Kartei!BA = "DECLINED"`, колонка A залита красным.
  - **В веб-интерфейсе**: Редактирование через Kartei UI заблокировано. Для исправления используется DeclinedOverview (`/approvals/declined/`).

**Важно**: SAFE-применение изменений (без создания `PendingChange`) разрешено только для записей со статусом NORMAL. Для записей PENDING/DECLINED любые изменения tracked-полей должны пройти через approvals-флоу.

---

## 4. Tracked vs Non-Tracked Fields (History & Risk)

История пишется не по всем колонкам. Политика `TRACKED_FIELDS`:

**Tracked поля** (участвуют в истории и считаются “risk” при изменении):

- A (1) – `FamilyID`
- B (2) – `Parent`
- D (4) – `Child`
- E (5) – `Birthdate`
- F (6) – `Address`
- G (7) – `Phone`
- H (8) – `Mobile`
- I (9) – `Email`
- J (10) – `Subject1`
- M (13) – `Price1`
- O (15) – `Subject2`
- R (18) – `Price2`
- U–AF (21–32) – `Months[1..12]`
- AK (37), AL (38), AM (39) – `Extra1..3`

**Non-tracked поля** (изменение не попадает в историю и обычно считается safe):

- Все прочие колонки (C, K, L, N, P, Q, S, T, AG–AJ, AN–AS, AU, AW–AY, BA).
- Примеры:
  - `SepaMarker` (AU) — влияет на права, но не фиксируется в истории.
  - `LastChangeRole/Date/Time` — служебные.

Риск‑классификация (`Export_RiskClassification.IsRiskyChange` при `RISK_POLICY_MODE = "TRACKED_FIELDS"`):

- Для существующих ID:
  - Если меняется хотя бы одно tracked‑поле → изменение считается risky → идёт в `pre_tblKartei`.
  - Если меняются только non‑tracked поля → изменение safe → можно обновлять `tblKartei` напрямую.
- Для новых ID:
  - Всегда считаются safe (нет исходной записи в `tblKartei`).

---

## 5. History Field (AZ / Value52)

Поле истории:

- Excel: колонка AZ (52) на листе `Kartei`.
- Access: поле `Value52` в таблицах `tblKartei`, `pre_tblKartei`, `decl_tblKartei`.

Формат:

- Сейчас используется **новый, структурированный формат**, формируемый модулем `Export_HistoryBuilder` и частично преобразуемый `Export_HistoryConverter`.
- История представляет собой конкатенацию “сессий” с разделителем `HD_SESSION = "||"`, каждая сессия содержит:
  - изменения по полям (tags вроде `FID`, `PAR`, `CHD`, `M01..M12`, `EX1..3`),
  - комментарий в формате `/@...@/`,
  - дату.

Для этой модели достаточно знать:

- История — **одна строка на запись**, накапливающаяся по мере изменений.
- Она должна переноситься intact при миграции; дополнительная нормализация истории (как отдельные строки `HistoryEvent`) — задача следующего уровня.

### 5.1 Mapping to Django history

Поле `Value52` (AZ) является исходным полем для парсинга и нормализации истории:

- **Исходные данные**: `KarteiRecord.history_raw` (TextField) хранит сырую строку из AZ/Value52.
- **Нормализация**: Сервис `history.services.parse_raw_history(raw)` парсит строку и возвращает список `HistoryEventData`.
- **Хранение**: Каждая "сессия" (сегмент между `||`) становится отдельной записью `history.models.HistoryEvent`.

Структура `HistoryEvent`:

| Поле                   | Источник                                                                     |
| ---------------------- | ---------------------------------------------------------------------------- |
| `event_time`           | Парсится из даты в конце сегмента (DD.MM.YYYY)                               |
| `event_type`           | Определяется по тегам: `DCL` → DECLINE, `RUCK:` → RUCK, иначе CHANGE         |
| `changes`              | Словарь: `{"field_name": {"old": ..., "new": ...}}` из тегов `TAG(OLD->NEW)` |
| `comment`              | Парсится из `/@...@/`                                                        |
| `raw_history_fragment` | Исходный сегмент строки для отладки                                          |

Синхронизация:

- `history.services.sync_history_from_raw(record)` — парсит `history_raw` и создаёт отсутствующие `HistoryEvent`.
- Существующие события не изменяются (односторонняя совместимость).
- Исходная строка `history_raw` остаётся "источником истины" до полного перехода.

---

## 6. Годовой контекст

Каждый год сейчас представлен **отдельной** связкой файлов:

- `KindElternDaten_24_Admin.xlsm`, `KindElternDaten_24_Suprime.xlsm`, `KindElternDaten_24_Data.xlsm`, `KindElternDaten_24_front.accdb`.
- `KindElternDaten_25_...` и т.д.

Внутри каждого года:

- Структура листа `Kartei` и таблиц `tblKartei/pre/decl` одинакова.
- ID уникальны внутри года, но **формально никак не связаны** между годами.

При переходе на единую веб‑БД целесообразно добавить явное поле `year` (или эквивалент) для всех записей `Kartei` / `pre` / `decl`, чтобы заменить "по файлу на год" на "по полю год в единой таблице".

---

## 7. Mapping to Django Models

Лист `Kartei` и таблица `tblKartei` проецируются на Django-модель `karteien.models.KarteiRecord`.

### 7.1 Соответствие полей

| Excel Col | Excel Idx | Access Field | Django Field       | Django Type                 |
| --------- | --------- | ------------ | ------------------ | --------------------------- |
| -         | -         | -            | `pkid`             | `BigAutoField` (PK)         |
| -         | -         | -            | `year`             | `PositiveSmallIntegerField` |
| AV        | 48        | `ID`         | `id`               | `PositiveIntegerField`      |
| A         | 1         | `Value1`     | `family_id`        | `CharField(50)`             |
| B         | 2         | `Value2`     | `parent_name`      | `CharField(255)`            |
| D         | 4         | `Value4`     | `child_name`       | `CharField(255)`            |
| E         | 5         | `Value5`     | `birthdate`        | `DateField`                 |
| F         | 6         | `Value6`     | `address`          | `CharField(500)`            |
| G         | 7         | `Value7`     | `phone`            | `CharField(50)`             |
| H         | 8         | `Value8`     | `mobile`           | `CharField(50)`             |
| I         | 9         | `Value9`     | `email`            | `EmailField(255)`           |
| J         | 10        | `Value10`    | `subject1`         | `CharField(255)`            |
| M         | 13        | `Value13`    | `price1`           | `DecimalField(10,2)`        |
| O         | 15        | `Value15`    | `subject2`         | `CharField(255)`            |
| R         | 18        | `Value18`    | `price2`           | `DecimalField(10,2)`        |
| U–AF      | 21–32     | `Value21–32` | `month_1..12`      | `DecimalField(10,2)`        |
| AK        | 37        | `Value37`    | `extra1`           | `CharField(255)`            |
| AL        | 38        | `Value38`    | `extra2`           | `CharField(255)`            |
| AM        | 39        | `Value39`    | `extra3`           | `CharField(255)`            |
| AU        | 47        | `Value47`    | `sepa_marker`      | `CharField(20)`             |
| BA        | 53        | —            | `status`           | `CharField` (choices enum)  |
| AW        | 49        | `Value49`    | `last_change_role` | `CharField(20)`             |
| AX        | 50        | `Value50`    | `last_change_date` | `DateField`                 |
| AY        | 51        | `Value51`    | `last_change_time` | `TimeField`                 |
| AZ        | 52        | `Value52`    | `history_raw`      | `TextField`                 |

### 7.2 Ключевые особенности

- **Primary Key**: Django PK — поле `pkid` (surrogate, BigAutoField).
- **Access/Excel ID**: Поле `id` хранит Access ID / Excel AV (48). Не глобально уникально: доменный ключ записи — `(year, id)`.
- **Год**: Поле `year` заменяет отдельные файлы баз по годам. Уникальность обеспечивается ограничением `(year, id)`.
- **Статус**: Поле `status` - enum (`RecordStatus`) со значениями `''` (normal), `'PENDING'`, `'DECLINED'`.
- **История**: Поле `history_raw` сохраняет legacy-формат истории. Нормализованная история - в отдельной модели `history.HistoryEvent`.
- **Tracked Fields**: Константа `TRACKED_FIELDS` в `models.py` определяет поля, изменения которых фиксируются в истории и считаются "risky".

### 7.3 Месячные поля

Константы для работы с месяцами определены в `karteien/models.py`:

```python
MONTH_FIELD_NAMES = ("month_1", "month_2", ..., "month_12")
MONTH_NUM_TO_FIELD = {1: "month_1", 2: "month_2", ..., 12: "month_12"}
EXCEL_COL_TO_MONTH = {21: 1, 22: 2, ..., 32: 12}
```

Модель предоставляет методы `get_month_value(n)`, `set_month_value(n, val)`, `get_all_months()` для удобной работы с месяцами.

### 7.4 Catalog Reference Fields (Web Extension)

Веб-система расширяет `KarteiRecord` nullable FK-полями для связи со справочниками:

| Django Field    | FK Target             | Описание                                         |
| --------------- | --------------------- | ------------------------------------------------ |
| `subject1_ref`  | `catalog.Subject`     | Ссылка на предмет 1-го полугодия (мес. 1-6)      |
| `teacher1_ref`  | `catalog.Teacher`     | Ссылка на преподавателя для предмета 1           |
| `price1_ref`    | `catalog.PriceOption` | Ссылка на цену для предмета 1                    |
| `subject2_ref`  | `catalog.Subject`     | Ссылка на предмет 2-го полугодия (мес. 7-12)     |
| `teacher2_ref`  | `catalog.Teacher`     | Ссылка на преподавателя для предмета 2           |
| `price2_ref`    | `catalog.PriceOption` | Ссылка на цену для предмета 2                    |
| `start_month_1` | —                     | Стартовый месяц начисления 1-го полугодия (1-6)  |
| `start_month_2` | —                     | Стартовый месяц начисления 2-го полугодия (7-12) |

**Семантика:**

- Поля `subject1/price1` относятся к 1-му полугодию (месяцы 1-6).
- Поля `subject2/price2` относятся к 2-му полугодию (месяцы 7-12).
- `start_month_*` определяет, с какого месяца начинается начисление (месяцы до старта = 0.00).

**Синхронизация с legacy-полями:**

- При сохранении формы, если выбран `subject*_ref`, legacy-поле `subject*` заполняется именем предмета.
- Если выбран `price*_ref`, legacy-поле `price*` заполняется суммой.
- Это обеспечивает обратную совместимость с импортом и экспортом в Access.

**Валидация:**

- `price*_ref.year` должен совпадать с `record.year`.
- `price*_ref.subject` должен совпадать с `subject*_ref` (если оба выбраны).
- `teacher*_ref` должен иметь активный `TeachingAssignment` для `(year, subject)`.

**Обратная совместимость:**

- Все ref-поля nullable; импортированные записи с заполненными legacy-полями работают без ref.
- При редактировании форма пытается prefill ref из legacy (по совпадению имени предмета/суммы).

**Дополнительные поля в Postgres (web extension):**

| Django Field                      | Django Type          | Описание |
| -------------------------------- | -------------------- | -------- |
| `months_mode`                     | `CharField`          | Режим начислений: `LEGACY` / `AUTO` / `OVERRIDE`. |
| `base_amounts`                    | `JSONField`          | Базовые суммы по месяцам до скидок (в `AUTO`). |
| `hours_amounts`                   | `JSONField`          | UE по месяцам для Stundenfächer (в `AUTO`). |
| `discounts_disabled`              | `BooleanField`       | Отключить скидки для записи. |
| `discounts_disabled_months`       | `JSONField`          | Месяцы (1-12), где скидки отключены точечно. |
| `contract_terminated_from_month`  | `PositiveSmallIntegerField` | Месяц (1-12), с которого начисления должны быть 0 при расторгнутом контракте. |

Дополнительные legacy-поля, импортируемые из Access (учителя/маркеры контракта), описаны в разделе §8.2.

### 7.5 Pending / Declined

Таблицы `pre_tblKartei` и `decl_tblKartei` проецируются на модели в приложении `approvals`:

- `approvals.models.PendingChange`
- `approvals.models.DeclinedChange`

#### Mapping to Django approvals models

**Access `pre_tblKartei` → `approvals.PendingChange`**

| Access Field      | Django Field      | Notes                                              |
| ----------------- | ----------------- | -------------------------------------------------- |
| `ID`              | `record` (FK)     | Ссылка на `KarteiRecord`, not duplicated as field. |
| `Value1..Value52` | `snapshot` (JSON) | Tracked fields хранятся в JSON-снимке.             |
| `InteriorColor*`  | —                 | Цвета не переносятся в pending (только данные).    |

Ключевые отличия от Access:

- В Access `pre_tblKartei` повторяет структуру `tblKartei` (52 поля).
- В Django `PendingChange` хранит только tracked-поля в JSON-поле `snapshot`.
- Связь с оригинальной записью через `OneToOneField` (одна pending-запись на ID).
- Поле `is_processed` отслеживает, обработано ли изменение.

**Access `decl_tblKartei` → `approvals.DeclinedChange`**

| Access Field      | Django Field      | Notes                                      |
| ----------------- | ----------------- | ------------------------------------------ |
| `ID`              | `record` (FK)     | Ссылка на `KarteiRecord`.                  |
| `Value1..Value52` | `snapshot` (JSON) | Снимок отклонённых значений.               |
| —                 | `decline_reason`  | Причина отклонения (нет аналога в Access). |
| —                 | `declined_by`     | Superadmin, отклонивший изменение.         |

Ключевые отличия от Access:

- В Access `decl_tblKartei` хранит только данные, причина записывается в историю (DCL).
- В Django `DeclinedChange` явно хранит `decline_reason` и ссылку на `declined_by`.
- Используется `ForeignKey` (не OneToOne), т.к. возможны множественные отклонения одной записи.

---

## 8. Legacy Import Mapping

Раздел описывает маппинг данных при импорте из Access в Django через
`legacy_import.import_access_year`.

### 8.1 Ключ записи: (year, ID)

- **pkid** — surrogate PK Django (BigAutoField), используется во всех URL/FK внутри веб‑системы.
- **id** — Access/Excel ID из Access поля `ID` (Excel: AV/48), **не** глобально уникален.
- **year** — год, передаваемый параметром `--year` команды импорта.
- Уникальное ограничение (доменный ключ) в Django: `(year, id)` в `KarteiRecord`. Импорт и patch‑импорт ищут записи по нему.

**Границы анализа:** Команда `legacy_import.import_access_year` анализирует данные
и FamilyID на уровне одного года (`--year`). Для расширенного межгодового анализа
FamilyID предполагаются отчёты поверх PostgreSQL без изменения текущего импорта.

### 8.2 Access Value Fields → Django Fields

| Access Field | Django Field       | Тип                  | Примечание                       |
| ------------ | ------------------ | -------------------- | -------------------------------- |
| `ID`         | `id`               | PositiveIntegerField | Access/Excel ID (не Django PK; доменный ключ с `year`) |
| —            | `pkid`             | BigAutoField         | Django PK (surrogate), отсутствует в Access |
| —            | `year`             | PositiveSmallInt     | Из параметра `--year`            |
| `Value1`     | `family_id`        | CharField(50)        | A - FamilyID                     |
| `Value2`     | `parent_name`      | CharField(255)       | B - Parent                       |
| `Value3`     | —                  | —                    | Пропускается (visual)            |
| `Value4`     | `child_name`       | CharField(255)       | D - Child                        |
| `Value5`     | `birthdate`        | DateField            | E - Birthdate                    |
| `Value6`     | `address`          | CharField(500)       | F - Address                      |
| `Value7`     | `phone`            | CharField(50)        | G - Phone                        |
| `Value8`     | `mobile`           | CharField(50)        | H - Mobile                       |
| `Value9`     | `email`            | EmailField           | I - Email                        |
| `Value10`    | `subject1`         | CharField(255)       | J - Subject1                     |
| `Value11`    | `teacher1_legacy_name` | CharField(255)   | K - Lehrer (1. HJ, legacy text)  |
| `Value12`    | —                  | —                    | Пропускается                     |
| `Value13`    | `price1`           | DecimalField         | M - Price1                       |
| `Value14`    | `contract_type_raw`| CharField(255)       | N - Vertragstyp (Rohtext; содержит `O/V` и др.) |
| —            | `is_monthly_contract` | BooleanField      | Derived: содержит `O/V` (case-insensitive) |
| `Value15`    | `subject2`         | CharField(255)       | O - Subject2                     |
| `Value16`    | `teacher2_legacy_name` | CharField(255)   | P - Lehrer (2. HJ, legacy text)  |
| `Value17`    | —                  | —                    | Пропускается                     |
| `Value18`    | `price2`           | DecimalField         | R - Price2                       |
| `Value19`    | —                  | —                    | Пропускается                     |
| `Value20`    | `contract_status_raw` | CharField(255)    | T - Vertragsstatus (Rohtext; содержит `KN` и др.) |
| —            | `is_contract_terminated` | BooleanField    | Derived: токен `KN` (case-insensitive, separated by whitespace) |
| `Value21`    | `month_1`          | DecimalField         | U - Month 1                      |
| `Value22`    | `month_2`          | DecimalField         | V - Month 2                      |
| ...          | ...                | ...                  | ...                              |
| `Value32`    | `month_12`         | DecimalField         | AF - Month 12                    |
| `Value33-36` | —                  | —                    | Пропускаются                     |
| `Value37`    | `extra1`           | CharField(255)       | AK - Extra1                      |
| `Value38`    | `extra2`           | CharField(255)       | AL - Extra2                      |
| `Value39`    | `extra3`           | CharField(255)       | AM - Extra3                      |
| `Value40-46` | —                  | —                    | Пропускаются                     |
| `Value47`    | `sepa_marker`      | CharField(20)        | AU - SepaMarker                  |
| `Value48`    | —                  | —                    | Игнорируется (используется `ID`) |
| `Value49`    | `last_change_role` | CharField(20)        | AW - LastChangeRole              |
| `Value50`    | `last_change_date` | DateField            | AX - LastChangeDate              |
| `Value51`    | `last_change_time` | TimeField            | AY - LastChangeTime              |
| `Value52`    | `history_raw`      | TextField            | AZ - History (memo)              |

Patch‑режим импорта: `python manage.py import_access_year --patch-fields` дополняет уже импортированные записи (по `(year, id)`) учителями/контрактными маркерами/SEPA, не трогая остальные поля.

### 8.3 Маркерные строки "Zahlung"

Маркерные строки — визуальные разделители семей в Excel, которые **НЕ импортируются**:

- `Value4` (Child) = `" Zahlung"` (с пробелом) или `"Zahlung"`
- Остальные поля обычно пусты
- Строка залита зелёным (`InteriorColor1` ∈ {5287936, 5296274, 32768, 65280})

При импорте такие строки:

- Не создают `KarteiRecord`
- Учитываются в статистике `marker_skipped`
- Записываются в отчёт

### 8.4 FamilyID и Family Key

**Проблема legacy-данных:**

- FamilyID не контролировался строго
- Одна семья могла иметь разные FamilyID в разные годы
- Один FamilyID мог использоваться разными семьями

**Family Key:**
Для анализа используется `family_key = normalize(parent_name) | normalize(email)`:

- `normalize()` — trim, lowercase, удаление управляющих символов

**Обнаруживаемые проблемы:**

1. `SAME_FAMILY_DIFFERENT_FAMILYIDS` — один family_key, несколько FamilyID
2. `SAME_FAMILYID_DIFFERENT_FAMILIES` — один FamilyID, несколько family_key

**Политики:**

- `report` (по умолчанию) — только отчёт
- `auto-merge` — для случая #1 выбирается канонический FamilyID (первый по алфавиту)

**Web‑расширение (отчёты по FamilyID):**

- Для SUPERADMIN реализован сводный multi-year отчёт начислений по семье (группировка по `family_id` по всем годам в БД).
- В отчёт попадают только значения из `RecordStatus.NORMAL`; если в каком-то году у семьи есть `PENDING`, в отчёте используются NORMAL-значения, но отображается пометка о наличии альтернативы (ожидающей одобрения).

### 8.5 Импорт pre_tblKartei и decl_tblKartei

**pre_tblKartei → PendingChange:**

- Находит/создаёт `KarteiRecord` по `(year, ID)`
- Устанавливает `status = PENDING`
- Создаёт `PendingChange` с `snapshot` из tracked-полей
- Маркерные строки пропускаются

**decl_tblKartei → DeclinedChange:**

- Находит/создаёт `KarteiRecord` по `(year, ID)`
- Устанавливает `status = DECLINED`
- Создаёт `DeclinedChange` с `snapshot` и `decline_reason = "Imported from legacy..."`
- Маркерные строки пропускаются

### 8.6 Синхронизация истории

При `--sync-history`:

- Для каждой записи с непустым `history_raw` вызывается `history.sync_history_from_raw()`
- Парсит `history_raw` и создаёт `HistoryEvent` для каждой сессии
- Существующие события не перезаписываются (идемпотентно)
- Если запись уже была в базе, её история **не переписывается** — новые события
  добавляются только для фрагментов, которых ещё нет

---

## 9. Reference / Catalog Tables (Справочники)

Справочные таблицы для стандартизации данных о преподавателях и предметах.
Эти таблицы **не заменяют** существующие поля `KarteiRecord.subject1/subject2`
и не влияют на импорт из Access. Это отдельные справочники для будущих форм
и отчётности.

### 9.1 Teacher (Преподаватель)

| Поле       | Тип          | Описание                                |
| ---------- | ------------ | --------------------------------------- |
| last_name  | CharField    | Фамилия (обязательное, max 100)         |
| first_name | CharField    | Имя (обязательное, max 100)             |
| is_active  | BooleanField | Активен ли преподаватель (default=True) |

**Ограничения:**

- UniqueConstraint на `(last_name, first_name)` — уникальность по полному имени.

**Отображение:** `__str__` → `"LastName, FirstName"`

### 9.2 Subject (Предмет)

| Поле      | Тип          | Описание                            |
| --------- | ------------ | ----------------------------------- |
| name      | CharField    | Название предмета (unique, max 200) |
| is_active | BooleanField | Активен ли предмет (default=True)   |

**Отображение:** `__str__` → `name`

### 9.3 TeachingAssignment (Назначение преподавателя)

Связь "в каком году какой преподаватель вёл какой предмет".

| Поле      | Тип                       | Описание                             |
| --------- | ------------------------- | ------------------------------------ |
| year      | PositiveSmallIntegerField | Учебный год (2024, 2025, ...)        |
| subject   | FK → Subject              | Предмет (on_delete=PROTECT)          |
| teacher   | FK → Teacher              | Преподаватель (on_delete=PROTECT)    |
| is_active | BooleanField              | Активно ли назначение (default=True) |

**Ограничения:**

- UniqueConstraint на `(year, subject, teacher)` — уникальность назначения.

**Индексы:**

- `(year, subject)` — для быстрого поиска преподавателей по году и предмету.
- `(year, teacher)` — для быстрого поиска предметов по году и преподавателю.

**Отображение:** `__str__` → `"YEAR: SubjectName — LastName, FirstName"`

### 9.4 PriceOption (Прайс-лист)

Цены по году и дисциплине. Позволяет управлять прайс-листом с разными вариантами цен для одного предмета.

| Поле      | Тип                       | Описание                                            |
| --------- | ------------------------- | --------------------------------------------------- |
| year      | PositiveSmallIntegerField | Учебный год (2024, 2025, ...)                       |
| subject   | FK → Subject              | Предмет (on_delete=PROTECT)                         |
| amount    | DecimalField(10,2)        | Сумма в € (не может быть отрицательной)             |
| comment   | TextField                 | Комментарий — почему такая цена (blank, default="") |
| is_active | BooleanField              | Активна ли цена (default=True)                      |

**Ограничения:**

- UniqueConstraint на `(year, subject, amount, comment)` — уникальность цены.
- CheckConstraint: `amount >= 0` — запрет отрицательных значений.

**Индексы:**

- `(year, subject)` — для быстрого поиска цен по году и предмету.

**Семантика цены (`amount`):**

- По умолчанию: **цена за месяц** (€/Monat) — для групповых занятий.
- Для предметов с "Ind." или "VSpE\_" в названии (индивидуальные занятия): **цена за академический час** (€/UE).
- Для предметов с "NH" или "Nachhilfe" в названии (наххильфе): **цена за академический час** (€/UE).

**Методы:**

- `get_price_unit()` → `"€/Monat"` или `"€/UE"` (определяется по названию предмета).
- `is_per_hour()` → `True` если цена за академический час.

**Отображение:** `__str__` → `"YEAR: SubjectName – AMOUNT UNIT (comment)"`

### 9.6 FamilyIdReservation (резервация FamilyID, Web Extension)

Отдельная таблица для сценария «клиент звонит сейчас, а данные внесём позже»: админ резервирует следующий свободный FamilyID и может позже использовать его при создании семьи.

| Поле        | Тип           | Описание |
| ----------- | ------------- | -------- |
| family_id   | CharField(50) | Зарезервированный FamilyID (`"1. <number>"`), уникальный |
| reserved_at | DateTimeField | Дата/время резервации (auto_now_add) |
| reserved_by | FK → User      | Кто зарезервировал (nullable) |
| is_used     | BooleanField   | Использована ли резервация |
| used_at     | DateTimeField  | Когда использована (nullable) |
| note        | CharField(200) | Опциональная заметка |

**Семантика:**

- «Следующий FamilyID» рассчитывается глобально по всем годам как `max(использованные, активные резервации) + 1` (для формата `"1. <number>"`).
- Резервация исключает номер из автоподбора до момента использования/отмены.

### 9.5 Примеры запросов

```python
from apps.catalog.models import Teacher, Subject, TeachingAssignment, PriceOption

# Все преподаватели, которые вели "Gitarre" в 2024 году
teachers = Teacher.objects.filter(
    assignments__year=2024,
    assignments__subject__name="Gitarre",
    assignments__is_active=True
)

# Все предметы преподавателя за все годы
subjects = Subject.objects.filter(
    assignments__teacher__last_name="Müller",
    assignments__teacher__first_name="Hans"
).distinct()

# Список всех назначений за 2025 год
assignments_2025 = TeachingAssignment.objects.filter(
    year=2025,
    is_active=True
).select_related('subject', 'teacher')

# Все годы, в которые велся предмет
years = TeachingAssignment.objects.filter(
    subject__name="Klavier"
).values_list('year', flat=True).distinct()

# Все активные цены для предмета в 2025 году
prices_klavier_2025 = PriceOption.objects.filter(
    year=2025,
    subject__name="Klavier",
    is_active=True
).order_by('amount')

# Получить цену с единицей измерения
for price in prices_klavier_2025:
    unit = price.get_price_unit()  # "€/Monat" или "€/UE"
    print(f"{price.amount} {unit}: {price.comment}")

# Все предметы с ценами за академчас (индивидуальные, наххильфе)
per_hour_subjects = Subject.objects.filter(
    price_options__year=2025,
    price_options__is_active=True
).filter(
    models.Q(name__icontains='Ind.') |
    models.Q(name__icontains='VSpE_') |
    models.Q(name__icontains='NH') |
    models.Q(name__icontains='Nachhilfe')
).distinct()
```

---

## 10. NeuList (Новые записи)

Функциональность NeuList позволяет Superadmin отслеживать новые записи, появившиеся с момента последнего просмотра.

### 10.1 VBA реализация (valid_NeuList.bas)

В VBA NeuList реализован через:

- **NeuConfig sheet**: Скрытый лист с `LastSeenID` (последний просмотренный ID).
- **Neu sheet**: Отображает все записи из `tblKartei` с `ID > LastSeenID`.
- **RefreshNeuList**: Обновляет список, загружая новые записи из Access.
- **MarkAllSeen**: Устанавливает `LastSeenID = max(ID)`.

### 10.2 Django реализация

В веб-интерфейсе NeuList реализован через:

**Модель `SuperadminState`** (`apps/approvals/models.py`):

| Поле             | Тип                  | Описание                                           |
| ---------------- | -------------------- | -------------------------------------------------- |
| `user`           | OneToOneField        | Ссылка на User (Superadmin).                       |
| `last_seen_id`   | PositiveIntegerField | Последний просмотренный ID.                        |
| `last_seen_date` | DateTimeField        | Альтернативный трекер (дата последнего просмотра). |
| `updated_at`     | DateTimeField        | Время последнего обновления состояния.             |

**Сервисы** (`apps/approvals/services.py`):

- `get_or_create_superadmin_state(user)` — получает или создаёт состояние.
- `get_new_records(user, year)` — возвращает записи с `id > last_seen_id`.
- `get_new_records_count(user, year)` — количество новых записей.
- `update_last_seen_id(user, max_id=None)` — обновляет `last_seen_id`.

**URL-эндпоинты:**

| URL                                | View                     | Описание                        |
| ---------------------------------- | ------------------------ | ------------------------------- |
| `/approvals/superadmin/neu/`       | `SuperadminNeuListView`  | Список новых записей.           |
| `/approvals/superadmin/mark-seen/` | `SuperadminMarkSeenView` | Отметить все как просмотренные. |

### 10.3 Критерий "новой" записи

Запись считается "новой" если:

- `record.id > state.last_seen_id`
- Опционально: `record.created_at > state.last_seen_date` (для случаев, когда ID не монотонно возрастает)

По умолчанию используется ID-based подход (аналог VBA).

### 10.4 Отображение в UI

NeuList страница показывает:

- Таблицу новых записей (ID, FamilyID, Parent, Child, Birthdate, Status)
- Текущее значение `last_seen_id`
- Количество новых записей
- Кнопку "Alle als gesehen markieren" для сброса
- Фильтр по году

---

## 11. Справочник скидок (Discounts)

Система скидок позволяет назначать процентные и фиксированные скидки на семьи
или отдельные записи с привязкой к конкретным месяцам.

### 11.1 Модели скидок

**Discount (Справочник скидок):**

| Поле          | Тип          | Описание                                              |
| ------------- | ------------ | ----------------------------------------------------- |
| `id`          | AutoField    | Primary key.                                          |
| `kind`        | CharField    | Тип: `PERCENT` или `FIXED`.                           |
| `value`       | DecimalField | Значение: для PERCENT 0.00-0.99, для FIXED сумма в €. |
| `description` | TextField    | Описание скидки (например "Geschwisterrabatt").       |
| `is_active`   | BooleanField | Активна ли скидка.                                    |

**FamilyDiscount (Скидка на семью):**

| Поле          | Тип                       | Описание                                          |
| ------------- | ------------------------- | ------------------------------------------------- |
| `id`          | AutoField                 | Primary key.                                      |
| `year`        | PositiveSmallIntegerField | Учебный год.                                      |
| `family_id`   | CharField                 | Идентификатор семьи (как в KarteiRecord).         |
| `discount`    | FK → Discount             | Применяемая скидка.                               |
| `start_month` | PositiveSmallIntegerField | Начальный месяц (1-12), default=1.                |
| `end_month`   | PositiveSmallIntegerField | Конечный месяц (1-12), default=12.                |
| `months`      | JSONField                 | Опционально: список месяцев [1,3,5] вместо range. |
| `created_at`  | DateTimeField             | Время создания.                                   |
| `updated_at`  | DateTimeField             | Время обновления.                                 |

**RecordDiscount (Скидка на запись):**

| Поле          | Тип                       | Описание                                          |
| ------------- | ------------------------- | ------------------------------------------------- |
| `id`          | AutoField                 | Primary key.                                      |
| `record`      | FK → KarteiRecord         | Запись, к которой применяется скидка.             |
| `discount`    | FK → Discount             | Применяемая скидка.                               |
| `start_month` | PositiveSmallIntegerField | Начальный месяц (1-12), default=1.                |
| `end_month`   | PositiveSmallIntegerField | Конечный месяц (1-12), default=12.                |
| `months`      | JSONField                 | Опционально: список месяцев [1,3,5] вместо range. |
| `created_at`  | DateTimeField             | Время создания.                                   |
| `updated_at`  | DateTimeField             | Время обновления.                                 |

### 11.2 Логика применимости месяцев

Метод `get_applicable_months()` возвращает список месяцев, к которым применяется скидка:

1. Если `months` заполнено (например `[1, 3, 5]`) — используется этот список.
2. Иначе используется диапазон `start_month..end_month` (включительно).

Примеры:

- `start_month=1, end_month=12, months=None` → весь год `[1,2,3,4,5,6,7,8,9,10,11,12]`
- `start_month=3, end_month=6, months=None` → `[3,4,5,6]`
- `months=[1,6,12]` → только январь, июнь, декабрь (range игнорируется)

### 11.3 Типы скидок

**Процентная скидка (PERCENT):**

- Значение хранится как дробь: 0.25 = 25%, 0.10 = 10%.
- Допустимый диапазон: 0.00 - 0.99.
- Применяется первой (до фиксированной).

**Фиксированная скидка (FIXED):**

- Значение в EUR: 10.00 = вычесть 10€.
- Допустимый диапазон: >= 0.
- Применяется после процентной скидки.

### 11.4 Иерархия применения

Скидки применяются в следующем порядке:

1. Семейные скидки (FamilyDiscount) — ко всем записям семьи.
2. Записные скидки (RecordDiscount) — к конкретной записи.
3. Внутри каждой категории: сначала PERCENT, затем FIXED.

**Пример расчёта для месяца:**

- Базовая цена: 100€
- Семейная скидка: 10% (PERCENT, value=0.10)
- Записная скидка: 5€ (FIXED, value=5.00)
- Расчёт: 100€ × (1 - 0.10) - 5€ = 90€ - 5€ = 85€

### 11.5 URL-эндпоинты

| URL                          | Описание                  |
| ---------------------------- | ------------------------- |
| `/catalog/discounts/`        | Справочник скидок (CRUD). |
| `/catalog/family-discounts/` | Скидки на семьи (CRUD).   |
| `/catalog/record-discounts/` | Скидки на записи (CRUD).  |
