# DOMAIN_MODEL – KindEltern Data Model

Примечание по исходникам VBA: все выгруженные модули Excel/VBA (для справки и сопоставления логики) находятся в `legacy_VBA/`:

- `legacy_VBA/*.bas` и `legacy_VBA/*.cls` — модули из `KindElternDaten_XX_Admin.xlsm`
- `legacy_VBA/admin_forms/` — VBA-код форм из `KindElternDaten_XX_Admin.xlsm`
- `legacy_VBA/Superadmin/` — модули из `KindElternDaten_XX_Suprime.xlsm`
- `legacy_VBA/alt/` — модули из `KindElternDaten_XX_Data.xlsm`

Этот документ описывает **фактическую** модель данных системы в текущей реализации на Excel+Access. Он служит источником истины при проектировании Django‑моделей и миграции.

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
| —         | —         | —            | `year`             | `PositiveSmallIntegerField` |
| AV        | 48        | `ID`         | `id`               | `PositiveIntegerField` (PK) |
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

- **Primary Key**: Поле `id` задаётся явно (`PositiveIntegerField`), а не автоинкремент. Значение берётся из Access ID / Excel AV.
- **Год**: Новое поле `year` заменяет отдельные файлы баз по годам. Уникальное ограничение `(year, id)`.
- **Статус**: Поле `status` — enum (`RecordStatus`) со значениями `''` (normal), `'PENDING'`, `'DECLINED'`.
- **История**: Поле `history_raw` сохраняет legacy-формат истории. Нормализованная история — в отдельной модели `history.HistoryEvent`.
- **Tracked Fields**: Константа `TRACKED_FIELDS` в `models.py` определяет поля, изменения которых фиксируются в истории и считаются "risky".

### 7.3 Месячные поля

Константы для работы с месяцами определены в `karteien/models.py`:

```python
MONTH_FIELD_NAMES = ("month_1", "month_2", ..., "month_12")
MONTH_NUM_TO_FIELD = {1: "month_1", 2: "month_2", ..., 12: "month_12"}
EXCEL_COL_TO_MONTH = {21: 1, 22: 2, ..., 32: 12}
```

Модель предоставляет методы `get_month_value(n)`, `set_month_value(n, val)`, `get_all_months()` для удобной работы с месяцами.

### 7.4 Pending / Declined

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

- **ID** — числовой идентификатор из Access поля `ID` (Excel: AV/48).
- **year** — год, передаваемый параметром `--year` команды.
- Уникальное ограничение в Django: `(year, id)` в `KarteiRecord`.
- ID уникален в пределах одного года; один ID может существовать в разных годах.

**Границы анализа:** Команда `legacy_import.import_access_year` анализирует данные
и FamilyID на уровне одного года (`--year`). Для расширенного межгодового анализа
FamilyID предполагаются отчёты поверх PostgreSQL без изменения текущего импорта.

### 8.2 Access Value Fields → Django Fields

| Access Field | Django Field       | Тип                  | Примечание                       |
| ------------ | ------------------ | -------------------- | -------------------------------- |
| `ID`         | `id`               | PositiveIntegerField | PK в Django, ключ записи         |
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
| `Value11-12` | —                  | —                    | Пропускаются                     |
| `Value13`    | `price1`           | DecimalField         | M - Price1                       |
| `Value14`    | —                  | —                    | Пропускается                     |
| `Value15`    | `subject2`         | CharField(255)       | O - Subject2                     |
| `Value16-17` | —                  | —                    | Пропускаются                     |
| `Value18`    | `price2`           | DecimalField         | R - Price2                       |
| `Value19-20` | —                  | —                    | Пропускаются                     |
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

## 9. NeuList (Новые записи)

Функциональность NeuList позволяет Superadmin отслеживать новые записи, появившиеся с момента последнего просмотра.

### 9.1 VBA реализация (valid_NeuList.bas)

В VBA NeuList реализован через:

- **NeuConfig sheet**: Скрытый лист с `LastSeenID` (последний просмотренный ID).
- **Neu sheet**: Отображает все записи из `tblKartei` с `ID > LastSeenID`.
- **RefreshNeuList**: Обновляет список, загружая новые записи из Access.
- **MarkAllSeen**: Устанавливает `LastSeenID = max(ID)`.

### 9.2 Django реализация

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

### 9.3 Критерий "новой" записи

Запись считается "новой" если:

- `record.id > state.last_seen_id`
- Опционально: `record.created_at > state.last_seen_date` (для случаев, когда ID не монотонно возрастает)

По умолчанию используется ID-based подход (аналог VBA).

### 9.4 Отображение в UI

NeuList страница показывает:

- Таблицу новых записей (ID, FamilyID, Parent, Child, Birthdate, Status)
- Текущее значение `last_seen_id`
- Количество новых записей
- Кнопку "Alle als gesehen markieren" для сброса
- Фильтр по году
