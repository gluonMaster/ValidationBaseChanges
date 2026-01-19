# Инструкция для разработчика: внедрение мультигод‑системы в `KindElternDaten_Suprime.xlsm`

Документ описывает “под ключ”, как собрать новый централизованный файл суперадмина на базе существующего `KindElternDaten_26_Suprime.xlsm`, чтобы он работал с базами 2024/2025/2026 из одного файла.

## 0) Что получается в итоге

Один файл `KindElternDaten_Suprime.xlsm`:

- подключается к трём Access‑базам (`..._24_front.accdb`, `..._25_front.accdb`, `..._26_front.accdb`);
- импортирует pending‑изменения (`pre_tblKartei`) в годовые листы `KarteiYY`;
- строит годовые листы решений `grossGeschichteYY` и умеет синхронизировать решения обратно в соответствующую базу;
- строит отчёты истории по записи (`GeschichteYY`) и за период (`GeschichteYY_Alle`);
- умеет выгружать оплаты (`FamZahlungen`) по выбранному году;
- поддерживает список новых записей `Neu` только для 2026.

## 1) База для сборки

1. Возьмите `KindElternDaten_26_Suprime.xlsm` как исходник.
2. Сохраните копию как `KindElternDaten_Suprime.xlsm` (без года в имени).
3. Включите макросы (Enable Content) и откройте VBA‑редактор (`Alt+F11`).

## 2) VBA References (обязательно)

Код использует DAO типы (например, `DAO.Database`, `DAO.Recordset`). Поэтому в `Tools → References` включите один из вариантов:

- `Microsoft Office XX.0 Access database engine Object Library` (рекомендуется), или
- `Microsoft DAO 3.6 Object Library` (если ACE не доступен).

Примечания:

- `Scripting.Dictionary` и `VBScript.RegExp` используются через `CreateObject`, отдельные references для них не нужны.
- Для выбора папки используется `Application.FileDialog`, это часть Excel/Office.

После включения reference выполните `Debug → Compile VBAProject`.

## 3) Какие листы должны быть в книге

Сейчас в исходном файле (по скрину) есть: `Kartei`, `grossGeschichte`, `Neu`, `geschichteForm`, `Geschichte`.

### 3.1 Обязательные листы (новая система)

**UI**

- `Dashboard` (основная точка входа, выбор года + действия).

**Годовые листы (по каждому году 24/25/26)**

- Pending: `Kartei24`, `Kartei25`, `Kartei26`
- Решения/валидация: `grossGeschichte24`, `grossGeschichte25`, `grossGeschichte26`
- История по записи: `Geschichte24`, `Geschichte25`, `Geschichte26`
- История за период: `Geschichte24_Alle`, `Geschichte25_Alle`, `Geschichte26_Alle`

**Служебные (скрытые)**

- `DBConfig` (VeryHidden) — хранит пути к базам (cloud root + overrides).
- `NeuConfig` (VeryHidden) — хранит `LastSeenID` для листа `Neu`.

### 3.1.1 Что нужно создавать/заполнять вручную, а что создаётся макросами

Общее правило: **кнопки не обязательны**, а **шапки/данные на годовых листах заполняются макросами**. Но есть важный нюанс про “шаблонные” листы.

**Нюанс (важно):** автосоздание годовых листов (`Kartei24/25/26`, `grossGeschichte24/25/26`, `Geschichte24/25/26`, `GeschichteYY_Alle`) работает так, что копирует строки `1:2` из “шаблонных” legacy‑листов:

- `Kartei` → для `KarteiYY`
- `grossGeschichte` → для `grossGeschichteYY`
- `Geschichte` → для `GeschichteYY` и `GeschichteYY_Alle`

Поэтому **шаблонные листы должны существовать** и в них должны быть корректные заголовки в строках 1–2 (в базе `KindElternDaten_26_Suprime.xlsm` они уже есть).

Далее по каждому листу:

- `Dashboard`: создаётся и заполняется макросом `ShowDashboard` (вызывается при открытии книги). Кнопки на нём **не создаются автоматически** — там выводится “таблица подсказок” с названиями макросов. Пользователь может запускать макросы через `Developer → Macros`, а разработчик может (необязательно) добавить кнопки/фигуры и привязать их к макросам.
- `Dashboard`: если лист уже существует, `ShowDashboard` **дополни́т/инициализирует** разметку (не удаляя фигуры/кнопки). Поэтому в норме достаточно просто открыть книгу или вручную запустить `ShowDashboard` один раз.
- `Kartei24/25/26`: создаются макросами (через `valid_YearConfig.EnsureYearSheetsExist`). Заголовки копируются из `Kartei` (строки 1–2). Данные заполняются `Dashboard_LoadPending` (или `LoadPendingChangesForYear`/`LoadPendingAndBuildDecisionYY`) и пишутся с 3‑й строки. При повторном импорте строки 3+ очищаются, заголовки сохраняются.
- `grossGeschichte24/25/26`: создаются макросами, заголовки копируются из `grossGeschichte` (строки 1-2). Заполнение происходит при построении листа решений (`Dashboard_LoadPending` или `Dashboard_BuildDecisionSheet`). Период для агрегации комментариев задаётся на `Dashboard` (ячейки `E4:J4`) и перед запуском копируется в `grossGeschichteYY!B1/C1` (значения не должны затираться на "сегодня").
- `Geschichte24/25/26`: создаются макросами, заголовки копируются из `Geschichte`. Заполняются макросом истории по записи (`Dashboard_SingleRecordHistory` / `GeschichteMachenYY`). Период фильтра по датам задаётся на `Dashboard` (`E4:J4`) и копируется в `GeschichteYY!B1/C1` перед запуском.
- `Geschichte24_Alle/25_Alle/26_Alle`: создаются макросами, заголовки копируются из `Geschichte`. Период задаётся на `Dashboard` (`E4:J4`) и копируется в `GeschichteYY_Alle!B1/C1`, далее запускается `Dashboard_DateRangeHistory` / `GrossGeschichteMachenYY`.
- `Neu`: можно оставить существующий. Если листа нет — он создаётся макросом `RefreshNeuList` при первом запуске. Заголовок строки 2 при обновлении копируется с `Kartei26` (если есть) иначе с `Kartei`; данные пишутся с 3‑й строки. Никаких ручных обязательных значений на листе не требуется.
- `DBConfig` (VeryHidden): создаётся макросами при настройке БД (`Dashboard_ConfigureDatabase` / `ConfigureAllDatabases`). Вручную значения не задавать. Используемые ячейки: `A1` (CloudRoot), `B1/C1/D1` (override 24/25/26).
- `NeuConfig` (VeryHidden): создаётся макросами при первом запуске `RefreshNeuList`. Вручную значения не задавать. Используемые ячейки: `A1` (LastSeenID), `B1` (подпись).

### 3.2 Существующие листы (оставить)

Оставьте без удаления:

- `Neu` (используется как список новых записей; логика — только для 2026),
- `geschichteForm` (используется макросами как форма/вспомогательный лист),
- `Kartei`, `grossGeschichte`, `Geschichte` (как legacy/backward compatibility).

Важно:

- В новом пайплайне суперадмин работает с `KarteiYY` и `grossGeschichteYY`.
- Legacy‑листы могут оставаться для совместимости со старой логикой/старыми кнопками, но их рекомендуется “не использовать” в обычной работе.

### 3.3 Как создать листы

Есть два подхода:

**A) Ручная подготовка (желательно для “чистого” файла)**

- Создайте перечисленные листы вручную (пустые).
- `DBConfig` и `NeuConfig` сделайте `xlSheetVeryHidden`.

**B) Автосоздание макросами (допустимо)**

- Достаточно иметь только исходные листы + импортированные модули.
- При первом запуске макросы сами создадут недостающие листы.
- Для принудительного создания запустите `valid_YearConfig.EnsureYearSheetsExist 24/25/26`.

Примечание: чтобы “раздать” суперадминам файл уже с готовыми вкладками, удобно один раз (у разработчика) запустить `valid_YearConfig.EnsureYearSheetsExist` для 24/25/26 и сохранить книгу. Настройку `DBConfig` всё равно делает каждый суперадмин на своём ПК.

## 4) Какие модули должны быть в VBA‑проекте

### 4.1 Импорт из репозитория

Источник: папка `Superadmin/` в репозитории.

Импортируйте/обновите все модули `.bas` из `Superadmin/`.

Ключевые модули (должны точно присутствовать):

- `valid_YearConfig.bas` — мультигод‑конфиг, пути к базам, имена листов, автосоздание листов.
- `valid_Dashboard.bas` — макросы `Dashboard_*` и создание листа `Dashboard`.
- `valid_ImportPending.bas` + `valid_FormatMonths.bas` — импорт pending из `pre_tblKartei` в `KarteiYY`.
- `valid_GrossGeschichteDecision.bas` — построение `grossGeschichteYY` (War/Ist + решения).
- `valid_ApproveFlow.bas` — пайплайн “загрузить → построить → синхронизировать решения” по году.
- `valid_HistoryPerYear.bas` — история по записи/за период по году.
- `valid_ParseHistory.bas` — парсер истории (используется в history‑отчётах).
- `pay_Main.bas`, `pay_DataProcessor.bas`, `pay_FileGenerator.bas`, `pay_Utils.bas` — выгрузки оплат.
- `valid_NeuList.bas` — список новых записей (`Neu`) для 2026.
- `valid_DatabasePath.bas` — совместимость со старым API (делегирует в `valid_YearConfig`).

Примечание по обновлению:

- В VBE проще “удалить старый модуль → Import File… новый”, чем пытаться вручную править.
- Перед заменой можно экспортировать текущие модули как backup.

### 4.2 События книги (ThisWorkbook)

Файл `Superadmin/DieseArbeitsmappe.cls` содержит код для `ThisWorkbook`:

- `Workbook_Open`: показывает/создаёт `Dashboard`, (опционально) обновляет `Neu` без навязчивых диалогов.
- `Workbook_BeforeClose`: очищает годовые рабочие листы `KarteiYY` и `grossGeschichteYY` (строки 3+), сохраняет книгу.

Как внедрить:

1. В VBE откройте объект `ThisWorkbook`.
2. Полностью замените код содержимым из `Superadmin/DieseArbeitsmappe.cls`.

Настраиваемые флаги в `DieseArbeitsmappe`:

- `AUTO_REFRESH_NEU` — автообновление `Neu` при открытии.
- `CLEAR_ON_CLOSE` — очистка `KarteiYY/grossGeschichteYY` при закрытии.

## 5) Какие макросы считаются “входными точками” (UX)

### 5.1 Для пользователя через Dashboard

Год выбирается через:

- `SelectYear24`, `SelectYear25`, `SelectYear26` (сохраняют год в `Dashboard!B2`)

Текущая сборка (по скрину) использует кнопки на `Dashboard` со следующими привязками:

- `2024` → `SelectYear24`
- `2025` → `SelectYear25`
- `2026` → `SelectYear26`
- `Wartend` → `Dashboard_LoadPending`
- `Ganz Datenbank` → `Dashboard_ImportFromBase`
- `Zahlungen` → `Dashboard_FamZahlungen`
- `Datums Geschichte` → `Dashboard_DateRangeHistory`
- `Database-Pfad` → `Dashboard_ConfigureDatabase`

Основные действия (используют выбранный год):

- `Dashboard_LoadPending` — импорт `pre_tblKartei` → `KarteiYY` + сборка `grossGeschichteYY`.
- `Dashboard_BuildDecisionSheet` — пересобрать `grossGeschichteYY` без повторного импорта pending.
- `Dashboard_SyncDecisions` — применить решения из `grossGeschichteYY` в БД (approve/decline).
- `Dashboard_SingleRecordHistory` — история по выбранной строке `KarteiYY` → `GeschichteYY`.
- `Dashboard_DateRangeHistory` — история за период → `GeschichteYY_Alle`.
- `Dashboard_FamZahlungen` — выгрузка оплат по `KarteiYY` в `C:\FamZahlung` (файлы с суффиксом года).
- `Dashboard_RefreshNeu` — обновить `Neu` (только 2026).
- `Dashboard_ConfigureDatabase` - первичная настройка путей к базам.
- `Dashboard_ShowConfig` - показать текущие пути/статусы (неинтерактивно, без диалогов).
- `Dashboard_ImportFromBase` - импорт полной базы `tblKartei` → `KarteiYY` (режим просмотра; `grossGeschichteYY` не строится).

Создание/показ Dashboard:

- `ShowDashboard` (создаёт лист, если его нет).

### 5.2 “Прямые” годовые макросы (для кнопок/меню)

Аналоги без Dashboard (удобно привязать к кнопкам на листах):

- `LoadPendingAndBuildDecision24/25/26`
- `BuildPendingDecisionSheet24/25/26`
- `SyncDecisions24/25/26`
- `GeschichteMachen24/25/26`
- `GrossGeschichteMachen24/25/26`
- `ImportFromBase24/25/26`
- `FamZahlungen24/25/26`

## 6) Как должна работать первичная настройка БД (поведение)

Пути на каждом ПК разные, поэтому конфиг хранится в книге:

- `DBConfig!A1` — CloudRoot (корневая папка, где есть `2024/2025/2026`).
- `DBConfig!B1/C1/D1` — overrides для 24/25/26 (если год лежит не в CloudRoot).

Ожидаемая структура по умолчанию:
`<CloudRoot>\<YYYY>\Alarm\KindElternDaten_<YY>_front.accdb`

Если DB не найдена:

- рабочие макросы могут запустить guided‑setup (выбор папки/базы),
- но `Dashboard_ShowConfig` и `Workbook_Open` не должны инициировать “внезапные” диалоги (это уже учтено).

## 7) Чек‑лист приёмки сборки (перед выдачей суперадмину)

1. Открыть книгу → должен появиться `Dashboard` (создастся сам, если его нет).
2. Запустить `Dashboard_ConfigureDatabase` → указать CloudRoot.
3. Запустить `Dashboard_ShowConfig` → по всем годам статус `ok`.
4. Для каждого года 24/25/26:
   - `SelectYearYY`
   - `Dashboard_LoadPending` → заполняется `KarteiYY`, создаётся/заполняется `grossGeschichteYY`
   - на `grossGeschichteYY` выставить несколько `Approved/Declined` (AC/AD)
   - `Dashboard_SyncDecisions` → изменения уходят в правильную БД года
   - `Dashboard_SingleRecordHistory` / `Dashboard_DateRangeHistory` → отчёты попадают в `GeschichteYY` / `GeschichteYY_Alle`
   - `Dashboard_FamZahlungen` → файлы создаются в `C:\FamZahlung` с суффиксом года
5. `Dashboard_RefreshNeu`:
   - работает (2026),
   - не требуется для 2024/2025.
6. Закрыть книгу:
   - `KarteiYY` и `grossGeschichteYY` очищены (строки 3+),
   - `DBConfig`, `NeuConfig`, `Neu`, `Geschichte*`, `Dashboard` не очищены.

## 8) Набор минимально обязательных вкладок (кратко)

Если нужно просто “списком для сборки”, то в `KindElternDaten_Suprime.xlsm` должны быть:

- `Dashboard`
- `Kartei24`, `Kartei25`, `Kartei26`
- `grossGeschichte24`, `grossGeschichte25`, `grossGeschichte26`
- `Geschichte24`, `Geschichte25`, `Geschichte26`
- `Geschichte24_Alle`, `Geschichte25_Alle`, `Geschichte26_Alle`
- `Neu`
- `DBConfig` (VeryHidden)
- `NeuConfig` (VeryHidden)
- (оставить legacy): `Kartei`, `grossGeschichte`, `Geschichte`, `geschichteForm`

## 9) Про кнопки (Form Controls / Shapes)

Кнопки не являются техническим требованием: весь функционал доступен через `Developer → Macros`.

Если хотите “готовый UX”:

- добавьте на `Dashboard` обычные фигуры/кнопки (Insert → Shapes или Form Controls),
- назначьте им макросы из раздела 5.1 (`SelectYear24/25/26`, `Dashboard_*`),
- (опционально) продублируйте кнопки на годовых листах `KarteiYY`/`grossGeschichteYY` и привяжите к прямым макросам (`LoadPendingAndBuildDecisionYY`, `SyncDecisionsYY`, и т.д.).

Важно: `ShowDashboard` сам рисует только “табличку‑подсказку” с именами макросов и инструкциями — реальные кнопки он не создаёт.
