# Spec - KindEltern Web (контракты, инварианты, требования)

Этот документ фиксирует "опорные" контракты и правила, которые нельзя случайно сломать при дальнейшем развитии.

## 0) Актуальность: current implementation / known defects / target v2

Дата актуализации этого раздела: 2026-05-27.

### Current implementation

- SAFE/RISKY workflow существует, но SAFE direct-write в live `KarteiRecord` еще используется.
- Admin может редактировать `PENDING`/`DECLINED` через стандартный `/karteien/<pkid>/edit/` editor; Operator в этих статусах блокируется.
- `PendingChange.snapshot` совместим со старым flat tracked-fields форматом и дополнительно может содержать reserved metadata keys.
- PRICELIST V2 уже имеет category/group models и UI, но billing pipeline еще не консолидирован.

### Known defects

- Preview / form save / category apply / bulk apply / approve / decline могут идти разными расчетными путями.
- Часть billing/context fields prewrite'ится в live record до approval.
- `base_amounts` и `month_*` могут временно относиться к разным расчетам.
- Legacy contract fields и timeline entries еще не полностью сведены к единому source of truth.

### Target stabilization v2

Целевой контракт описан в `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`:

- live record остается approved-state, пока change находится в `PENDING`;
- billing/context changes хранятся во frozen snapshot v2 payload;
- `APPROVED` применяет reviewed payload, а не пересчитывает proposal заново;
- `base_amounts` в `AUTO` остается historical base history.

## 1) Ключевые сущности и идентификаторы

### 1.1 KarteiRecord

- Модель: `karteien.KarteiRecord` (таблица `karteien_record`).
- `pkid`: Django PK (surrogate).
- `id`: Access/Excel ID (исторический идентификатор).
- Доменный ключ записи: `(year, id)` (уникален).

Следствие: любые ссылки/URL/поиски по "ID записи" должны явно понимать, что используется: `pkid` или `(year, id)`.

### 1.2 Статус записи

- `RecordStatus.NORMAL` хранится как пустая строка `""`.
- Проверка "normal" должна делаться как `not record.status`.

## 2) Роли и доступы (server-side)

- `ADMIN`: list/detail/live-search/create/edit/delete.
- `OPERATOR`: list/detail/live-search/create/edit, но без delete (и другие ограничения по проекту).
- `SUPERADMIN`: approvals (pending/declined) + read-only `/karteien/` (list/detail/live search), без create/edit/delete.
- `USER`: read-only `/user/...` и read-only `/karteien/` (list/detail/live search), без create/edit/delete.

UI-кнопки должны соответствовать реальным server-side ограничениям.

## 3) Approvals workflow (pending/declined)

### 3.1 Safe vs risky

- Изменения классифицируются как SAFE/RISKY (tracked-fields политика).
- SAFE правки в текущей реализации могут применяться напрямую к `KarteiRecord`.
- RISKY правки создают/обновляют `PendingChange`.
- Для stabilization v2 финансово значимые non-tracked billing/context changes не должны считаться safe только потому, что они не входят в legacy `TRACKED_FIELDS`.

### 3.2 Pending/declined edit workflow

- Current implementation: Admin редактирует `PENDING`/`DECLINED` через стандартный record editor.
- Для `PENDING` editor работает поверх существующего pending snapshot и обновляет pending proposal.
- Для `DECLINED` editor prefill'ится из `DeclinedChange.snapshot`; сохранение создает новый `PendingChange`.
- Operator в `PENDING`/`DECLINED` блокируется.
- Старые snapshot-only формы approvals могут оставаться compatibility layer, но не являются единственным workflow.

## 4) Импорт из Access (legacy_import)

### 4.1 Общие правила

- Источник: `.accdb` (обычно запускается на Windows-хосте из-за ACE ODBC).
- `ACCESS_CONN_STRING_TEMPLATE` собирается через Python `.format(file_path=...)`.
  Поэтому имя драйвера должно быть экранировано:
  `DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};`
- Маркерные строки "Zahlung" (визуальные разделители семей) не импортируются.

### 4.2 Patch-режим (`import_access_year --patch-fields`)

Назначение: "доимпорт" отдельных полей без полного реимпорта.

Контракт:
- читает только `tblKartei`,
- ищет существующие записи по `(year, id)`,
- обновляет только patch-поля (teachers/contract markers/SEPA),
- требует примененных миграций к той же БД (иначе `UndefinedColumn ...`).

## 5) Teachers / Subjects / Price (catalog)

- Учитель выбирается на каждую дисциплину отдельно (subject1/subject2).
- `extra1..3` считаем legacy и не развиваем.
- Прайс по умолчанию "за месяц".
- Для дисциплин с `Ind.` или `VSpE_` (индивидуальные) и `NH`/`Nachhilfe` в названии: цена "за академчас".

### 5.1 Pricelist v2 (категории и группы)

- Граница семестра по году: `catalog.SemesterConfig` (по умолчанию 1-6 / 7-12).
- Категории предметов на год: `catalog.SubjectCategory` (kind=GROUP/INDIVIDUAL) + привязка `catalog.SubjectCategoryLink` (1 предмет → не более 1 категории в год).
- Группы: `catalog.DisciplineGroup` (1 группа на `Subject` в год для GROUP-категорий) + история:
  - `DurationEntry` — длительность занятия, действует с месяца.
  - `GroupSizeEntry` — ручной размер группы, действует с месяца (в пределах семестра).
- Авторазмер группы считается из `karteien.KarteiRecord` на месяц с учетом контрактного статуса и billing-окна; ручной override имеет приоритет (см. `apps/catalog/group_size_service.py`).

## 6) AUTO начисления (month_1..month_12)

- `Fach 1 / Preis 1` относятся к 1-му полугодию, `Fach 2 / Preis 2` - ко 2-му.
- Для не-Ind/NH дисциплин месяцы полугодия считаются автоматически (цена за месяц + скидки).
- Для Ind/NH дисциплин админ вводит academic hours по месяцам (до 2 знаков), суммы считаются как `hours * price_per_hour`.
- При изменении цены админ выбирает диапазон применения в полугодии: `ab Monat` и опционально `bis Monat`.
- В `months_mode=AUTO` поле `base_amounts` хранит базовые суммы по месяцам (до скидок) и служит “историей цен”; при сохранении без биллинговых изменений оно не должно пересчитываться/перезаписываться.
- Месяцы до start_month должны быть `0.00` (не `NULL`).
- Ручное редактирование `month_*` запрещено в обычной форме (есть отдельный override-инструмент).
- Денежные поля нормализуются до 2 знаков, округление всегда вверх.
- Негативные суммы clamp-ятся к 0, с подтверждением в UI (по правилам проекта).

### 6.1 `months_mode` и правила редактирования месяцев

- `months_mode=LEGACY`: импортированные значения `month_*` считаются наследованными. В edit по умолчанию месяцы read-only (клик показывает объяснение).
- `months_mode=AUTO`: `month_*` рассчитываются из цен/скидок/UE и сохраняются в записи.
- `months_mode=OVERRIDE`: `month_*` считаются вручную заданными через `months-override` (аварийный режим).

Для legacy-записей доступна кнопка **Monate neu berechnen**: переводит запись в `AUTO` и пересчитывает месяцы по текущим правилам (с учётом цен/скидок/UE).

### 6.2 Подсветка “подозрительных” начислений (UI)

- В detail/edit месяцы, которые выглядят подозрительно относительно ожидаемой суммы по текущим ценам, подсвечиваются красной рамкой.
- Исключение: Stundenfächer (Ind./Nachhilfe/VSpE) не подсвечиваются, т.к. сумма зависит от UE.
- В `OVERRIDE` красная рамка не показывается; вместо этого рядом с месяцем отображается бейдж `Override`.

## 7) Скидки

- Процентные скидки хранятся как доля (`0.25` = 25%) и суммируются.
- Фиксированные (EUR) применяются после процентов и только когда base(month) > 0.
- Nachhilfe: скидки не применяются никогда.
- Есть флаг записи "скидки не применяются" (исключения).
- Скидки действуют по месяцам (по умолчанию с указанного месяца до конца года; поддерживаются "конкретные месяцы").

Важно для мастера "Neue Familie":
- допускается несколько `RecordDiscount` на одну запись;
- при назначении скидок нужно сразу пересчитать и сохранить `month_*` по правилам AUTO (без PendingChange);
- нельзя допустить двойного применения скидок при повторных пересчетах.

## 8) Контрактные маркеры (legacy Value14/Value20)

Хранение: raw text + вычисляемые флаги (чтобы не терять "другой текст в поле").

- Value14: `contract_type_raw`, `is_monthly_contract`
  - маркер `O/V` (case-insensitive), может быть вплотную к числам, ищется как подстрока.
- Value20: `contract_status_raw`, `is_contract_terminated`
  - маркер `KN` (case-insensitive), отдельный токен (границы строки/пробелы).

Дополнительно:
- `contract_terminated_from_month` (1-12): если контракт расторгнут, месяцы `>=` этого значения должны быть `0.00` в `AUTO` расчётах.

UI:
- отображение типа (месячный/годичный) и статуса (активен/разорван),
- фильтры/быстрая выборка.

## 9) FamilyID

- Формат: `"N. NNNN..."`, где N - цифра, после точки всегда пробел, количество цифр после пробела может расти.
- Нумерация должна быть глобальной (сквозной по годам).
- В мастере "Neue Familie": префикс `N.` пока фиксирован как `1.` (как в VBA).
- Резервации FamilyID:
  - админ может зарезервировать «следующий свободный» номер (номер + дата/время резервации),
  - автоподбор следующего номера должен учитывать активные резервации,
  - в мастере "Neue Familie" можно выбрать «использовать зарезервированный FamilyID».

## 10) UI и локализация

- UI тексты/сообщения: немецкий.
- Комментарии в коде: английский.
- `/karteien/`: list + live-search (AJAX), detail, create/edit для админов, read-only для superadmin/user.
  - фильтры списка включают Vertragstyp/Vertragsstatus (в т.ч. SEPA-варианты) и динамические фильтры Unterricht/Lehrer по семестру.
  - панель Filter может быть сворачиваемой (collapsed по умолчанию).
- `/karteien/create/`: autocomplete Eltern/Kind + prefill FamilyID/контактов + fallback по прошлым годам.
  - UX: подсказка возле FamilyID (max + рекомендация), предупреждение для Freitext.
  - entry-point из Family Dashboard: `prefill_family_id=...` (предзаполнение из существующей семьи, включая выбор из нескольких значений).
- Eintragsrabatte: выбор записи через autocomplete (вместо ручного ввода PKID).
- Мастер "Neue Familie": `/karteien/family/new/`.
- Family Dashboard: `/karteien/family/?year=YYYY&family_id=...` (в т.ч. запуск создания новой записи для семьи).
- Superadmin NeuList: `/approvals/superadmin/neu/?year=YYYY` (список “новых” записей per-year).
- Быстрые сценарии “починки каталога” из `/karteien/<pkid>/edit/`: create-формы справочников поддерживают предзаполнение через query-params и `next` возвратом обратно в редактирование.

## 11) Где искать детали

- Подробная архитектура: `ARCHITECTURE.md`
- Модель данных и маппинг Access Value-полей: `DOMAIN_MODEL.md`
- Пайплайн промптов и "почему это так": `PROMPTS_OVERVIEW.md` и `docs/Decisions.md`

## 12) Testing / QA contract

На дату 2026-05-27 после `PROMPT_179` в репозитории есть минимальная backend pytest-база:

- `pytest`, `pytest-django`;
- `pytest.ini`;
- `backend/tests/`.

После `PROMPT_182` есть PRICELIST V2 regression/xfail cases. После `PROMPT_185` есть minimal fixture-backed browser smoke layer через `pytest-playwright` + `pytest-django live_server`, но broad UI automation и MCP-runtime config не подтверждены. `package.json` не добавлен.

Контракт для дальнейшей разработки:

- `python backend/manage.py check` и `makemigrations --check --dry-run` — sanity-check, не функциональные тесты.
- `pytest -m "not browser" --reuse-db` запускает текущие backend regression/sanity tests; browser tests excluded by default through `pytest.ini`.
- `pytest -m browser backend/tests/browser --browser chromium --reuse-db` запускает minimal browser smoke и не заменяет manual browser QA.
- Нельзя считать UI/UX сценарии автоматически покрытыми за пределами явно перечисленного smoke scope.
- Не запускать backend и browser pytest параллельно против одной PostgreSQL test DB.
- Любое изменение поведения должно сопровождаться автоматическим тестом или явным manual QA checklist.
- Для UI/navigation/template/role/AJAX prompts использовать `docs/BROWSER_QA.md`; если browser verification не выполнен, prompt/final answer должен явно назвать blocker и оставшиеся ручные проверки.
