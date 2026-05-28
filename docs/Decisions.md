# Decisions (ADR-lite) - KindEltern Web

Этот файл фиксирует ключевые спорные решения, чтобы не возвращаться к ним каждый раз.

Дата актуализации: 2026-05-27.

Актуальный статус проекта и prompt-chain фиксируется в `CLAUDE.md` и `docs/Status.md`. Решения ниже могут иметь исторический контекст; если старый текст расходится с current/target блоками в `docs/Status.md`, приоритет у `docs/Status.md`.

Формат записи:
- Decision: что решили
- Context: почему возник вопрос
- Rationale: почему так
- Consequences: что это означает для кода/данных

## Current / Known Defects / Target

### Current implementation

- D03 фиксирует current hybrid workflow: Admin редактирует `PENDING`/`DECLINED` через стандартный record editor; Operator блокируется.
- D16 фиксирует канон документов для PRICELIST V2 stabilization.
- D17-D18 фиксируют целевые инварианты approvals/billing.

### Known defects

- Текущий PRICELIST V2 имеет split-brain между live record и pending snapshot.
- Billing preview, form save, category apply и approval path не полностью используют единый pipeline.
- Legacy contract fields и timeline entries еще не полностью сведены к одному operational source of truth.

### Target stabilization v2

- Frozen snapshot payload для billing/context changes.
- Live record остается approved-state до approval.
- `base_amounts` остается historical base history.
- Все поведенческие prompts обязаны обновлять `docs/Status.md` и давать test/manual QA evidence.

## D01 - Surrogate PK `pkid` в KarteiRecord

- Decision: `KarteiRecord.pk` = `pkid` (surrogate), а Access ID хранится в поле `id`; доменный ключ `(year, id)`.
- Context: Access ID не уникален по годам и неудобен как Django PK.
- Rationale: стабильные FK и корректная мультигодовая модель.
- Consequences: UI и API всегда явно различают `pkid` и Access `id`.

## D02 - `RecordStatus.NORMAL` хранится как пустая строка

- Decision: `NORMAL == ""`, проверки делаем как `not record.status`.
- Context: legacy хранение статуса.
- Rationale: минимальная несовместимость с импортом и существующими данными.
- Consequences: нельзя сравнивать со строкой `"NORMAL"`.

## D03 - Pending/Declined редактируются через стандартный record editor для Admin

- Decision: current implementation использует стандартный `/karteien/<pkid>/edit/` editor для Admin при `PENDING` и `DECLINED`; Operator в этих статусах блокируется.
- Context: snapshot-only flow оказался слишком узким для сложных форм, billing settings и UX исправлений.
- Rationale: единая форма уменьшает рассинхрон UI и позволяет исправлять pending/declined proposals через тот же validation/form stack.
- Consequences: старый variant A считать историческим/compatibility layer. Все рискованные изменения по-прежнему должны пройти approval Superadmin; прямой обход approval недопустим.

## D04 - Процентные скидки храним как долю

- Decision: percent скидки хранятся как `0.25` (25%), в UI отображаем как `25%`.
- Context: удобство вычислений и суммирования процентов.
- Rationale: упрощает формулы и снижает риск ошибок.
- Consequences: миграции/формы/шаблоны всегда учитывают формат доли.

## D05 - Nachhilfe: скидки не применяются

- Decision: для NH/Nachhilfe скидки не применяются никогда.
- Context: бизнес-правило.
- Rationale: исключение из общих правил скидок.
- Consequences: калькулятор начислений обязан учитывать исключение.

## D06 - AUTO начисления и отдельный override-инструмент

- Decision: `month_*` считаются автоматически, ручное редактирование запрещено в обычной форме; для исключений есть отдельный override-инструмент.
- Context: ручные правки ломают консистентность и мешают аудиту.
- Rationale: предсказуемость расчетов и единые правила.
- Consequences: UI формы должны блокировать ручной ввод `month_*`, кроме override.

## D07 - Patch-импорт `--patch-fields` ограничен полями и источником

- Decision: patch-импорт читает только `tblKartei` и обновляет только patch-поля у существующих записей, найденных по `(year, id)`.
- Context: нужно дописывать недостающие поля без полного реимпорта и без влияния на approvals таблицы.
- Rationale: минимальный риск повредить данные и workflow.
- Consequences: patch-режим не создает новые записи и не трогает остальную модель.

## D08 - Контрактные маркеры храним как raw text + derived flags

- Decision: Value14/Value20 сохраняются как сырой текст + вычисляемые булевы флаги.
- Context: в legacy полях мог быть "другой текст", который нельзя терять.
- Rationale: сохраняем исходник и получаем удобные фильтры.
- Consequences: при редактировании/импорте обновляется и raw, и флаги.

## D09 - Live search на `/karteien/` через AJAX

- Decision: live-search реализован как AJAX обновление таблицы (без SPA).
- Context: нужен быстрый поиск без тяжелого фронтенда.
- Rationale: минимальный JS, проще поддержка.
- Consequences: API для live-search должен проверять права и не раскрывать лишнего.

## D10 - UX скидок: shortcuts + autocomplete record picker

- Decision: админ назначает скидки прямо из create/edit записи; в Eintragsrabatte запись выбирается через autocomplete, а не через ручной ввод PKID.
- Context: ручной ввод PKID приводил к ошибкам и неудобству.
- Rationale: снижение риска привязки скидки не к той записи и ускорение работы.
- Consequences: UI должен показывать `pkid` явно и давать "говорящий" поиск по семье/ребенку.

## D11 - Языки и формат работы

- Decision: UI тексты - немецкий; комментарии в коде - английский; промпты - русский.
- Context: проект используется немецкоязычными пользователями; разработка идет на русском; код должен быть читаемым технически.
- Rationale: удобство пользователей и поддерживаемость.
- Consequences: любые новые UI-тексты добавляются на немецком.

## D12 - `months_mode`: LEGACY/AUTO/OVERRIDE и правила редактирования месяцев

- Decision: месячные поля `month_*` редактируются только косвенно (через цены/скидки/UE) в `AUTO`; в `LEGACY` по умолчанию read-only с пояснением; для ручных правок используется отдельный `months-override` (режим `OVERRIDE`).
- Context: прямое редактирование `month_*` в обычной форме приводило к конфликтам с ценами/скидками и усложняло аудит.
- Rationale: сохранить предсказуемость расчётов и при этом оставить аварийный ручной механизм.
- Consequences: UI в `/karteien/<pkid>/edit/` не должен показывать “спиннеры” для месяцев по умолчанию; “Override” доступен как отдельная операция.

## D13 - “Подозрительные начисления”: подсветка несовпадений с ожидаемой ценой

- Decision: в detail/edit подсвечиваем красной рамкой месяцы, где начисление подозрительно относительно ожидаемой суммы по текущим ценам (для не-Stundenfächer); в `OVERRIDE` рамку не показываем и вместо этого ставим бейдж `Override`.
- Context: в legacy-данных часто встречаются значения `month_*`, противоречащие цене/правилам, это нужно видеть сразу.
- Rationale: быстрый визуальный контроль без вмешательства в данные.
- Consequences: правило подсветки всегда исключает Stundenfächer (Ind./Nachhilfe/VSpE), где сумма зависит от UE и может отличаться от `Preis`.

## D14 - Изменение цены по диапазону месяцев: `ab Monat` / `bis Monat`

- Decision: при изменении цены админ задаёт диапазон применения в рамках полугодия (start/end), а UI показывает панель применения только если цена действительно изменена.
- Context: без “bis Monat” невозможно корректно исправлять цену “задним числом” без перезатирания будущих месяцев; постоянная панель применения дублировала основной UX и вводила в заблуждение.
- Rationale: минимальный, но практичный инструмент для реальных сценариев (включая backdated fixes).
- Consequences: в preview и при сохранении диапазон применяется только к затронутым месяцам соответствующего полугодия.

## D15 - Legacy значения `Fach/Lehrer/Preis`: бейджи + быстрые ссылки на каталоги

- Decision: если legacy-значения отсутствуют в каталогах, показываем под полями информативные бейджи и даём быстрые ссылки на create-формы справочников с предзаполнением + `next` возвратом.
- Context: “подмешивание” отсутствующих значений в dropdown как выбранный пункт маскирует проблему и может приводить к ошибочному сохранению/потере данных.
- Rationale: явная индикация нарушения процесса, но без разрушения legacy-данных при сохранении “других” правок.
- Consequences: если пользователь добавил нужные сущности в каталоги, бейджи исчезают и поля переходят на нормальный выбор из dropdown.

## D16 - Канон для PRICELIST V2 stabilization

- Decision: `PRICELIST_V2_STATE_AUDIT.md` считать current broken state, `PRICELIST_V2_STABILIZATION_TZ_REVISED.md` считать target state / master-spec.
- Context: документация и prompt-chain рассинхронизировались после итераций прайслиста v2.
- Rationale: агенту нужен один набор приоритетов, чтобы не чинить локально симптомы split-brain.
- Consequences: `ARCHITECTURE.md`, `DOMAIN_MODEL.md` и `docs/Spec.md` могут содержать historical text; их current/known defects/target блоки имеют приоритет.

## D17 - Snapshot v2 как целевой контракт approvals для billing/context

- Decision: для stabilization v2 финансово значимые billing/context changes должны храниться во frozen pending payload и применяться только после approval.
- Context: текущий код частично prewrite'ит non-tracked billing/context state в live record до решения Superadmin.
- Rationale: live record должен оставаться approved-state, иначе decline/approve становятся асимметричными.
- Consequences: prompt 166-178 должны постепенно убрать live prewrite paths и привести preview/save/apply/approve к единому payload contract.

## D18 - `base_amounts` в AUTO не является кэшем

- Decision: `base_amounts` в `months_mode=AUTO` — historical base history по месяцам.
- Context: category/group/duration/contract drift может менять suggested price, но не должен тихо переписывать утвержденную историю начислений.
- Rationale: бухгалтерский audit trail важнее удобства пересчета.
- Consequences: изменение `base_amounts` допускается только через явный action (`APPLY_CATEGORY`, `PRICE_OPTION_RECALC`, `LEGACY_TO_AUTO`, миграционный/admin action с явным scope).

## D19 - Prompt discipline: Status + test или manual QA

- Decision: каждый prompt, меняющий поведение приложения, должен обновить `docs/Status.md` и добавить автоматический тест или явный manual QA checklist.
- Context: предыдущая итерация прайслиста v2 создала значительный UI/UX и state рассинхрон без надежной фиксации статуса и проверки.
- Rationale: агент должен оставлять после себя проверяемый след и актуальную точку старта для следующего агента.
- Consequences: prompt без Status-update и test/QA checklist считается неполным, если только пользователь явно не ограничил задачу анализом.

## D20 - Backend pytest baseline и minimal browser smoke существуют, но broad UI automation/MCP не подтверждены

- Decision: после `PROMPT_179...PROMPT_185` документация должна различать backend pytest coverage, minimal `pytest-playwright` browser smoke, mandatory manual browser QA и фактическую MCP-конфигурацию runtime.
- Context: в репозитории есть `pytest`, `pytest-django`, `pytest-playwright`, `pytest.ini`, `backend/tests/` и `backend/tests/browser/`; после `PROMPT_182` есть PRICELIST V2 regression/xfail cases, после `PROMPT_185` есть fixture-backed non-destructive browser smoke. При этом `package.json` не добавлен, а `mcpServers` / Playwright MCP / Chrome DevTools MCP config в репозитории отсутствуют.
- Rationale: `manage.py check` и `makemigrations --check --dry-run` полезны как sanity-check, backend tests покрывают regression/invariant сценарии, а browser smoke покрывает только перечисленные role/route checks. UI/UX, AJAX, visual/responsive и destructive flows по-прежнему требуют manual browser QA или отдельной автоматизации.
- Consequences: default backend gate должен исключать browser tests (`pytest -m "not browser" --reuse-db`), browser smoke запускать явно (`pytest -m browser backend/tests/browser --browser chromium --reuse-db`), backend/browser pytest не запускать параллельно против одной PostgreSQL test DB, и не заявлять MCP evidence без реального MCP/runtime config.
