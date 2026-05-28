# PRICELIST V2 — Stabilization TZ (Revised, decision-complete)

Дата: 2026-04-13  
Статус: **готово для пошаговой реализации через последовательные AI prompts**  
Этот документ **заменяет** `PRICELIST_V2_STABILIZATION_SPEC.md` там, где есть расхождения.

---

## 0. Назначение документа

Цель этого ТЗ — стабилизировать PRICELIST V2 **без rewrite с нуля**, но с обязательным закрытием архитектурных дыр, которые уже подтверждены кодом:

- split-brain между live record и pending snapshot;
- конкуренция legacy contract bool-полей и timeline entries;
- рассинхрон между preview / form save / apply / approve / decline;
- нарушение инварианта `base_amounts` как historical base в `AUTO`;
- неполная approval semantics для non-tracked billing/context fields;
- schema/data conflict вокруг dual `Nachhilfe`.

Этот документ не пересобирает весь домен заново. Он **фиксирует точные решения**, чтобы дальнейшие prompt-цепочки для AI-агентов не принимали архитектурные решения по ходу реализации.

---

## 1. Подтверждённая база: что реально есть в коде

Ниже перечислено то, от чего надо отталкиваться как от validated current state:

- `backend/apps/karteien/billing.py` остаётся legacy billing engine для `build_base_amounts()`, `recalculate_record_months()`, `recalculate_legacy_to_auto()`, `detect_meaningful_changes()`.
- `backend/apps/catalog/pricing.py` уже считает category-based suggestion через `ContractTypeEntry`, `DurationEntry`, `GroupSizeEntry`.
- `backend/apps/karteien/category_price.py` умеет менять `base_amounts` и `month_*` отдельно от form/save pipeline.
- `backend/apps/karteien/api.py::billing_preview_api()` считает по отдельному legacy-path и не совпадает с category apply.
- `backend/apps/karteien/forms.py` и `backend/apps/karteien/views.py` по-прежнему позволяют general form менять legacy contract поля.
- `backend/apps/karteien/views.py::ContractTypeChangeView` и `ContractStatusChangeView` уже используют pending metadata (`_pending_contract_type_entry`, `_pending_contract_status_entry`).
- `backend/apps/approvals/services.py` строит `PendingChange.snapshot` только из `TRACKED_FIELDS`; approval/decline не умеют применять полный non-tracked billing/context payload.
- risky save-path в `KarteiRecordUpdateView` prewrite'ит часть non-tracked state через `_build_safe_fields_update()` прямо в live record до approval; в текущем коде туда ошибочно попадает и `sepa_marker`, хотя он влияет на permission/filter semantics.
- `ApplyCategoryPriceView` сейчас prewrite'ит `base_amounts` и `months_mode` в live record до approval.
- `BulkApplyCategoryPriceView` имеет аналогичный live prewrite category-apply state до approval.
- `QuickSetSubjectRefView` имеет отдельный direct `.update(subject*_ref_id, status=PENDING)` path вне `_build_safe_fields_update()`.
- `DisciplineGroupPrepareLegacyView` уже делает full-year `LEGACY_TO_AUTO`, но всё ещё prewrite'ит `subject*_ref_id`, `months_mode`, `base_amounts`, `hours_amounts` и связанный context в live record до approval.
- `MonthsOverrideView` уже является отдельным pending-path и тоже должен быть приведён к frozen payload semantics вместо частичной live-синхронизации.
- `SubjectCategory` сейчас имеет unique constraint `(year, name)`, а `ensure_default_categories(year)` работает по принципу “если в году уже есть хоть одна категория — ничего не создавать”.

---

## 2. Непереговорные инварианты

Это locked invariants. AI-агенты не должны их “оптимизировать” или пересматривать.

### 2.1 Billing truth

1. `base_amounts` в `months_mode=AUTO` — это **historical base history**, а не вычислимый cache.  
2. `month_*` — это **approved persisted truth** на записи.  
3. Новый набор `month_*` может быть сгенерирован только через общий billing pipeline.  
4. Нельзя получать состояние, где `base_amounts` и `month_*` относятся к разным расчётам.

### 2.2 Contract truth

1. Операционное чтение договора по месяцам идёт через:
   - `get_contract_type_for_month()`
   - `get_contract_status_for_month()`
2. `ContractTypeEntry` / `ContractStatusEntry` — целевая operational truth модель.  
3. Legacy contract fields временно остаются как fallback / compatibility layer, но **не как primary source**.

### 2.3 Approval truth

1. Пока change находится в `PENDING`, live record должен оставаться в **approved-state**.  
2. `APPROVED` обязан применять **frozen reviewed payload**, а не пересчитывать proposal заново.  
3. `DECLINED` для новых snapshot v2 в целевой модели не должен ничего откатывать в live billing/context state, потому что этот state не должен был быть prewrite'ен. На переходный период допускается compatibility rollback, если snapshot v2 всё ещё несёт legacy rollback markers от ещё не удалённого prewrite path.

### 2.4 Status semantics

1. `PAUSED` и `TERMINATED` дают `month_* = 0.00` **во всех режимах**, включая `OVERRIDE`.  
2. `PAUSED` участвует в group size, но не в billable count.  
3. `TERMINATED` не участвует в group size и даёт `0.00`.

### 2.5 No silent rewrite

1. Category/group/duration/contract-type drift влияет на suggestion/warnings, но **не переписывает** уже approved `base_amounts`.  
2. `base_amounts` меняются только через явный action:
   - `APPLY_CATEGORY`
   - `PRICE_OPTION_RECALC`
   - `LEGACY_TO_AUTO` (explicit action)
   - админский/миграционный сервис с явным scope.

---

## 3. Locked decisions

## 3.1 Единый snapshot contract для approvals

### Решение

`PendingChange.snapshot` остаётся JSON dict, **совместимым** со старым flat snapshot, но получает versioned v2-расширение.

### Формат snapshot v2

Top-level tracked fields остаются как сейчас:

- все поля из `TRACKED_FIELDS` на верхнем уровне snapshot;
- это нужно для обратной совместимости с существующим UI и approve-path.

Дополнительно snapshot v2 обязан содержать reserved keys:

```python
{
    "_snapshot_version": 2,
    "_pending_action": "FORM_SAVE" | "MONTHS_OVERRIDE" | "APPLY_CATEGORY" |
                       "PRICE_OPTION_RECALC" | "LEGACY_TO_AUTO" |
                       "CONTRACT_TYPE_CHANGE" | "CONTRACT_STATUS_CHANGE" |
                       "QUICK_SET_SUBJECT_REF" | "GROUP_PREPARE_LEGACY",
    "_pending_nontracked_payload": {...},
    "_rollback_nontracked_payload": {...} | None,
    "_pending_contract_type_entry": {...} | None,
    "_pending_contract_status_entry": {...} | None,
    "_pending_meta": {
        "touched_months": [1, 2, ...],
        "warnings": [...],
        "audit_summary": "...",
        "ui_summary": "...",
    },
}
```

### `_pending_nontracked_payload`

Это **frozen full post-approval state** для non-tracked полей, а не patch “на удачу”.  
Минимальный набор, который должен поддерживаться pipeline для billing-affecting actions:

```python
(
    "subject1_ref_id", "teacher1_ref_id", "price1_ref_id",
    "start_month_1", "end_month_1", "months_csv_1",
    "subject2_ref_id", "teacher2_ref_id", "price2_ref_id",
    "start_month_2", "end_month_2", "months_csv_2",
    "sepa_marker",
    "months_mode",
    "base_amounts",
    "hours_amounts",
    "legacy_base_amounts_enabled",
    "discounts_disabled",
    "discounts_disabled_months",
    "is_monthly_contract",
    "contract_type_raw",
    "is_contract_terminated",
    "contract_status_raw",
    "contract_terminated_from_month",
)
```

Для generic form risky save сюда должны попадать **все changed non-tracked form fields**, а не только те, что удобно сохранить через `_build_safe_fields_update()`.

`_rollback_nontracked_payload` в snapshot v2 не является целевой steady-state частью модели.  
Он допускается только как transitional compatibility payload для тех write-paths, которые во время rollout ещё не избавлены от live prewrite и поэтому обязаны уметь симметрично откатывать уже записанный non-tracked/context state.

### Обязательная backward compatibility

- snapshot без `_snapshot_version` считается legacy/v1;
- v1 должен продолжать читаться в PENDING/DECLINED UI;
- `apply_decision()` обязан поддерживать и v1, и v2;
- массовая миграция старых `PendingChange` / `DeclinedChange` snapshot'ов **не требуется**.

---

## 3.2 Approval semantics меняются: tracked diff больше не единственный критерий risk

### Решение

`TRACKED_FIELDS` остаётся legacy/history concept, но перестаёт быть единственным критерием того, нужен ли approval.

### Новый принцип

Approval обязателен, если proposal содержит хотя бы одно из:

1. tracked diff;
2. timeline diff (`_pending_contract_type_entry` / `_pending_contract_status_entry`);
3. billing-affecting non-tracked diff;
4. mode change (`LEGACY`/`AUTO`/`OVERRIDE`).

### Approval-relevant non-tracked diff

Минимально сюда входят:

- `subject*_ref_id`
- `teacher*_ref_id`, если изменение идёт вместе с risky proposal и должно быть frozen в reviewed payload
- `teacher*_ref_id`, если в том же запросе есть любой approval-relevant diff; teacher-only change без risky diff может оставаться safe direct-save
- `price*_ref_id`
- `start_month_*`
- `end_month_*`
- `months_csv_*`
- `months_mode`
- `base_amounts`
- `hours_amounts`
- `legacy_base_amounts_enabled`
- `discounts_disabled`
- `discounts_disabled_months`
- `sepa_marker`
- legacy contract fields

### Safe direct-save fields

По умолчанию safe direct-save допускается только для non-billing, non-contract полей.  
Для начала стабилизации к ним относятся:

- `teacher1_ref_id`
- `teacher2_ref_id`
- прочие changed non-tracked form fields, которые не влияют на billing / contract / permissions semantics.

`teacher1_ref_id` / `teacher2_ref_id` считаются safe direct-save **только если** тот же запрос не содержит:

- tracked diff;
- timeline diff;
- billing-affecting non-tracked diff;
- mode change.

Если teacher refs изменены **вместе** с risky proposal, они обязаны замораживаться внутри `_pending_nontracked_payload`, а не prewrite'иться в live record.

`sepa_marker` **не является safe direct-save field**, потому что меняет permission semantics (`record.is_sepa`) и влияет на list/API filters.

---

## 3.3 APPROVED/DECLINED работают только с frozen payload

### APPROVED

Для snapshot v2 порядок действий фиксируется так:

1. взять live approved record как baseline;
2. применить top-level tracked snapshot;
3. применить `_pending_nontracked_payload`;
4. создать/обновить contract entries из `_pending_contract_type_entry` / `_pending_contract_status_entry`;
5. сохранить запись;
6. перевести `status -> NORMAL`;
7. записать APR history entry.

### DECLINED

Для snapshot v2:

1. создать `DeclinedChange` с тем же snapshot;
2. перевести запись в `status -> DECLINED`;
3. **не менять** `base_amounts`, `month_*`, refs, mode, timeline и прочий live state;
4. записать DCL history entry.

### Transitional compatibility during rollout

Пока не закрыты все legacy prewrite paths из `Prompt 03`, `Prompt 05`, `Prompt 06` и `Prompt 07`, decline для snapshot v2 обязан сохранять compatibility branch:

- если v2 snapshot содержит явные rollback markers вроде `_old_base_amounts` или `_old_month_values`, decline всё ещё обязан откатить эти поля;
- это временное поведение нужно только для безопасного перехода через промежуточный релизный период;
- после удаления live prewrite paths новые v2 proposals больше не должны создавать такие markers.

### Path-scoped v2 rollout rule

Rollout snapshot v2 выполняется **по отдельным write-paths, а не глобально на весь проект сразу**.

Любой path, который всё ещё prewrite'ит live non-tracked / billing / context state, до своего исправления обязан подчиняться одному из двух режимов:

1. либо он остаётся на legacy/v1 decision semantics;
2. либо он использует v2 snapshot, но тогда обязан сохранять **полный** `_rollback_nontracked_payload` для всех live-prewritten полей и поддерживать симметричный compatibility rollback.

Следствие:

- `Prompt 02` не имеет права переключать такие paths на clean-v2 decline semantics преждевременно;
- clean-v2 semantics применяются только там, где конкретный write-path уже не делает live prewrite или уже несёт полный rollback payload.

### Legacy compatibility

Для snapshot v1 допустим legacy fallback:

- если в snapshot есть `_old_base_amounts`, текущая legacy decline-логика может его откатить;
- это поведение поддерживается только для обратной совместимости со старыми pending changes.

---

## 3.4 Pending / Declined UI обязаны уметь собирать projected state из snapshot v2

### Решение

Вводится общий helper уровня service, например:

- `build_projected_record_from_snapshot(record, snapshot)`

Этот helper обязан:

1. взять live record;
2. in-memory применить top-level tracked snapshot;
3. in-memory применить `_pending_nontracked_payload`;
4. отдельно вернуть projected timeline context для contract type/status, если в snapshot есть `_pending_contract_*_entry`.

### Почему это обязательно

Существующие detail/edit flows для `PENDING` / `DECLINED` сейчас умеют читать только top-level tracked snapshot.  
Без projection helper новый snapshot contract сломает:

- detail pending preview;
- edit form prefill для `PENDING` / `DECLINED`;
- корректный diff в UI.

---

## 3.5 Contract semantics: transitional policy и final target

### Transitional policy (Phase 1)

На внутреннем уровне системы правдой уже считается **резолв через helper'ы**:

- `get_contract_type_for_month()`
- `get_contract_status_for_month()`

То есть:

- list/API/detail/filter/read-paths внутри этого репо должны перейти на helper-based resolution уже в stabilization phase;
- helper имеет право fallback'иться к legacy fields для месяцев, не покрытых entries.

### Hard rollout dependency for non-lossless timeline states

Timeline proposals, которые **не могут быть losslessly представлены** через legacy compatibility summary, нельзя выпускать в approved/live semantics раньше helper-based migration внутренних read-paths.

Причина:

- legacy compatibility слой умеет честно различать только `ACTIVE` и `TERMINATED`;
- `PAUSED` не имеет корректного legacy summary representation;
- internal read-paths до `Prompt 09` всё ещё местами читают legacy bool/raw поля;
- если non-lossless timeline proposal уже можно approve/apply, а internal read-paths всё ещё читают legacy summary, система временно врёт в list/API/detail/filter/search.

Следствие:

- migration month-aware read-paths является **hard dependency** для rollout approved non-lossless timeline proposals;
- если prompt-order временно оставляет contract timeline actions раньше полной read migration, такие actions должны оставаться rollout-blocked:
  - нельзя показывать их как fully supported live behavior;
  - нельзя считать задачу завершённой, пока internal reads не перешли на helper-based resolution.
- rollout-blocked должен быть **технически enforced**, а не только описан в документации:
  - approval layer вводит явный guard, например `ALLOW_NONLOSSLESS_TIMELINE_APPROVAL = False` по умолчанию;
  - если pending timeline proposal не losslessly representable в legacy summary и guard выключен, `apply_decision(APPROVED)` обязан вернуть явную admin-facing ошибку и не применять change;
  - guard может быть снят только после завершения `Prompt 09`.

Минимальный набор blocked proposals до завершения `Prompt 09`:

- `_pending_contract_status_entry.kind == PAUSED`;
- любая `ACTIVE` reactivation после более раннего non-`ACTIVE` status;
- любой `_pending_contract_type_entry`, который создаёт intra-year contract type change.

### Final target (Phase 2)

Для записей, где timeline начал использоваться как primary model, вводится explicit baseline materialization.

### Baseline materialization algorithm

#### Contract type

Если для записи нет `ContractTypeEntry`, baseline materialization создаёт:

- `ContractTypeEntry(record=..., effective_from_month=1, is_monthly=record.is_monthly_contract)`

#### Contract status

Если для записи нет `ContractStatusEntry`, baseline materialization использует **только** `contract_terminated_from_month` как legacy-operational source.

Алгоритм:

- если `contract_terminated_from_month is None` → создать `ACTIVE @ 1`
- если `contract_terminated_from_month == 1` → создать `TERMINATED @ 1`
- если `contract_terminated_from_month in 2..12` → создать `ACTIVE @ 1` и `TERMINATED @ contract_terminated_from_month`

### Dirty legacy contract rows

Автоматическая baseline materialization **запрещена**, если найден один из кейсов:

- `is_contract_terminated=True`, но `contract_terminated_from_month` пустой;
- `contract_status_raw` содержит `KN`, но `contract_terminated_from_month` отсутствует;
- записи уже имеют partial/dirty timeline, который не проходит валидацию.

Для таких записей action должен завершаться blocking error + audit warning, без silent guesses.

### Legacy compatibility projection

Legacy contract fields остаются в записи как best-effort compatibility summary, но не используются как primary read-path после миграции list/API/detail/filter на helper'ы.

Legacy contract fields обновляются только тогда, когда approved timeline state losslessly representable через legacy model.  
Никакой синтетический encoding ради “эмуляции” `PAUSED`, intra-year contract type changes или иных non-lossless timeline states не допускается.

---

## 3.6 Contract actions не равны billing recalc

### ContractTypeChange

`CONTRACT_TYPE_CHANGE`:

- создаёт timeline proposal;
- не переписывает `base_amounts`;
- не переписывает `month_*`;
- меняет только contract truth и suggestion/warnings для будущих explicit apply.

Это обязательное решение. Нельзя silently recalc `base_amounts` только из-за смены типа договора.

### ContractStatusChange

`CONTRACT_STATUS_CHANGE`:

- всегда создаёт timeline proposal;
- для `PAUSED` / `TERMINATED` zeroes affected `month_*` in proposal;
- для `ACTIVE`:
  - `AUTO` → восстанавливает affected `month_*` из frozen `base_amounts` + discounts + projected status timeline;
  - `OVERRIDE` / `LEGACY` → **не делает silent restore**.

---

## 3.7 `base_amounts` history semantics и `LEGACY -> AUTO`

### Решение

Автоматическая partial-конверсия `LEGACY -> AUTO` через обычный form save запрещается.

### Почему

Текущий touched-month conversion ломает инвариант:

- `month_*` частично остаются legacy;
- `base_amounts` уже переписаны новым расчётом на все месяцы.

Это делает `AUTO` семантически нечестным.

### Новый policy

1. Обычный form save на `LEGACY` записи **не переводит** запись в `AUTO` автоматически.  
2. Если edit затрагивает billing-affecting fields и для корректного применения требует `AUTO`, generic form должен завершаться validation error с понятным CTA:  
   **сначала выполнить explicit `LEGACY_TO_AUTO`, потом делать billing-affecting edit**.
3. `LEGACY_TO_AUTO` существует только как **explicit dedicated action** и всегда пересчитывает **все 12 месяцев**, а не touched subset.
4. До появления отдельного dedicated UI-action допустим временный bridge через уже существующий form control `force_recalculate_months=True` (`Monate neu berechnen`):
   - это считается **explicit** `LEGACY_TO_AUTO`, а не обычным generic form save;
   - такой submit должен идти через тот же proposal/approval semantics, что и любой другой explicit `LEGACY_TO_AUTO`;
   - обычный LEGACY form submit без `force_recalculate_months=True` не имеет права silently конвертировать запись в `AUTO`.
5. Group-side helper `DisciplineGroupPrepareLegacyView` допускается, но только как explicit combined action, который тоже создаёт proposal полного `LEGACY_TO_AUTO`.

### AUTO / OVERRIDE generic-form billing rule

Для `AUTO` / `OVERRIDE` любой generic-form edit, который влияет на billed months или month values, обязан идти через один явный billing builder / pipeline action.

Минимальный набор таких полей:

- `subject*_ref_id`
- `price*_ref_id`
- `start_month_*`
- `end_month_*`
- `months_csv_*`
- `hours_amounts`
- `discounts_disabled`
- `discounts_disabled_months`
- `apply_from_month_*`

Следствие:

- такие edits нельзя оставлять на ad-hoc form math;
- если конкретное поле ещё не маршрутизировано через pipeline в текущей фазе, оно становится read-only / disabled в risky generic form flow;
- approve не имеет права “додумывать” missing billing recalculation задним числом.

### Последствие

Утрачивается “тихая частичная конверсия”, но сохраняется честная data semantics и предсказуемость approval flow.

---

## 3.8 `OVERRIDE` / `LEGACY` restore policy

### Решение

Автоматическое восстановление сумм при `PAUSED/TERMINATED -> ACTIVE` поддерживается только для `AUTO`.

Для `OVERRIDE` и `LEGACY`:

- silent restore запрещён;
- UI должен явно сообщать, что после возврата в `ACTIVE` суммы автоматически не восстановятся;
- пользователю должны быть предложены только явные варианты:
  - manual override;
  - explicit apply / recalculation, который переводит запись в `AUTO`.

### Explicit apply в `OVERRIDE`

`APPLY_CATEGORY` и `PRICE_OPTION_RECALC` в `OVERRIDE` разрешены, но это считается **осознанным выходом из manual-mode**:

- proposal обязан выставить `months_mode_after = AUTO`;
- UI обязан предупреждать: “Diese Aktion beendet den Override-Modus nach Genehmigung.”

---

## 3.9 Group size / warnings / pending projection

### Решение на stabilization phase

Полная cross-record pending projection для group size **не является Phase-1 deliverable**, но это formalized deferral, а не забытый кейс.

### До реализации pending projection обязательно

1. `group_size_service.py` и зависящие warnings/detail/bulk-preview остаются based on approved/live data only;
2. в group detail, record warnings и bulk preview должен появиться явный banner:
   - pending changes других детей в размере группы не учтены;
3. это должно быть закреплено как explicit non-goal Phase 1;
4. отдельный prompt позже может добавить Python overlay поверх live queryset.

### Что нельзя делать

Нельзя оставлять молчаливую неоднозначность, когда UI выглядит “как будто всё точно посчитано”.

---

## 3.10 Filters и UI

### Canonical filter semantics

Вводится canonical query param:

- `contract_month` = `1..12`

### Default

Если `contract_month` не передан, используется **текущий календарный месяц** как номер месяца.  
Это правило едино для list view и API, независимо от выбранного `year`.
То есть для просмотра исторического года по умолчанию всё равно используется текущий номер месяца календаря пользователя/сервера, а не месяц из выбранного года данных. `contract_month` — это selector month-slot (`1..12`), а не дата.

### Canonical filter params

Новый canonical набор:

- `contract_type=monthly|yearly`
- `contract_status=active|paused|terminated`
- `sepa=1|0` (optional)
- `contract_month=1..12`

### Backward compatibility

Старые combined values `active_sepa` / `terminated_sepa` разрешается временно парсить, но они считаются deprecated.

### Internal implementation rule

Внутренние list/API filters должны использовать helper-based month resolution.  
Для selected year разрешён Python-side filtering после prefetch timeline entries; не нужно пытаться втиснуть эту логику в fragile SQL до стабилизации.

### UI rule

Любой экран, который фильтрует по договору, обязан явно показывать выбранный `contract_month`.

---

## 3.11 `Nachhilfe` schema decision

### Решение

`SubjectCategory` unique constraint меняется с:

- `(year, name)`

на:

- `(year, name, kind)`

### Что обязательно меняется вместе с этим

1. `ensure_default_categories(year)` становится **idempotent additive**, а не “all-or-nothing if any category exists”.  
2. После schema migration нужен отдельный backfill/service, который:
   - проходит по существующим годам;
   - создаёт missing `GROUP`-category `Nachhilfe`, если в году уже есть `INDIVIDUAL` `Nachhilfe`, но нет `GROUP`-версии.
3. UI для списков/форм выбора категорий обязан показывать `name + kind`, а не только `name`.
4. Ordering в category UI должно быть минимум:
   - `order_by("name", "kind")`
   либо explicit separate columns `Name` / `Art`.

### Что не требуется

- никаких destructive data migrations не нужно, потому что старый constraint уже не мог допустить dual duplicate в одной БД.

---

## 3.12 History / audit semantics

### Проблема

`APR` diff сейчас строится только из `TRACKED_FIELDS`.  
Для v2 proposal возможны случаи, где:

- timeline change approval важен;
- `base_amounts` changed;
- tracked diff пустой или неполный.

### Решение

Snapshot v2 обязан нести `audit_summary` в `_pending_meta`.

На approve/decline:

- обычный tracked diff строится как раньше;
- если есть `audit_summary`, он добавляется в history comment;
- timeline-only approval не должен записываться как “APR без осмысленного описания”.

Пример допустимого `audit_summary`:

```text
TL[type@5=monthly; status@9=PAUSED]; BASE[months=9,10,11,12]; MODE[OVERRIDE->AUTO]
```

Не требуется менять legacy history parser под новые теги; достаточно стабильного structured suffix в comment.

---

## 4. Canonical action matrix

| Action | Approval | Меняет `base_amounts` | Меняет `month_*` | Меняет timeline | Меняет `months_mode` | Locked rule |
|---|---|---:|---:|---:|---:|---|
| `FORM_PREVIEW` | нет | нет | нет | нет | нет | только projection, без side effects |
| `FORM_SAVE` safe | нет | только если action не billing-affecting — иначе запрещено | нет | нет | нет | direct save only for safe fields |
| `FORM_SAVE` risky on `AUTO` | да | только если billing-affecting diff routed through explicit pipeline action; иначе field must be read-only | да, frozen tracked snapshot | возможно | возможно | no live prewrite |
| `FORM_SAVE` on `LEGACY` with billing-affecting diff | блокируется | нет | нет | нет | нет | сначала explicit `LEGACY_TO_AUTO` |
| `MONTHS_OVERRIDE` | да | нет | да | нет | да (`-> OVERRIDE`) | months_mode must be in snapshot payload |
| `APPLY_CATEGORY` | да | да | да | нет | возможно (`OVERRIDE -> AUTO`) | live record stays approved until approve |
| `PRICE_OPTION_RECALC` | да | да | да | нет | возможно (`OVERRIDE -> AUTO`) | internal pipeline action for explicit price-option recalculation |
| `LEGACY_TO_AUTO` | да | да (all 12) | да (all 12) | нет | да (`-> AUTO`) | only explicit action |
| `CONTRACT_TYPE_CHANGE` | да | нет | нет | да | нет | no implicit recalc |
| `CONTRACT_STATUS_CHANGE` | да | нет | да (affected range only) | да | нет | `ACTIVE` restore only in `AUTO` |
| `QUICK_SET_SUBJECT_REF` | да | нет | нет | нет | нет | FK change lives only in snapshot payload |
| `GROUP_PREPARE_LEGACY` | да | да (full LEGACY_TO_AUTO) | да | нет | да (`-> AUTO`) | explicit combined action |
| `APPROVAL_APPLY` | — | применяет frozen payload | применяет frozen payload | применяет frozen payload | применяет frozen payload | no recompute |
| `APPROVAL_DECLINE` v2 | — | нет | нет | нет | нет | live state unchanged for clean v2; temporary compatibility rollback allowed only if legacy rollback markers still exist |

---

## 5. Shared implementation core

Не надо начинать с большого rewrite. Но нужен единый reusable core.

## 5.1 Обязательные shared services

### A. `backend/apps/karteien/services/pending_snapshot.py`

Обязанности:

- сериализация snapshot v2;
- backward-compatible чтение v1/v2;
- `build_projected_record_from_snapshot()`;
- `apply_snapshot_v2_to_record()`;
- helper'ы для `_pending_nontracked_payload`.

### B. `backend/apps/karteien/services/billing_pipeline.py`

Обязанности:

- shared builders для actions, которые реально меняют billing state;
- формирование frozen payload для approval;
- reuse существующих low-level calculators, а не переписывание математики заново.

Минимальные builders:

- `build_apply_category_proposal(...)`
- `build_price_option_recalc_proposal(...)`
- `build_legacy_to_auto_proposal(...)`
- `build_contract_status_proposal(...)`
- `build_contract_type_proposal(...)`
- `build_months_override_proposal(...)`
- `build_quick_set_subject_ref_proposal(...)`

`PRICE_OPTION_RECALC` в этом ТЗ — **internal pipeline action**, а не обязательный новый standalone URL в Phase 1.

Его источник в текущем UI/коде:

- generic form save для `AUTO` / `OVERRIDE`, когда меняются `price*_ref` и пользователь явно задаёт `apply_from_month_*`;
- при необходимости later phase может вынести это в отдельный explicit endpoint, но Phase 1 этого не требует.

Ограничения:

- на `LEGACY` обычный form save не может порождать `PRICE_OPTION_RECALC`;
- `force_recalculate_months=True` на `LEGACY` маршрутизируется как explicit `LEGACY_TO_AUTO`, а не как `PRICE_OPTION_RECALC`.

### C. Reused low-level calculators

Остаются как transitional ядро:

- `billing.py::calculate_month_values()` — low-level month finalizer;
- `billing.py::build_base_amounts()` — PRICE_OPTION strategy base builder;
- `catalog/pricing.py` — CATEGORY suggestion strategy.

### D. Вспомогательный contract service

Нужен helper, например:

- `backend/apps/karteien/services/contract_timeline.py`

Обязанности:

- projected timeline merge для pending preview;
- baseline materialization (Phase 2);
- legacy compatibility validation / dirty state detection.

---

## 6. Прямые ограничения на существующие write paths

До конца стабилизации запрещены следующие практики:

1. `KarteiRecordUpdateView._build_safe_fields_update()` больше не может быть источником live prewrite для risky billing/context changes.
2. `ApplyCategoryPriceView` / `BulkApplyCategoryPriceView` не могут писать `base_amounts` в live record до approval.
3. `QuickSetSubjectRefView` не может писать `subject*_ref_id` в live record до approval.
4. `MonthsOverrideView` не может создавать pending без frozen `months_mode=OVERRIDE` в snapshot payload.
5. `ContractTypeChangeView` / `ContractStatusChangeView` не могут полагаться на legacy bool-поля как на read truth для list/API/filter после миграции read-paths.
6. `billing_preview_api()` не может продолжать собственную логику touched-month / termination-preview отдельно от pipeline.
7. Generic form для existing records не является primary editor для contract type/status.

---

## 7. Последовательная prompt-декомпозиция для AI-агентов

Ниже canonical prompt order. Каждый следующий prompt предполагает, что locked decisions из этого ТЗ уже приняты и не обсуждаются заново.

## Prompt 01 — Snapshot v2 foundation

**Цель:** ввести snapshot v2 без поломки legacy snapshot v1.

**Файлы:**
- `backend/apps/approvals/services.py`
- новый `backend/apps/karteien/services/pending_snapshot.py`
- `backend/apps/karteien/views.py` (pending/declined projection helper integration)

**Сделать:**
- ввести `_snapshot_version=2`;
- добавить `_pending_nontracked_payload` и `_pending_meta`;
- сделать helper `build_projected_record_from_snapshot()`;
- сохранить backward compatibility для v1.

**Нельзя ломать:**
- existing PENDING/DECLINED records;
- current detail/edit ability for old snapshots.

**Acceptance:**
- old snapshot continues to render;
- new snapshot can project tracked + non-tracked state in-memory.

## Prompt 02 — Approval apply/decline refactor

**Цель:** перевести approvals на frozen payload semantics.

**Файлы:**
- `backend/apps/approvals/services.py`
- tests for approve/decline

**Сделать:**
- approve for v2 applies tracked + nontracked + timeline payload;
- decline for v2 leaves live record unchanged except `status=DECLINED`, но сохраняет transitional compatibility rollback branch для v2 snapshots, которые ещё несут explicit legacy rollback markers;
- v1 legacy rollback remains only as compatibility branch.
- rollout v2 clean-decline semantics сделать path-scoped, а не глобальным;
- ввести явный approval guard для non-lossless timeline proposals: пока month-aware read migration не завершена, approve для таких proposals запрещён.

**Acceptance:**
- `APPROVED` не пересчитывает proposal заново;
- `DECLINED` v2 не трогает live billing/context state для новых clean proposals;
- если v2 snapshot пришёл из ещё не устранённого prewrite path и несёт rollback markers, compatibility rollback остаётся рабочим до завершения Prompt 03/05/06/07;
- path, который ещё prewrite'ит live state, не считается переведённым на clean-v2 semantics, пока не перестанет prewrite'ить или не начнёт сохранять полный rollback payload.

## Prompt 03 — Pending actions without live prewrite

**Цель:** закрыть неполные pending write paths, которые не используют non-tracked payload.

**Файлы:**
- `backend/apps/karteien/views.py`

**Scope:**
- `MonthsOverrideView`
- `ContractTypeChangeView`
- `ContractStatusChangeView`
- `QuickSetSubjectRefView`

**Сделать:**
- все эти views создают snapshot v2;
- `months_mode`, refs и contract legacy compatibility fields попадают в frozen payload;
- live record получает только `status=PENDING`.
- `ContractStatusChangeView` и `ContractTypeChangeView` не считаются fully rolled out для non-lossless timeline proposals, пока не завершён prompt по month-aware read-path migration.
- contract action paths не имеют права live-prewrite'ить legacy contract compatibility fields; если такие поля нужны для compatibility summary, они живут только во frozen payload до approve.

**Acceptance:**
- months override после approve реально переводит запись в `OVERRIDE`;
- quick-set subject_ref не меняет live FK до approve.
- non-lossless timeline proposals остаются rollout-blocked до завершения internal helper-based reads;
- пока rollout guard не снят, approve non-lossless timeline proposals технически блокируется в approval layer, а не только скрывается в UI.

## Prompt 04 — Billing pipeline core for explicit actions

**Цель:** ввести shared proposal builders для billing-affecting actions.

**Файлы:**
- новый `backend/apps/karteien/services/billing_pipeline.py`
- `backend/apps/karteien/category_price.py` (reuse or thin wrapper)
- `backend/apps/karteien/billing.py` (только минимальный extract/reuse)

**Сделать:**
- builder для `APPLY_CATEGORY`;
- builder для `PRICE_OPTION_RECALC`;
- builder для `LEGACY_TO_AUTO` (full-year only);
- builder для `MONTHS_OVERRIDE`.
- зафиксировать, что `PRICE_OPTION_RECALC` — internal action generic form flow для explicit price changes (`price*_ref` + `apply_from_month_*`), а не отдельный обязательный endpoint этой фазы.

**Acceptance:**
- builders возвращают frozen proposal, а не пишут в БД;
- pipeline может работать без form/view context.

## Prompt 05 — Single-record apply + preview parity

**Цель:** перевести single apply и его preview на pipeline.

**Файлы:**
- `backend/apps/karteien/views.py`
- `backend/apps/karteien/category_price.py`

**Scope:**
- `ApplyCategoryPriceView`
- `apply_price_preview`

**Сделать:**
- убрать live prewrite `base_amounts`;
- `OVERRIDE -> AUTO` делать только внутри proposal;
- preview использовать тот же builder, что и apply.

**Acceptance:**
- preview и фактический pending proposal дают одинаковые числа.

## Prompt 06 — Bulk apply + prepare legacy

**Цель:** перевести group-side mass actions на тот же pipeline.

**Файлы:**
- `backend/apps/catalog/views.py`

**Scope:**
- `BulkApplyCategoryPriceView`
- `BulkApplyPreviewView`
- `DisciplineGroupPrepareLegacyView`

**Сделать:**
- bulk apply использует same proposal builder;
- group prepare делает explicit combined `QUICK_SET_SUBJECT_REF + LEGACY_TO_AUTO`, без live prewrite;
- зафиксировать, что проблема `DisciplineGroupPrepareLegacyView` в текущем коде — не touched-month semantics, а live prewrite до approval при уже full-year recalculation.

**Acceptance:**
- bulk preview совпадает с per-record proposal logic;
- legacy prepare не создаёт split-brain между pending snapshot и live record.

## Prompt 07 — General form save stabilization

**Цель:** убрать `_build_safe_fields_update()` как источник split-brain и закрыть automatic LEGACY->AUTO conversion.

**Файлы:**
- `backend/apps/karteien/forms.py`
- `backend/apps/karteien/views.py`

**Сделать:**
- risky generic save создаёт snapshot v2, а не prewrite live non-tracked state;
- existing record contract section становится read-only summary + links to dedicated actions;
- automatic LEGACY->AUTO via generic form save удаляется;
- billing-affecting edit on LEGACY blocks with explicit CTA;
- любой generic-form change на `AUTO` / `OVERRIDE`, который влияет на billed months или month values, маршрутизируется через явный billing builder / pipeline action;
- минимально это включает `subject*_ref_id`, `price*_ref_id`, `start_month_*`, `end_month_*`, `months_csv_*`, `hours_amounts`, `discounts_disabled*`, `apply_from_month_*`;
- если конкретное billing-affecting поле ещё не заведено в pipeline текущей фазы, оно становится read-only / disabled в risky generic form flow;
- explicit price changes с `apply_from_month_*` маршрутизируются через internal `PRICE_OPTION_RECALC` builder, а не через отдельную ad-hoc математику формы;
- `force_recalculate_months=True` остаётся только как temporary explicit bridge к `LEGACY_TO_AUTO` и должен идти через тот же proposal/approval flow, а не через обычный risky generic save.

**Acceptance:**
- no risky path writes billing/context truth into live record before approve;
- existing record form no longer edits contract truth directly;
- generic form не оставляет billing-affecting fields в “полуручном” состоянии между ad-hoc form math и pipeline logic.

## Prompt 08 — `billing_preview_api` parity

**Цель:** убрать отдельный preview engine.

**Файлы:**
- `backend/apps/karteien/api.py`
- `backend/apps/karteien/services/billing_pipeline.py`

**Сделать:**
- `billing_preview_api()` строит projected proposal тем же shared core;
- contract status preview использует same restore/zero rules;
- contract type preview не пересчитывает `base_amounts` silently.
- любой UI flow, который показывает billing preview для form/apply/status actions, не считается rollout-complete до завершения этого prompt; до этого preview должен быть либо скрыт, либо явно помечен как non-authoritative.

**Acceptance:**
- preview/save/apply/status-change share one core semantics.

## Prompt 09 — Month-aware contract read paths

**Цель:** выровнять list/API/detail/filter semantics.

**Файлы:**
- `backend/apps/karteien/views.py`
- `backend/apps/karteien/api.py`
- `backend/apps/accounts/views.py`
- `backend/templates/karteien/_record_list_table.html`
- `backend/templates/karteien/record_detail.html`
- `backend/templates/karteien/record_form.html`
- `backend/templates/accounts/user_search.html`
- `backend/templates/accounts/user_record_detail.html`
- templates for filter/search UI

**Сделать:**
- выполнить inventory/grep всех внутренних read-paths по:
  - `is_monthly_contract`
  - `is_contract_terminated`
  - `contract_terminated_from_month`
  - прямому legacy contract summary без helper resolution;
- canonical `contract_month`;
- `paused` support in filters/UI;
- internal read paths switch to helper-based month resolution;
- detail summary, list badges, search results и user-facing internal pages перестают использовать legacy bool как primary source of truth;
- deprecated combined `*_sepa` parsing remains only as compatibility parser.
- прямые legacy reads допускаются только в allowlist:
  - compatibility fallback внутри helper'ов;
  - baseline materialization;
  - migration/backfill code;
  - debug/admin diagnostics, если это явно оговорено.
- после завершения month-aware read migration снимается rollout guard для approved non-lossless timeline proposals.

**Acceptance:**
- grep по внутреннему репо не находит новых primary read-paths legacy contract fields вне allowlist;
- `PAUSED` корректно отображается в list/API/detail/filter/search;
- detail summary и list badges не используют legacy bool как primary source of truth.

## Prompt 10 — Contract baseline materialization + dirty-state audit

**Цель:** подготовить contract timeline к полному operational truth.

**Файлы:**
- новый `backend/apps/karteien/services/contract_timeline.py`
- contract-related views/tests
- optional management command for audit

**Сделать:**
- baseline materialization algorithm;
- dirty legacy contract detection;
- safe blocking errors instead of silent guesses.

**Acceptance:**
- first timeline consolidation on clean rows is deterministic;
- dirty rows are reported, not silently materialized.

## Prompt 11 — `Nachhilfe` schema + bootstrap/backfill

**Цель:** закрыть conflict между ТЗ и схемой данных.

**Файлы:**
- `backend/apps/catalog/models.py`
- `backend/apps/catalog/services.py`
- relevant catalog views/templates/tests
- migration files

**Сделать:**
- unique constraint `(year, name, kind)`;
- additive idempotent `ensure_default_categories(year)`;
- backfill missing `GROUP Nachhilfe` for existing years;
- show `kind` in category UI.

**Acceptance:**
- one year can contain both `Nachhilfe (GROUP)` and `Nachhilfe (INDIVIDUAL)`;
- bootstrap/backfill no longer depends on “empty year only”.

## Prompt 12 — Cleanup, warnings, test matrix

**Цель:** удалить transitional write paths и закрыть основные regression risks.

**Файлы:**
- impacted modules from previous prompts
- test suite
- docs/prompt references

**Сделать:**
- formal Phase-1 banner about missing pending group-size projection;
- remove dead prewrite helpers;
- update docs and prompt references;
- add required tests.

**Acceptance:**
- no known split-brain write path remains;
- main invariants covered by tests.

---

## 8. Обязательный pre-implementation audit

До начала code changes нужен хотя бы read-only audit script / command, который найдёт:

1. `PendingChange` / `DeclinedChange` со snapshot v1;
2. записи со статусом `DECLINED`, где legacy rollback уже мог оставить split-brain;
3. открытые `PENDING` записи, созданные legacy prewrite paths, где live non-tracked / billing / context state уже изменён и не полностью представлен в snapshot;
4. dirty legacy contract rows:
   - `is_contract_terminated=True`, но month пустой;
   - `contract_status_raw` содержит `KN`, но month пустой;
   - partial timeline anomalies;
5. годы, где после schema change потребуется добавить missing `GROUP Nachhilfe`.

Для найденных legacy `PENDING` split-brain rows нужен отдельный compatibility/repair decision до rollout новых approval semantics; просто “оставить как есть” нельзя.

Audit должен быть read-only. Repair — отдельный controlled step.

---

## 9. Минимальный test matrix

Ниже не wish-list, а обязательный минимум.

### 9.1 Approvals

- snapshot v1 backward compatibility;
- snapshot v2 projection applies tracked + nontracked payload;
- approve applies frozen payload without recompute drift;
- decline v2 leaves live state unchanged for clean v2 proposals; temporary rollback remains only for compatibility with not-yet-removed legacy prewrite paths;
- path-scoped rollout: write-paths with remaining live prewrite either stay on legacy semantics or carry full rollback payload;
- approve/decline remain symmetric for category apply / quick subject ref / months override / contract actions.

### 9.2 Billing atomicity

- category apply no longer prewrite'ит `base_amounts`;
- bulk apply behaves the same as single apply;
- `OVERRIDE -> AUTO` only via explicit apply/recalc proposal;
- months override persists `months_mode=OVERRIDE` after approve.

### 9.3 Contract semantics

- contract type change does not rewrite `base_amounts` or `month_*`;
- contract status `PAUSED/TERMINATED` zeroes months in all modes;
- contract status `ACTIVE` restores only in `AUTO`;
- rollout guard blocks non-lossless timeline proposals before Prompt 09;
- month-aware filters (`contract_month`) for active/paused/terminated.

### 9.4 LEGACY policy

- generic form cannot silently convert `LEGACY -> AUTO`;
- explicit `LEGACY_TO_AUTO` recalculates full year only;
- `force_recalculate_months=True` routes only as explicit `LEGACY_TO_AUTO`;
- billing-affecting generic-form fields on `AUTO` / `OVERRIDE` either route through pipeline or are read-only;
- group prepare uses same explicit semantics.

### 9.5 `Nachhilfe`

- migration allows `(same year, same name, different kind)`;
- additive bootstrap/backfill creates missing `GROUP Nachhilfe`;
- UI renders categories unambiguously.

---

## 10. Explicit non-goals

Это специально **не** входит в первую стабилизацию:

1. полный rewrite всех billing calculators;
2. удаление legacy contract fields из схемы;
3. полная cross-record pending projection для group size;
4. destructive rewrite старых pending/declined snapshot'ов;
5. silent restoration of manual modes;
6. автоматическое пересчитывание `base_amounts` при любом category/group drift.

---

## 11. Legacy prompt mapping

- `PROMPT_158` больше нельзя применять как isolated fix: он должен жить внутри нового `CONTRACT_STATUS_CHANGE` proposal path.
- `PROMPT_160` нельзя применять отдельно от schema migration и bootstrap/backfill rules.
- `PROMPT_162` остаётся безопасным независимым UX fix и может быть выполнен рано.

---

## 12. Итоговая формулировка ready-state

Стабилизация считается завершённой, только если одновременно выполняются все условия:

1. нет write path, который prewrite'ит billing/context truth в live record до approval;
2. snapshot v2 покрывает tracked + non-tracked + timeline payload;
3. `APPROVED` применяет frozen payload без recompute drift;
4. `DECLINED` v2 не оставляет split-brain live state;
5. `billing_preview_api`, form save, single apply, bulk apply, months override и contract status change используют один shared core semantics;
6. existing record general form больше не редактирует contract truth напрямую;
7. internal list/API/detail/search/filter read-paths month-aware и helper-based;
8. `LEGACY -> AUTO` больше не происходит молча и частично;
9. `OVERRIDE` / `LEGACY` честно не делают auto-restore при `ACTIVE`;
10. schema+bootstrap для dual `Nachhilfe` больше не конфликтуют.
