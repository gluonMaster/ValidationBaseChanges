# Логика прайслиста: расчёт цен и привязка к записям

> Документ описывает **текущее состояние** системы (актуализация: 2026-05-27), а не целевую stabilization v2.  
> Файлы реализации: `billing.py`, `pricing.py`, `category_price.py`, `catalog/pricing.py`.
>
> Важно: текущая реализация гибридная. `billing_preview_api()`/form-save всё еще идут через legacy billing path, а category pricing/apply живет отдельно в `catalog/pricing.py` + `karteien/category_price.py`. Category apply сейчас может prewrite'ить live `base_amounts` до Superadmin approval; это known defect, а не target-state. Target pipeline описан в `PRICELIST_V2_STABILIZATION_TZ_REVISED.md`.

---

## 1. Два пути ценообразования

Для каждой записи (`KarteiRecord`) и каждого семестра система **динамически** определяет источник цены:

```
Если subject*_ref заполнен И существует SubjectCategoryLink(subject, year, category.is_active=True):
    → CATEGORY  (новая логика: ставки категории)
Иначе:
    → PRICE_OPTION  (старая логика: PriceOption / price*_ref)
```

Источник определяет `determine_pricing_source()` в `catalog/pricing.py`. **Нигде не хранится** — вычисляется каждый раз.

---

## 2. Путь PRICE_OPTION (для некатегорийных дисциплин)

### Кто попадает сюда
- Дисциплина без `SubjectCategoryLink`
- Дисциплина без заполненного `subject*_ref` (legacy-записи)

### Расчёт базы (`build_base_amounts`)

Файл: `billing.py` → `build_base_amounts()`

Для каждого месяца (1–12):

1. Определяется семестр месяца (через `SemesterConfig` → `get_semester_month_ranges(year)`)
2. Берётся предмет и цена семестра:
   - Приоритет: `subject*_ref.name` / `price*_ref.amount`
   - Фолбэк: legacy `subject*` / `price*`
3. Определяется тип предмета:
   - **Per-month** (обычные): `base = price`
   - **Per-hour** (Individual/VSpE_/NH/Nachhilfe): `base = hours × price`
4. Если месяц вне диапазона `[start_month..end_month]` семестра → `base = 0.00`
5. Если задан `months_csv_*` → начисляются только перечисленные месяцы, `start/end` игнорируются
6. Округление: `CEILING` до 2 знаков

Результат: `dict {"month_1": Decimal, ..., "month_12": Decimal}` — **базы до скидок**.

---

## 3. Путь CATEGORY (для категорийных дисциплин)

### Кто попадает сюда
- Дисциплина привязана через `SubjectCategoryLink` к активной `SubjectCategory(kind=GROUP|INDIVIDUAL)`

### Предлагаемая цена (read-only, для отображения)

Файл: `catalog/pricing.py` → `calculate_suggested_price()`

**Это расчёт для UI, не записывается в базу автоматически.**

#### GROUP-дисциплина (месяц M):

```
1. contract_type(M) = тип договора в месяце M (ContractTypeEntry или fallback is_monthly_contract)
2. rate = category.monthly_rate  если contract_type=monthly
         category.yearly_rate    если contract_type=yearly
3. duration_minutes(M) = последний DurationEntry с effective_from_month <= M
4. ue_count = duration_minutes / 45
5. base_price = rate × ue_count
6. size(M) = авто-размер группы (см. раздел 5)
7. Если auto_scaling_enabled И 0 < size(M) < threshold:
       price = base_price × threshold / size(M)
   Иначе:
       price = base_price
8. Округление: CEILING до 2 знаков
```

#### INDIVIDUAL-дисциплина (месяц M):

```
rate = category.monthly_rate или category.yearly_rate (по типу контракта)
price = hours_amounts[M] × rate
(без часов — цена = 0)
```

### Применение к записи ("Применить с месяца X")

Файл: `karteien/category_price.py` → `apply_category_price_to_record()`

1. Рассчитывает предлагаемую цену для выбранных месяцев
2. Записывает новые значения в `record.base_amounts`
3. Сохраняет `_old_base_amounts` в snapshot (для отката при DECLINE)
4. Создаёт `PendingChange` → Superadmin одобряет/отклоняет
5. При DECLINE: `base_amounts` откатывается к старым значениям

**Массово** — через `BulkApplyCategoryPriceView` на странице группы.

Current-state caveat: шаги выше описывают существующий hybrid path. В текущем коде proposal может попасть в live `base_amounts` до решения Superadmin, а proposed `month_*` живут в `PendingChange.snapshot`; это создает split-brain между approved record state и pending proposal. Целевое состояние — держать весь billing/context proposal во frozen snapshot и применять его только после approval.

---

## 4. Финальный расчёт начислений (`calculate_month_values`)

Файл: `billing.py` → `calculate_month_values()`

После получения `base_amounts` (любым путём) применяется:

### Шаг 1: ContractStatus → обнуление

Для каждого месяца определяется `ContractStatusEntry` (или fallback):

| Статус | Начисление | В размере группы |
|--------|-----------|-----------------|
| ACTIVE | рассчитывается по шагам 2–4 | да |
| PAUSED | **0.00** (месяц пропускается) | да |
| TERMINATED | **0.00** (месяц пропускается) | нет |

> Это правило сильнее `months_mode=OVERRIDE`: даже ручной override = 0 в PAUSED/TERMINATED.

### Шаг 2: Проверка исключений скидок

Скидки **не применяются** если:
- `record.discounts_disabled = True` (и месяц в `discounts_disabled_months` или список пуст = все месяцы)
- Предмет месяца — Nachhilfe (`\bNH\b` или `Nachhilfe`)
- `base == 0.00`

### Шаг 3: Сбор скидок месяца

Собираются все `FamilyDiscount` (для семьи+год) и `RecordDiscount` (для конкретной записи), действующие в данном месяце.

### Шаг 4: Применение скидок

```
1. percent_sum = сумма всех %-скидок (clamp: max 99%)
2. after_percent = base × (1 − percent_sum)
3. fixed_sum = сумма всех €-скидок
4. after_fixed = after_percent − fixed_sum
5. Если after_fixed < 0 → clamp to 0.00 (с флагом предупреждения)
6. Округление: CEILING до 2 знаков
```

### Итоговый `month_1..month_12`

Результат `calculate_month_values()` — финальные начисления, которые записываются в `record.month_1..month_12`.

---

## 5. Размер группы

Файл: `catalog/group_size_service.py` → `get_group_size_for_month()`

Для `DisciplineGroup(subject=S, year=Y)` в месяце M:

```
size = COUNT записей WHERE:
  year = Y
  AND (
    (semester(M,Y)=1 AND subject1_ref=S AND M в активном диапазоне слота 1)
    OR
    (semester(M,Y)=2 AND subject2_ref=S AND M в активном диапазоне слота 2)
  )
  AND ContractStatus(M) IN {ACTIVE, PAUSED}  (TERMINATED не считается)
```

Фолбэк: если `subject*_ref` не заполнен — сравнивать по `subject1/2` (нормализация: trim + casefold).

**Ручной override**: если для месяца M есть `GroupSizeEntry` → использовать `manual_size` вместо авто-размера.

---

## 6. Режимы начислений (`months_mode`)

| Режим | Поведение |
|-------|-----------|
| `LEGACY` | Значения `month_*` вводятся вручную. `build_base_amounts()` не вызывается. "Применить цену" недоступно до конвертации. |
| `AUTO` | `base_amounts` хранит базы по месяцам. `month_*` пересчитываются через `calculate_month_values()` при любом billing-изменении. |
| `OVERRIDE` | `month_*` заданы вручную (через инструмент MonthsOverride). Могут быть перезаписаны при применении цены. |

Конвертация LEGACY → AUTO: кнопка "Monate neu berechnen" (создаёт PendingChange для Superadmin).

---

## 7. Семестры

Граница семестра определяется через `SemesterConfig(year)`:
- `last_month_sem1` (дефолт: 6)
- Семестр 1 = месяцы 1..boundary, семестр 2 = boundary+1..12

В записи `KarteiRecord`:
- `subject1_ref` + `price1_ref` → для месяцев семестра 1
- `subject2_ref` + `price2_ref` → для месяцев семестра 2
- `start_month_1/2`, `end_month_1/2`, `months_csv_1/2` — активные месяцы каждого слота

---

## 8. История цен (`base_amounts`)

`KarteiRecord.base_amounts` — JSONField вида `{"month_1": "100.00", ..., "month_12": "0.00"}`.

**Инвариант**: при AUTO-режиме это **история базовых цен по месяцам**. Нельзя молча перезаписывать прошлые месяцы.

При сохранении записи в AUTO-режиме без billing-изменений — `base_amounts` НЕ перезаписывается (сохраняется история).

---

## 9. Approval-поток для ценовых изменений

Все финансово значимые изменения проходят через approval:

| Действие | OPERATOR | ADMIN | SUPERADMIN |
|----------|----------|-------|------------|
| Изменить `base_amounts`/`month_*` | PENDING (только текущий/будущий месяц) | PENDING (любой месяц) | только одобряет |
| "Применить цену с месяца X" | PENDING (только текущий/будущий) | PENDING | только одобряет |
| ContractStatus/ContractType изменения | PENDING | PENDING | только одобряет |

При DECLINE: `base_amounts` откатывается к значениям до применения (через `_old_base_amounts` в snapshot).

Уточнение current-state: `_old_base_amounts` — transitional rollback metadata для отдельных apply-paths. Это не полноценный v2 frozen payload и не доказательство атомарности approval flow: non-tracked billing/context changes пока не проходят через единый reviewed payload.

---

## 10. Статус применения промптов 152–162

| Промпт | Описание | Применён |
|--------|----------|---------|
| PROMPT_152 | Fix: billing не падает на несохранённой записи (pk is None) | ✓ |
| PROMPT_153 | Fix: ACTIVE-restore учитывает предлагаемый entry при пересчёте snapshot | ✓ |
| PROMPT_154 | Fix: KarteiRecordForm без хардкода 1–6/7–12 | ✓ |
| PROMPT_155 | Fix: динамические подписи семестров в UI и Kosten-Report | ✓ |
| PROMPT_156 | Fix: auto-scaling блокируется при size=0 в группе | ✓ |
| PROMPT_157 | Fix: approvals edit forms сохраняют metadata-ключи snapshot (`_pending_*`) | ✓ |
| PROMPT_158 | Fix: ContractStatusChangeView ACTIVE не дублирует entry с одним effective_from_month | **НЕ применён** |
| PROMPT_159 | Fix: откат base_amounts при DECLINE применения категорийной цены | ✓ |
| PROMPT_160 | Fix: добавить "Nachhilfe" в DEFAULT_GROUP_CATEGORIES | **НЕ применён** |
| PROMPT_161 | Fix: Operator получает доступ к bulk apply (CatalogEditorMixin) | ✓ |
| PROMPT_162 | Fix: Operator-навигация к группам (DisciplineGroupListView, navbar, breadcrumbs) | **НЕ применён** |

**Не применены (3 промпта):** 158, 160, 162 — требуют применения.
