# PRICELIST V2 — Edge Cases Checklist

Этот чеклист нужен для GPT PRO как список обязательных сценариев, которые нельзя потерять при упрощении логики.

## 1. Границы семестров и активность месяцев

- `SemesterConfig` может задавать не только 1..6 / 7..12.
- `start_month_*` может быть позже первого месяца семестра.
- `end_month_*` может ограничивать семестр раньше его конца.
- `months_csv_*` может задавать непоследовательные месяцы.
- `months_csv_*` может конфликтовать с `start/end`, и правило приоритета должно быть единым.
- Если предмет или цена отсутствуют только в одном семестре, второй семестр не должен ломаться.
- Месяцы вне активного окна должны давать нулевую базу.

## 2. Источник цены

- В одном семестре запись может быть на `CATEGORY`, а в другом на `PRICE_OPTION`.
- `CATEGORY` должен включаться только при наличии активного `SubjectCategoryLink`.
- При отсутствии category link должен быть корректный fallback в `PRICE_OPTION`.
- Если category link появился после того, как `base_amounts` уже были сохранены, старая база не должна молча перезаписаться.
- Suggestion и applied base history должны быть разными понятиями.

## 3. GROUP vs INDIVIDUAL

- GROUP использует duration, contract type и effective group size.
- INDIVIDUAL использует hourly base по `hours_amounts`.
- `Nachhilfe` может быть и групповым, и индивидуальным сценарием, но текущая схема `SubjectCategory(year, name)` это не всегда позволяет.
- Правила per-hour должны быть одинаковыми в billing, preview и warnings.

## 4. Размер группы

- В effective size считаются `ACTIVE + PAUSED`.
- `TERMINATED` не считаются.
- Billable count — только `ACTIVE`.
- Manual size override должен побеждать auto size только в своём месяце/семестре.
- Если manual size пуст, должен использоваться auto size.
- Если auto size = 0, scaling и price suggestion должны вести себя предсказуемо.
- Если группа есть, но нет billable учеников, UI и расчёт должны показать корректный warning.
- Изменение статуса одного ребёнка должно влиять на size других детей той же группы в тех же месяцах.

## 5. Scaling

- Scaling включается только для GROUP.
- Scaling не должен включаться, если effective size <= 0.
- Scaling не должен применяться, если size >= threshold.
- При size < threshold цена должна быть вычислена стабильно и одинаково в preview, warnings и apply.
- Toggle scaling не должен быть доступен пользователям без прав.

## 6. Duration

- Если для GROUP нет duration entry на активный месяц, это warning.
- Если есть gap в duration timeline, suggestion должен быть неполным, но система не должна silently invent values.
- Если duration меняется посреди семестра, affected months должны быть только соответствующие месяцы.

## 7. Contract Type Timeline

- В одном году договор может быть yearly в одни месяцы и monthly в другие.
- При смене contract type в том же effective month должен корректно заменяться существующий entry.
- Расчёт ставки yearly/monthly должен быть консистентен в category suggestion, preview и final apply.
- Mixed contract types внутри одной группы должны корректно отображаться в group detail.

## 8. Contract Status Timeline

- `ACTIVE` даёт начисление.
- `PAUSED` входит в group size, но не должен давать billable month value.
- `TERMINATED` не входит в group size и должен давать ноль.
- Смена статуса в том же месяце должна заменять существующий entry, а не создавать дубль.
- Возврат из `PAUSED` в `ACTIVE` должен восстанавливать начисление из сохранённой базы, а не из нуля.
- Возврат из `TERMINATED` в `ACTIVE` должен работать симметрично.

## 9. `base_amounts` и финальные месяцы

- `base_amounts` в `AUTO` — историческая база, не просто вычислимый кэш.
- `month_*` всегда должны быть производными от `base_amounts` + discounts + status logic.
- Нельзя получить состояние, где `base_amounts` и `month_*` относятся к разным расчётам.
- Частичное обновление touched months не должно ломать untouched months.
- LEGACY -> AUTO conversion должна сохранять нетронутые legacy months, если так задумано workflow.

## 10. Discounts

- FamilyDiscount и RecordDiscount могут действовать одновременно.
- Процентные скидки применяются раньше фиксированных.
- Итог не должен уходить в минус.
- `discounts_disabled` может отключать скидки глобально.
- `discounts_disabled_months` может отключать скидки только частично.
- Отключение скидок должно одинаково работать в preview и final save.
- Nachhilfe/individual rules не должны случайно обходить скидки, если это не задано явно.

## 11. Approval Flow

- Category apply не должен оставлять live record в полусохранённом состоянии при pending.
- Decline должен откатывать все связанные billing-поля, а не только часть.
- Approval должен создавать/обновлять timeline entries без дублей и без потери старой истории.
- Повторный risky edit поверх pending должен обновлять snapshot консистентно.
- Declined -> re-edit -> pending снова не должен приводить к рассинхрону record/snapshot.

## 12. UI / Preview / Permissions

- Live preview должен совпадать с тем же pipeline, который потом применится при save/apply.
- Operator не должен видеть кнопки, ведущие на гарантированный 403.
- Breadcrumbs и навигация должны вести только на доступные страницы.
- Group bulk preview и record preview должны показывать одинаковые цифры для одной и той же базы.
- UI должен явно предупреждать, если category apply игнорирует несохранённые изменения формы.

## 13. Data Model Conflicts

- Если `"Nachhilfe"` нужна и как `GROUP`, и как `INDIVIDUAL`, current unique constraint `(year, name)` это запрещает.
- Нужно выбрать и зафиксировать одно решение:
  - либо unique на `(year, name, kind)`;
  - либо разные имена категорий;
  - либо иной доменный компромисс.
- Это решение должно быть согласовано с lookup-логикой, UI и миграциями.

## 14. Regression Scenarios

- Record в `LEGACY` без meaningful change не должен случайно перейти в `AUTO`.
- Forced recalculation из `LEGACY` должен использовать тот же движок, что и обычный `AUTO`.
- Record в `OVERRIDE` не должен неожиданно пересчитываться при unrelated edits.
- Category mismatch warning должен исчезать после корректного apply и approval.
- Один и тот же сценарий должен давать одинаковый результат при:
  - form preview;
  - apply category price;
  - save;
  - approve;
  - reopen detail page.
