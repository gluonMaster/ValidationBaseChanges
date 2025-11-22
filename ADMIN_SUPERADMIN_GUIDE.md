## Инструкция для администратора (Admin)

- **Источник истины:** поле `ID` в базе Access (`tblKartei.ID`) должно совпадать с `Kartei!AV`. Все сопоставления делаются по этому полю.
- **Открытие файла:** при `Workbook_Open` загружаются данные из `tblKartei`, затем накладываются pending/declined из `pre_tblKartei`/`decl_tblKartei`. После наложения формируется `Kartei_Original`.
- **Статусы на листе `Kartei`:**  
  - `BA="PENDING"` (колонка 53) + голубая заливка `A` — запись из `pre_tblKartei`, ждёт решения Суперадмина.  
  - `BA="DECLINED"` + алая заливка `A` — запись из `decl_tblKartei`, нужна ручная работа через инструмент DeclinedOverview.  
  - Если один ID есть и в pre_, и в decl_ — приоритет decl_, появляется предупреждение.
- **Роли:** `J1` определяет роль. Operator ограничен, Admin — без ограничений Operator’а. Operator не может менять прошедшие месяцы и строки AU="SEPA". Все попытки SEPA‑правок откатываются при синхронизации с сообщением.
- **Редактирование и история:** любые изменения в месяцах U–AF и полях F/J/O пишутся в историю AZ. Для прошлых месяцев у Admin добавляется префикс `Ruck:`; Notitzen обязательно для изменений. При отмене окна Notitzen обновление строки и запись в базу отменяются.
- **Синхронизация (`CompareAndSyncKartei`):**  
  - Перед запуском убедитесь, что BA не содержит PENDING/DECLINED для строк, которые хотите записать прямо в базу.  
  - Изменения делятся на **safe** (идут в `tblKartei`) и **risky** (идут только в `pre_tblKartei`). Новые ID всегда считаются safe.  
  - После синхронизации выводится сводка по количеству safe/risky.
- **Работа с отклонёнными:**  
  - Запустите `ShowDeclinedOverview` для построения листа `DeclinedOverview`.  
  - Внесите корректировки (на `DeclinedOverview` или напрямую на `Kartei`) и выполните `ApplyDeclinedFixes` — выбранные ID уйдут из `decl_tblKartei` в `pre_tblKartei` и получат статус PENDING.

## Инструкция для суперадмина (Superadmin)

- **Источник данных:** pending‑записи из `pre_tblKartei` (ID = AV) импортируются на лист `Kartei` файла Суперадмина.
- **Загрузка pending:** выполните `valid_ImportPending.LoadPendingChangesFromPre` (или `valid_ApproveFlow.LoadPendingChanges`) — подтянутся все pending по ID, затем построится лист `GrossGeschichte` для анализа истории.
- **Просмотр истории:**  
  - `GrossGeschichte` показывает события по каждому ID (месяцы, Address/Subject1/Subject2, Decl_n).  
  - `Geschichte` можно вызывать на активной строке для детального просмотра отдельной записи.
- **Принятие решения:**  
  - На `GrossGeschichte` пометьте для нужных строк `Decision` = Approved или Declined (в заранее подготовленной колонке, обычно U).  
  - Для Declined система запросит комментарий; он будет добавлен в AZ как `Decl_n: Was(); Is(...комментарий...)`.
- **Применение решений:**  
  - Запустите `valid_ApproveFlow.SyncDecisions`.  
  - Approved: запись (по ID) переносится из `pre_tblKartei` в `tblKartei` и удаляется из `pre_`.  
  - Declined: запись (по ID) переносится из `pre_tblKartei` в `decl_tblKartei` с новым `Decl_n` и удаляется из `pre_`.
- **Инвариант ID:** после любых действий проверяйте, что AV на листе совпадает с ID в `tblKartei/pre_tblKartei/decl_tblKartei`; все операции сопоставляют строки по этому полю.
