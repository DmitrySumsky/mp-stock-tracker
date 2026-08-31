// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Триггеры — управление автозапуском
// ════════════════════════════════════════════════════════════

/**
 * Диалог настройки периодического обновления остатков маркетплейсов.
 * Триггер МойСклад настраивается отдельно из меню «Остатки МС».
 */
function setupTriggers() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '⏰ Настройка автозапуска',
    'Как часто обновлять остатки маркетплейсов (в часах)?\nНапример: 1, 2, 4, 6, 12',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const hours = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(hours) || hours < 1 || hours > 24) {
    ui.alert('Укажите число от 1 до 24');
    return;
  }

  // Удаляем старый триггер маркетплейсов
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runAllSupplyFunctions')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runAllSupplyFunctions')
    .timeBased()
    .everyHours(hours)
    .create();

  ui.alert(
    `✅ Триггер установлен:\n` +
    `• Обновление остатков маркетплейсов — каждые ${hours} ч.\n\n` +
    `Триггер для МойСклад настраивается отдельно через «Остатки МС → Включить ежедневный автозапуск».`
  );
}

/** Удаляет ВСЕ триггеры проекта. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Все триггеры удалены.');
}
