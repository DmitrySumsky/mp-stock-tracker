// ════════════════════════════════════════════════════════════
// МЕНЮ — точка входа при открытии таблицы
// ════════════════════════════════════════════════════════════

/**
 * Автоматически вызывается Google Sheets при открытии таблицы.
 * Инициализирует Панель управления и создаёт меню «Автоматизация».
 */
function onOpen() {
  initPanel();

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Автоматизация')

    // ── Модуль 1: Остатки маркетплейсов ──────────────────
    .addSubMenu(ui.createMenu('📦 Остатки маркеты')
      .addItem('Обновить все кабинеты',       'runAllSupplyFunctions')
      .addItem('Выбрать кабинеты...',          'runSelectedSupplyFunctions')
      .addItem('Один кабинет...',              'runSingleFromMenu')
      .addSeparator()
      .addItem('Только OZON',                  'runAllOzon')
      .addItem('Только WB',                    'runAllWB')
      .addSeparator()
      .addItem('Добавить кабинет',             'addNewCabinet')
      .addItem('Проверить низкие остатки',     'checkLowStock')
      .addItem('Проверка API',                 'healthCheck')
      .addItem('Обновить дашборд',             'updateDashboard')
      .addItem('Экспорт дашборда в CSV',       'exportDashboardCSV')
      .addSeparator()
      .addItem('Настроить автозапуск...',      'setupTriggers')
      .addItem('Удалить все триггеры',         'removeTriggers'))

    // ── Модуль 2: МойСклад ───────────────────────────────
    .addSubMenu(ui.createMenu('🏪 Остатки МС')
      .addItem('Загрузить остатки сейчас',         'fetchMsStock')
      .addSeparator()
      .addItem('Включить ежедневный автозапуск',   'createMsDailyTrigger')
      .addItem('Отключить ежедневный автозапуск',  'removeMsDailyTrigger'))

    // ── Модуль 3: Кредиты ────────────────────────────────
    .addSubMenu(ui.createMenu('💳 Кредиты')
      .addItem('Предпросмотр данных',              'previewCreditsSync')
      .addItem('Записать в последний столбец',     'syncCreditsToBalance')
      .addItem('Записать в конкретный столбец...', 'syncCreditsToSpecificColumn'))

    .addSeparator()

    // ── Сервис ───────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ Сервис')
      .addItem('Очистка старых данных...',     'cleanupOldData')
      .addItem('Миграция из старых листов',    'migrateFromOldConfig'))

    .addToUi();
}
