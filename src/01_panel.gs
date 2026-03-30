// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Панель управления — инициализация и миграция
// ════════════════════════════════════════════════════════════

/**
 * Создаёт лист «Панель управления», если он ещё не существует.
 * Вызывается автоматически из onOpen().
 */
function initPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PANEL);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_PANEL, 0);

  // Блок Telegram
  sheet.getRange('A1').setValue('🔔 TELEGRAM').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A2').setValue('Токен бота');
  sheet.getRange('A3').setValue('Chat ID (через ;)');
  sheet.getRange('B2').setValue('').setNote('Вставьте токен бота от @BotFather');
  sheet.getRange('B3').setValue('').setNote('Несколько чатов через точку с запятой: -100123;-100456');

  // Блок МойСклад
  sheet.getRange('A4').setValue('🏪 МОЙ СКЛАД').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A5').setValue('API Токен');
  sheet.getRange('B5').setValue('').setNote('Токен из МойСклад: Настройки → Доступ по API');

  // Заголовок кабинетов
  sheet.getRange('A7').setValue('📦 КАБИНЕТЫ').setFontWeight('bold').setFontSize(11);

  // Заголовки таблицы кабинетов
  const headers = [
    'Маркетплейс', 'ИП', 'Client ID (OZON)', 'API Key (OZON)',
    'Token (WB)', 'Имя листа', 'Активен', 'Последний запуск', 'Статус', 'Строк'
  ];
  const headerRange = sheet.getRange(8, 1, 1, headers.length);
  headerRange.setValues([headers]).setFontWeight('bold').setBackground('#e8eaf6');
  sheet.setFrozenRows(8);

  // Валидация
  const mpRule     = SpreadsheetApp.newDataValidation().requireValueInList(['OZON', 'WB']).build();
  const activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Нет']).build();
  sheet.getRange('A9:A200').setDataValidation(mpRule);
  sheet.getRange('G9:G200').setDataValidation(activeRule);

  // Визуальное выделение автозаполняемых столбцов
  sheet.getRange('H9:J200').setBackground('#f5f5f5').setFontColor('#666666');

  // Ширина столбцов
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(6, 240);
  sheet.setColumnWidth(7, 75);
  sheet.setColumnWidth(8, 155);
  sheet.setColumnWidth(9, 80);
  sheet.setColumnWidth(10, 70);

  return sheet;
}

/**
 * Переносит кабинеты из старых технических листов в Панель управления.
 *
 * Адаптируйте ozonMap и wbMap под структуру ваших старых листов:
 *   ozonMap: для каждого кабинета укажите id, ячейку Client ID и ячейку API Key
 *   wbMap:   для каждого кабинета укажите id и ячейку с токеном
 */
function migrateFromOldConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = initPanel();
  const ui    = SpreadsheetApp.getUi();

  const existingData = panel.getRange(PANEL.TABLE_START_ROW, 1).getValue();
  if (existingData) {
    const confirm = ui.alert('Миграция', 'В таблице уже есть данные. Перезаписать?', ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;
  }

  const rows = [];

  // --- OZON: замените имя листа и ячейки под свою структуру ---
  const ozonSheet = ss.getSheetByName('Технический Озон');
  if (ozonSheet) {
    const ozonMap = [
      // { id: 'Кабинет 1', cidCell: 'B1', keyCell: 'B2' },
      // { id: 'Кабинет 2', cidCell: 'B3', keyCell: 'B4' },
    ];
    ozonMap.forEach(m => {
      const cid = ozonSheet.getRange(m.cidCell).getValue().toString().trim();
      const key = ozonSheet.getRange(m.keyCell).getValue().toString().trim();
      rows.push(['OZON', m.id, cid, key, '', `Остатки OZON ${m.id}`, 'Да', '', '', '']);
    });
  }

  // --- WB: замените имя листа и ячейки под свою структуру ---
  const wbSheet = ss.getSheetByName('Технический ВБ');
  if (wbSheet) {
    const wbMap = [
      // { id: 'Кабинет 1', tCell: 'B1' },
      // { id: 'Кабинет 2', tCell: 'D1' },
    ];
    wbMap.forEach(m => {
      const token = wbSheet.getRange(m.tCell).getValue().toString().trim();
      rows.push(['WB', m.id, '', '', token, `Остатки WB ${m.id}`, 'Да', '', '', '']);
    });
  }

  if (rows.length > 0) {
    panel.getRange(PANEL.TABLE_START_ROW, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Telegram-настройки перенесите вручную
  // panel.getRange(PANEL.TG_TOKEN_CELL).setValue('<ваш токен>');
  // panel.getRange(PANEL.TG_CHATS_CELL).setValue('<ваш chat id>');

  ui.alert(
    `✅ Миграция завершена: ${rows.length} кабинетов перенесено.\n\n` +
    `Заполните Токен бота и Chat ID в Панели управления (B2, B3).\n` +
    `Старые технические листы можно удалить.`
  );
}
