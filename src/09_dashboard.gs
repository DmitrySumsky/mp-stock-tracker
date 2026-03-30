// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Дашборд и экспорт
// ════════════════════════════════════════════════════════════

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DASH);
  if (!sheet) sheet = ss.insertSheet(SHEET_DASH);

  sheet.clear();

  const headers = ['Маркетплейс', 'ИП', 'Лист', 'Строк данных', 'Последнее обновление', 'Статус'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8eaf6');
  sheet.setFrozenRows(1);

  const cabinets = loadCabinets();
  if (cabinets.length === 0) return;

  const rows = cabinets.map(c => {
    const dataSheet = ss.getSheetByName(c.sheetName);
    if (!dataSheet) return [c.mp, c.id, c.sheetName, 0, '—', '⚠️ Нет листа'];

    const lastRow  = dataSheet.getLastRow();
    const rowCount = Math.max(0, lastRow - 1);

    let lastUpdate = '—';
    if (lastRow > 1) {
      const val = dataSheet.getRange(lastRow, 1).getValue();
      if (val instanceof Date) {
        lastUpdate = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
      } else if (typeof val === 'string' && val) {
        lastUpdate = val;
      }
    }

    const status = !c.active ? '⏸ Неактивен' : (rowCount > 0 ? '✅' : '⚠️ Пусто');
    return [c.mp, c.id, c.sheetName, rowCount, lastUpdate, status];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const totalRow = rows.length + 3;
  sheet.getRange(totalRow, 1).setValue('📊 ИТОГО').setFontWeight('bold');
  sheet.getRange(totalRow, 4).setFormula(`=SUM(D2:D${rows.length + 1})`).setFontWeight('bold');
  sheet.getRange(totalRow + 1, 1).setValue('🕐 Обновлено').setFontColor('#999999');
  sheet.getRange(totalRow + 1, 2).setValue(new Date()).setFontColor('#999999');

  sheet.autoResizeColumns(1, headers.length);
}

function exportDashboardCSV() {
  updateDashboard();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DASH);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const csv  = data.map(row => row.map(cell => {
    const val = cell instanceof Date
      ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : String(cell);
    return `"${val.replace(/"/g, '""')}"`;
  }).join(',')).join('\n');

  const filename = 'stock_dashboard_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
  const file = DriveApp.createFile(filename, csv, 'text/csv');

  SpreadsheetApp.getUi().alert(`📁 CSV сохранён на Google Drive:\n${file.getUrl()}`);
}
