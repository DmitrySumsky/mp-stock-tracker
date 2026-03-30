// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Чтение конфигурации из Панели управления
// ════════════════════════════════════════════════════════════

function loadTelegramConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return { token: '', chatIds: [] };

  const token    = panel.getRange(PANEL.TG_TOKEN_CELL).getValue().toString().trim();
  const chatsRaw = panel.getRange(PANEL.TG_CHATS_CELL).getValue().toString().trim();
  const chatIds  = chatsRaw.split(';').map(s => s.trim()).filter(Boolean);

  return { token, chatIds };
}

function getMsToken_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return '';
  return panel.getRange(PANEL.MS_TOKEN_CELL).getValue().toString().trim();
}

function loadCabinets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return [];

  const lastRow = panel.getLastRow();
  if (lastRow < PANEL.TABLE_START_ROW) return [];

  const numRows = lastRow - PANEL.TABLE_START_ROW + 1;
  const data    = panel.getRange(PANEL.TABLE_START_ROW, 1, numRows, 10).getValues();

  return data
    .map((row, i) => ({
      rowIndex:  PANEL.TABLE_START_ROW + i,
      mp:        String(row[0]).trim().toUpperCase(),
      id:        String(row[1]).trim(),
      clientId:  String(row[2]).trim(),
      apiKey:    String(row[3]).trim(),
      token:     String(row[4]).trim(),
      sheetName: String(row[5]).trim(),
      active:    String(row[6]).trim().toLowerCase() === 'да'
    }))
    .filter(c => c.mp && c.id && c.sheetName);
}

function getActiveCabinets(mp) {
  const all = loadCabinets();
  if (!mp) return all.filter(c => c.active);
  return all.filter(c => c.active && c.mp === mp.toUpperCase());
}
