// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Утилиты — HTTP, запись в листы, форматирование
// ════════════════════════════════════════════════════════════

// ─── HTTP ────────────────────────────────────────────────

/**
 * Выполняет HTTP-запрос с повторными попытками при ошибках.
 * @param {string} url
 * @param {Object} options — параметры UrlFetchApp.fetch
 * @param {number} [maxRetries=3]
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse}
 */
function fetchWithRetry(url, options, maxRetries) {
  maxRetries = maxRetries || 3;
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) return res;
      lastError = new Error(`HTTP ${res.getResponseCode()}`);
    } catch (e) {
      lastError = e;
    }
    if (attempt < maxRetries) Utilities.sleep(2000 * attempt);
  }

  throw new Error(`${lastError.message} (после ${maxRetries} попыток)`);
}

// ─── Работа с листами ────────────────────────────────────

/**
 * Записывает строки данных в лист (дописывает после существующих).
 * Создаёт лист с заголовками, если он не существует.
 */
function writeToSheet(ss, name, rows, headers) {
  if (rows.length === 0) return;

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  } else {
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Добавляет запись в лог-лист.
 */
function writeLog(mp, ip, status, details) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Дата', 'Маркетплейс', 'ИП', 'Статус', 'Детали']);
  }
  sheet.appendRow([new Date(), mp, ip, status, details]);
}

/**
 * Возвращает лист по имени, создаёт если не существует.
 */
function getOrCreateSheet_(name) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/**
 * Удаляет строки старше заданного количества дней из всех листов кабинетов и лога.
 */
function cleanupOldData() {
  const ui = SpreadsheetApp.getUi();

  const resp = ui.prompt(
    '🧹 Очистка данных',
    'За сколько дней оставить данные? (всё старше будет удалено)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const days = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(days) || days < 1) { ui.alert('Укажите число дней больше 0'); return; }

  const confirm = ui.alert(
    'Подтверждение',
    `Удалить все данные старше ${days} дней?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cabinets = loadCabinets();
  const cutoff   = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  let totalDeleted = 0;

  cabinets.forEach(cab => {
    const sheet = ss.getSheetByName(cab.sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const dateCol = cab.mp === 'OZON' ? 9 : 1;
    const data    = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      const cellValue = data[i][dateCol - 1];
      let rowDate;

      if (cellValue instanceof Date) {
        rowDate = cellValue;
      } else if (typeof cellValue === 'string') {
        const parts = cellValue.match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (parts) rowDate = new Date(parts[3], parts[2] - 1, parts[1]);
      }

      if (rowDate && rowDate < cutoff) {
        sheet.deleteRow(i + 1);
        totalDeleted++;
      }
    }
  });

  const logSheet = ss.getSheetByName(SHEET_LOG);
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getDataRange().getValues();
    for (let i = logData.length - 1; i >= 1; i--) {
      if (logData[i][0] instanceof Date && logData[i][0] < cutoff) {
        logSheet.deleteRow(i + 1);
        totalDeleted++;
      }
    }
  }

  ui.alert(`Удалено ${totalDeleted} строк старше ${days} дней.`);
}

// ─── Работа со столбцами Баланса ─────────────────────────

/**
 * Находит последний заполненный столбец в строке заголовков листа «Управленческий баланс».
 */
function findLastHeaderCol() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const balanceSheet = ss.getSheetByName(SHEET_BALANCE);
  if (!balanceSheet) throw new Error(`Лист «${SHEET_BALANCE}» не найден`);

  const lastCol   = balanceSheet.getLastColumn();
  const headerRow = balanceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let lastFilledCol = 1;
  for (let j = 0; j < headerRow.length; j++) {
    if (headerRow[j] !== '' && headerRow[j] !== null && headerRow[j] !== undefined) {
      lastFilledCol = j + 1;
    }
  }
  return lastFilledCol;
}

// ─── Форматирование ──────────────────────────────────────

/** Число → буква(ы) столбца: 1 → A, 27 → AA */
function colLetter(col) {
  let letter = '';
  while (col > 0) {
    col--;
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col    = Math.floor(col / 26);
  }
  return letter;
}

/** Буква(ы) столбца → номер: A → 1, AA → 27 */
function letterToCol(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

/** Форматирует число с пробелами-разделителями тысяч. */
function formatNumber(num) {
  return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}
