// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Кредиты → Управленческий баланс
// ════════════════════════════════════════════════════════════

// ─── Публичные точки входа ───────────────────────────────

/** Показывает предпросмотр без записи данных. */
function previewCreditsSync() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const creditsSheet = ss.getSheetByName(SHEET_CREDITS);
  if (!creditsSheet) {
    SpreadsheetApp.getUi().alert(`Лист «${SHEET_CREDITS}» не найден`);
    return;
  }

  const targetCol  = findLastHeaderCol();
  const creditsData = creditsSheet.getDataRange().getValues();
  const headers    = creditsData[0];

  const colIndex = {};
  for (const [key, headerName] of Object.entries(CR_HEADERS)) {
    const idx = headers.findIndex(h => String(h).trim().toLowerCase() === headerName.toLowerCase());
    if (idx === -1) {
      SpreadsheetApp.getUi().alert(
        `Столбец «${headerName}» не найден!\n\nЗаголовки: ${headers.join(', ')}`
      );
      return;
    }
    colIndex[key] = idx;
  }

  const { debtByCredit, totalRows, paidRows, unpaidRows } =
    collectDebts_(creditsData, colIndex);

  const lines = [
    `ПРЕДПРОСМОТР (данные НЕ записаны)`,
    `Целевой столбец: ${colLetter(targetCol)}`,
    ``,
    `Найдены столбцы:`,
    `  Основной долг = ${colLetter(colIndex.PRINCIPAL + 1)}`,
    `  Кредит = ${colLetter(colIndex.CREDIT_NAME + 1)}`,
    `  Оплачено = ${colLetter(colIndex.PAID + 1)}`,
    ``,
    `Строк: ${totalRows} всего, ${paidRows} оплачено, ${unpaidRows} не оплачено`,
    ``
  ];

  const mappedCredits = new Set();
  const ipTotals = {};
  let grandTotal = 0;

  for (const m of CREDIT_MAP) {
    const debt = debtByCredit[m.credit] || 0;
    lines.push(`${m.credit} → ${formatNumber(debt)}`);
    mappedCredits.add(m.credit);
    ipTotals[m.ip] = (ipTotals[m.ip] || 0) + debt;
    grandTotal += debt;
  }

  lines.push('', '--- ИТОГО ПО ИП ---');
  for (const [ip, total] of Object.entries(ipTotals)) {
    lines.push(`${ip}: ${formatNumber(total)}`);
  }
  lines.push(`ВСЕГО: ${formatNumber(grandTotal)}`);

  const unmapped = Object.keys(debtByCredit).filter(n => !mappedCredits.has(n));
  if (unmapped.length > 0) {
    lines.push('', '--- БЕЗ МАППИНГА (не будут записаны) ---');
    unmapped.forEach(n => lines.push(`${n}: ${formatNumber(debtByCredit[n])}`));
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/** Синхронизирует кредиты в последний заполненный столбец баланса. */
function syncCreditsToBalance() {
  const targetCol = findLastHeaderCol();
  const result    = writeCreditDataToColumn(targetCol);
  sendTelegram(buildReport(result).join('\n'));
  showResultAlert(result);
}

/** Запрашивает букву столбца и записывает туда. */
function syncCreditsToSpecificColumn() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Выбор столбца',
    'Буква столбца для записи (например: J, K, L):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const letter = resp.getResponseText().trim().toUpperCase();
  if (!letter.match(/^[A-Z]{1,2}$/)) { ui.alert('Некорректная буква столбца'); return; }

  const col     = letterToCol(letter);
  const confirm = ui.alert(
    'Подтверждение',
    `Записать кредиты в столбец ${letter}?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const result = writeCreditDataToColumn(col);
  sendTelegram(buildReport(result).join('\n'));
  showResultAlert(result);
}

// ─── Основная логика ─────────────────────────────────────

function writeCreditDataToColumn(targetCol) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const creditsSheet = ss.getSheetByName(SHEET_CREDITS);
  if (!creditsSheet) throw new Error(`Лист «${SHEET_CREDITS}» не найден`);

  const creditsData = creditsSheet.getDataRange().getValues();
  const headers     = creditsData[0];

  const colIndex = {};
  for (const [key, headerName] of Object.entries(CR_HEADERS)) {
    const idx = headers.findIndex(h => String(h).trim().toLowerCase() === headerName.toLowerCase());
    if (idx === -1) throw new Error(
      `Столбец «${headerName}» не найден на листе «${SHEET_CREDITS}». ` +
      `Заголовки: ${headers.join(', ')}`
    );
    colIndex[key] = idx;
  }

  const { debtByCredit, totalRows, paidRows, unpaidRows } =
    collectDebts_(creditsData, colIndex);

  const balanceSheet = ss.getSheetByName(SHEET_BALANCE);
  if (!balanceSheet) throw new Error(`Лист «${SHEET_BALANCE}» не найден`);

  const balanceData = balanceSheet.getDataRange().getValues();
  const labelToRow  = {};
  for (let i = 0; i < balanceData.length; i++) {
    const label = String(balanceData[i][0]).trim();
    if (label) labelToRow[label] = i + 1;
  }

  const written  = [];
  const notFound = [];
  const ipTotals = {};
  let grandTotal = 0;

  for (const mapping of CREDIT_MAP) {
    const debt = debtByCredit[mapping.credit] || 0;
    const row  = labelToRow[mapping.balanceLabel];
    if (!row) {
      if (debt > 0) notFound.push(`${mapping.credit} → «${mapping.balanceLabel}»`);
      continue;
    }
    balanceSheet.getRange(row, targetCol).setValue(debt);
    written.push({ label: mapping.balanceLabel, debt });
    ipTotals[mapping.ip] = (ipTotals[mapping.ip] || 0) + debt;
    grandTotal += debt;
  }

  const mappedCredits = new Set(CREDIT_MAP.map(m => m.credit));
  const unmapped      = Object.keys(debtByCredit).filter(n => !mappedCredits.has(n));

  const diag = {
    headersFound: `Основной долг=${colLetter(colIndex.PRINCIPAL + 1)}, ` +
                  `Кредит=${colLetter(colIndex.CREDIT_NAME + 1)}, ` +
                  `Оплачено=${colLetter(colIndex.PAID + 1)}`,
    totalRows, paidRows, unpaidRows
  };

  return { written, notFound, ipTotals, grandTotal, unmapped, debtByCredit, targetCol, diag };
}

// ─── Вспомогательные функции ─────────────────────────────

function collectDebts_(creditsData, colIndex) {
  const debtByCredit = {};
  let totalRows = 0, paidRows = 0, unpaidRows = 0;

  for (let i = 1; i < creditsData.length; i++) {
    const row        = creditsData[i];
    const creditName = String(row[colIndex.CREDIT_NAME]).trim();
    if (!creditName) continue;

    totalRows++;
    if (isPaidValue(row[colIndex.PAID])) { paidRows++; continue; }
    unpaidRows++;

    const principal = Number(row[colIndex.PRINCIPAL]) || 0;
    if (principal === 0) continue;
    debtByCredit[creditName] = (debtByCredit[creditName] || 0) + principal;
  }

  return { debtByCredit, totalRows, paidRows, unpaidRows };
}

function isPaidValue(value) {
  if (value === true)  return true;
  if (!value && value !== 0) return false;
  const str = String(value).trim().toLowerCase();
  return str === 'да' || str === 'yes' || str === 'true' || str === '1';
}

function buildReport(result) {
  const colName = colLetter(result.targetCol);
  const report  = [`📊 *Кредиты → Баланс (столбец ${colName})*`, ``];

  if (result.diag) {
    report.push(`Столбцы: ${result.diag.headersFound}`);
    report.push(
      `Строк: ${result.diag.totalRows} всего, ` +
      `${result.diag.paidRows} оплачено, ${result.diag.unpaidRows} не оплачено`
    );
    report.push(``);
  }

  for (const [ip, total] of Object.entries(result.ipTotals)) {
    report.push(`  ${ip}: ${formatNumber(total)}`);
  }
  report.push(``, `*Итого кредиты: ${formatNumber(result.grandTotal)}*`);

  if (result.notFound.length > 0) {
    report.push(``, `⚠️ *Не найдены строки:*`);
    result.notFound.forEach(n => report.push(`  ${n}`));
  }
  if (result.unmapped.length > 0) {
    report.push(``, `⚠️ *Без маппинга:*`);
    result.unmapped.forEach(n => report.push(`  ${n}: ${formatNumber(result.debtByCredit[n])}`));
  }
  return report;
}

function showResultAlert(result) {
  const colName = colLetter(result.targetCol);
  SpreadsheetApp.getUi().alert(
    `Готово!\n\n` +
    `Столбец: ${colName}\n` +
    `Записано: ${result.written.length} кредитов\n` +
    `Итого долг: ${formatNumber(result.grandTotal)}` +
    (result.notFound.length > 0 ? `\n\n⚠️ Не найдены: ${result.notFound.length} строк` : '') +
    (result.unmapped.length  > 0 ? `\n\n⚠️ Без маппинга: ${result.unmapped.join(', ')}` : '')
  );
}
