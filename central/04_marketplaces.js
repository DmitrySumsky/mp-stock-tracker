// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Остатки маркетплейсов — оркестрация
// ════════════════════════════════════════════════════════════

// ─── Публичные точки входа ───────────────────────────────

function runAllSupplyFunctions() {
  const cabinets = getActiveCabinets();
  if (cabinets.length === 0) {
    SpreadsheetApp.getUi().alert('Нет активных кабинетов. Проверьте Панель управления.');
    return;
  }
  const report = ['🚀 *Обновление ВСЕХ остатков маркетплейсов*'];
  report.push(...processCabinets(cabinets));
  report.push('🏁 *Все операции завершены*');
  sendTelegram(report.join('\n'));
  updateDashboard();
}

function runAllOzon() {
  const cabinets = getActiveCabinets('OZON');
  const report   = ['🚀 *Обновление всех OZON*'];
  report.push(...processCabinets(cabinets));
  report.push('🏁 *Готово*');
  sendTelegram(report.join('\n'));
  updateDashboard();
}

function runAllWB() {
  const cabinets = getActiveCabinets('WB');
  const report   = ['🚀 *Обновление всех WB*'];
  report.push(...processCabinets(cabinets));
  report.push('🏁 *Готово*');
  sendTelegram(report.join('\n'));
  updateDashboard();
}

function runSelectedSupplyFunctions() {
  const ui       = SpreadsheetApp.getUi();
  const cabinets = getActiveCabinets();
  if (cabinets.length === 0) { ui.alert('Нет активных кабинетов.'); return; }

  let prompt = 'Введите номера через запятую (0 = все):\n\n';
  cabinets.forEach((c, i) => { prompt += `${i + 1}. ${c.mp} — ${c.id}\n`; });

  const resp = ui.prompt('Выбор кабинетов', prompt, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = resp.getResponseText().trim();
  if (input === '0') { runAllSupplyFunctions(); return; }

  const nums = input.split(',')
    .map(s => parseInt(s.trim(), 10))
    .filter(n => !isNaN(n) && n >= 1 && n <= cabinets.length);

  if (nums.length === 0) { ui.alert('Ничего не выбрано.'); return; }

  const selected = nums.map(n => cabinets[n - 1]);
  const report   = ['🚀 *Выборочное обновление*'];
  report.push(...processCabinets(selected));
  report.push('🏁 *Готово*');
  sendTelegram(report.join('\n'));
  updateDashboard();
}

function runSingleFromMenu() {
  const ui       = SpreadsheetApp.getUi();
  const cabinets = getActiveCabinets();
  if (cabinets.length === 0) { ui.alert('Нет активных кабинетов.'); return; }

  let prompt = 'Введите номер:\n\n';
  cabinets.forEach((c, i) => { prompt += `${i + 1}. ${c.mp} — ${c.id}\n`; });

  const resp = ui.prompt('Один кабинет', prompt, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const num = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(num) || num < 1 || num > cabinets.length) { ui.alert('Некорректный номер.'); return; }

  const cab    = cabinets[num - 1];
  const report = [`🚀 *${cab.mp} ${cab.id}*`];
  report.push(...processCabinets([cab]));
  sendTelegram(report.join('\n'));
  updateDashboard();
}

function addNewCabinet() {
  const ui    = SpreadsheetApp.getUi();
  const panel = initPanel();

  const mpResp = ui.prompt('Новый кабинет (1/2)', 'Маркетплейс (OZON или WB):', ui.ButtonSet.OK_CANCEL);
  if (mpResp.getSelectedButton() !== ui.Button.OK) return;
  const mp = mpResp.getResponseText().trim().toUpperCase();
  if (mp !== 'OZON' && mp !== 'WB') { ui.alert('Укажите OZON или WB'); return; }

  const ipResp = ui.prompt('Новый кабинет (2/2)', 'Название ИП (например: ИП12):', ui.ButtonSet.OK_CANCEL);
  if (ipResp.getSelectedButton() !== ui.Button.OK) return;
  const ipName = ipResp.getResponseText().trim();
  if (!ipName) { ui.alert('Название не может быть пустым'); return; }

  const sheetName = mp === 'OZON'
    ? `Остатки по кластерам ${ipName}`
    : `Остатки ВБ ${ipName}`;

  panel.appendRow([mp, ipName, '', '', '', sheetName, 'Да', '', '', '']);

  ui.alert(
    `✅ ${mp} ${ipName} добавлен!\n\n` +
    `Впишите ${mp === 'OZON' ? 'Client ID и API Key' : 'Token'} ` +
    `прямо на листе «${SHEET_PANEL}».\n` +
    `Лист «${sheetName}» создастся при первом запуске.`
  );
}

// ─── Обработчик кабинетов ────────────────────────────────

function processCabinets(cabinets) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  const C     = PANEL.COLS;
  const results = [];

  cabinets.forEach(cab => {
    const now = new Date();
    try {
      let rowCount = 0;

      if      (cab.mp === 'OZON') rowCount = fetchOzon(ss, cab);
      else if (cab.mp === 'WB')   rowCount = fetchWB(ss, cab);
      else throw new Error(`Неизвестный МП: ${cab.mp}`);

      const msg = `✅ ${cab.mp} ${cab.id}: +${rowCount}`;
      results.push(msg);
      writeLog(cab.mp, cab.id, 'Успех', msg);

      if (panel) {
        panel.getRange(cab.rowIndex, C.LAST_RUN).setValue(now);
        panel.getRange(cab.rowIndex, C.STATUS).setValue('✅');
        panel.getRange(cab.rowIndex, C.ROW_COUNT).setValue(rowCount);
      }
    } catch (e) {
      const errorMsg = `⛔ ${cab.mp} ${cab.id}: ${e.message}`;
      results.push(errorMsg);
      writeLog(cab.mp, cab.id, 'Ошибка', e.message);

      if (panel) {
        panel.getRange(cab.rowIndex, C.LAST_RUN).setValue(now);
        panel.getRange(cab.rowIndex, C.STATUS).setValue('❌');
        panel.getRange(cab.rowIndex, C.ROW_COUNT).setValue(0);
      }
    }
  });

  return results;
}
