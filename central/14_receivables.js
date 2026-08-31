// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Дебиторка — остатки кабинетов маркетплейсов в баланс
// ════════════════════════════════════════════════════════════
//
// Заполняет в блоке «Дебиторская задолженность» управленческого баланса две строки
// на каждый кабинет:
//   «Остаток на балансе кабинета ВБ»   — сколько Wildberries должен кабинету,
//   «Остаток на балансе кабинета ОЗОН» — сколько Ozon должен кабинету.
// Сама «Дебиторская задолженность» и итоги по ИП — формулы, они пересчитаются сами.
//
// ЧЕГО ЗДЕСЬ НЕТ И ПОЧЕМУ. Строки «Ожидаем поступление на РС от ВБ / от ОЗОН» модуль
// НЕ ТРОГАЕТ. Это деньги, которые площадка уже отправила, но которые ещё не дошли до
// расчётного счёта, и ни один публичный метод их не отдаёт (проверено 31.08.2026:
// у WB живёт только `/account/balance`, остальные пути 404; у Ozon `payments` в отчёте
// о взаиморасчётах — это УЖЕ сделанные выплаты за период, а `/v1/finance/payouts` не
// существует). Единственный источник — срез личного кабинета. Пустая ячейка честнее
// выдуманной, поэтому эти строки остаются человеку.
//
// Строки ищутся ПО НАЗВАНИЮ и привязываются к блоку ИП по заголовку выше — номера
// строк не зашиты: в баланс регулярно вставляют новые строки.

/**
 * Остаток кабинета Wildberries.
 * Требует у токена категорию «Финансы»: токен, выпущенный под «Статистику» и
 * «Аналитику», отвечает 403 `scope is not allowed` — это не поломка, а недостающее
 * право, и такой кабинет мы пропускаем, НЕ обнуляя ячейку.
 */
function fetchWbAccountBalance_(cab) {
  const res = fetchWithRetry(WB_BALANCE_URL, {
    method: 'get', headers: { Authorization: cab.token }, muteHttpExceptions: true
  }, 2);
  const body = JSON.parse(res.getContentText() || '{}');
  return { value: Number(body.current) || 0, extra: Number(body.for_withdraw) || 0 };
}

/**
 * Остаток кабинета Ozon — `end_balance_amount` САМОГО СВЕЖЕГО периода отчёта о
 * взаиморасчётах. Именно оно ведёт себя как «Баланс» в кабинете: пока неделя не
 * закрыта, Ozon отдаёт отдельный период «сегодня…сегодня» с текущим остатком.
 * Спрашивать прошедшую дату бесполезно — на середину недели возвращается период
 * целиком, вместе с ещё не наступившими днями.
 */
function fetchOzonAccountBalance_(cab) {
  const to   = new Date();
  const from = new Date(to.getTime() - OZON_CASHFLOW_LOOKBACK_DAYS * 86400000);
  const res  = fetchWithRetry(OZON_URL_CASH_FLOW, ozonOptions_(cab, {
    date: { from: from.toISOString(), to: to.toISOString() },
    page: 1, page_size: 50, with_details: true
  }));
  const details = ((JSON.parse(res.getContentText() || '{}').result) || {}).details || [];
  if (details.length === 0) throw new Error('отчёт о взаиморасчётах пуст');

  let last = details[0];
  details.forEach(d => {
    if (String(d.period.end) > String(last.period.end)) last = d;
  });
  return { value: Number(last.end_balance_amount) || 0,
           period: String(last.period.begin).slice(0, 10) + '…' + String(last.period.end).slice(0, 10) };
}

/** Собирает остатки по всем кабинетам панели. Пропуск — с причиной словами. */
function collectCabinetBalances_() {
  const out = [];
  loadCabinets().forEach(cab => {
    const row = { mp: cab.mp, ip: cab.id };
    if (!cab.active) { row.skip = 'кабинет выключен в панели'; out.push(row); return; }
    try {
      if (cab.mp === 'WB') {
        if (!cab.token) { row.skip = 'токен не заполнен'; out.push(row); return; }
        const r = fetchWbAccountBalance_(cab);
        row.value = r.value;
        row.note  = 'можно вывести ' + formatNumber(r.extra) + ' ₽';
      } else if (cab.mp === 'OZON') {
        if (!cab.clientId || !cab.apiKey) { row.skip = 'ключи не заполнены'; out.push(row); return; }
        const r = fetchOzonAccountBalance_(cab);
        row.value = r.value;
        row.note  = r.period;
      } else {
        row.skip = 'неизвестный маркетплейс';
      }
    } catch (e) {
      row.skip = receivablesReason_(cab.mp, e.message);
    }
    out.push(row);
  });
  return out;
}

/** Текст отказа маркетплейса — словами, а не кодом. */
function receivablesReason_(mp, message) {
  const text = String(message || '');
  if (text.indexOf('403') >= 0) {
    return mp === 'WB'
      ? 'у токена нет категории «Финансы» — перевыпустить в ЛК'
      : 'у ключа нет прав на финансы — перевыпустить в кабинете';
  }
  if (text.indexOf('401') >= 0) return 'токен отозван или истёк';
  if (text.indexOf('429') >= 0) return 'кабинет под лимитером — повторить позже';
  return text.slice(0, 120);
}

/** Лист баланса: берём первый, где есть колонка с нужной датой. */
function findBalanceSheet_(ss, wantDate) {
  for (let i = 0; i < BALANCE_SHEETS.length; i++) {
    const sheet = ss.getSheetByName(BALANCE_SHEETS[i]);
    if (!sheet) continue;
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    for (let c = 0; c < header.length; c++) {
      if (String(header[c]).trim() === wantDate) return { sheet: sheet, col: c + 1 };
    }
  }
  return null;
}

/**
 * Строки «Остаток на балансе кабинета …» с привязкой к блоку ИП.
 * Хвост заголовка блока («ИП11 50%») возвращается отдельно: он означает долю
 * холдинга в проекте, а строка называет ПОЛНЫЙ остаток кабинета. Путать нельзя,
 * и решение о доле принимает владелец, а не скрипт.
 */
function findReceivableRows_(sheet) {
  const last = sheet.getLastRow();
  const col  = sheet.getRange(1, 1, last, 1).getDisplayValues();
  const out  = [];
  let block = '', share = '';
  for (let i = 0; i < col.length; i++) {
    const a = String(col[i][0]).trim();
    if (!a) continue;
    const m = a.match(/^ИП\s*(\d+)\s*(.*)$/);
    if (m) { block = 'ИП' + m[1]; share = m[2].trim() ? a : ''; continue; }
    if (!block) continue;
    if (a === ROW_BALANCE_WB)  out.push({ row: i + 1, ip: block, mp: 'WB',   share: share });
    if (a === ROW_BALANCE_OZON) out.push({ row: i + 1, ip: block, mp: 'OZON', share: share });
  }
  return out;
}

/** План записи: что и куда встанет, что пропущено и почему. */
function planCabinetBalances_(wantDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = findBalanceSheet_(ss, wantDate);
  if (!target) {
    throw new Error('Ни в одном листе баланса нет колонки ' + wantDate +
      '. Листы, где искали: ' + BALANCE_SHEETS.join(', '));
  }
  const got = {};
  collectCabinetBalances_().forEach(r => { got[r.mp + '|' + r.ip] = r; });

  const rows = findReceivableRows_(target.sheet);
  const write = [], skip = [];
  rows.forEach(r => {
    const info = got[r.mp + '|' + r.ip];
    const was = Number(target.sheet.getRange(r.row, target.col).getValue()) || 0;
    if (!info) { skip.push({ row: r.row, ip: r.ip, mp: r.mp, why: 'кабинета нет в «Панели управления»' }); return; }
    if (info.skip) { skip.push({ row: r.row, ip: r.ip, mp: r.mp, why: info.skip }); return; }
    write.push({ row: r.row, ip: r.ip, mp: r.mp, was: was, value: info.value,
                 note: info.note || '', share: r.share });
  });
  return { sheet: target.sheet, col: target.col, date: wantDate, write: write, skip: skip };
}

/** Дата по умолчанию — сегодня: остатки кабинета API отдаёт только «на сейчас». */
function receivablesToday_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
}

/**
 * Пункт меню: собрать остатки кабинетов и записать в дебиторку баланса.
 * Сухой прогон и подтверждение С ЦИФРОЙ — до записи, а не после.
 */
function syncCabinetBalances() {
  const ui = SpreadsheetApp.getUi();
  const ask = ui.prompt('💰 Остатки кабинетов в дебиторку',
    'Дата колонки баланса (ДД.ММ.ГГГГ).\n\nОстатки кабинетов площадки отдают только ' +
    '«на сейчас» — истории по балансу нет ни у WB, ни у Ozon, поэтому осмысленна ' +
    'только сегодняшняя дата.\n\nПо умолчанию: ' + receivablesToday_(),
    ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;

  const wantDate = String(ask.getResponseText() || '').trim() || receivablesToday_();
  const plan = planCabinetBalances_(wantDate);

  if (plan.write.length === 0) {
    ui.alert('💰 Остатки кабинетов',
      'Записывать нечего.\n\n' + skipText_(plan.skip), ui.ButtonSet.OK);
    return;
  }

  let sum = 0;
  plan.write.forEach(w => { sum += w.value; });
  const preview = plan.write.map(w =>
    '  стр ' + w.row + '  ' + w.ip + ' ' + w.mp + ':  ' +
    formatNumber(w.was) + ' → ' + formatNumber(w.value) +
    (w.share ? '   ⚠ блок «' + w.share + '» — пишем ПОЛНЫЙ остаток кабинета' : '')
  ).join('\n');

  const answer = ui.alert('Записать ' + plan.write.length + ' значений на сумму ' +
    formatNumber(sum) + ' ₽?',
    'Лист «' + plan.sheet.getName() + '», колонка ' + colLetter(plan.col) + ' (' + wantDate + ')\n\n' +
    preview + '\n\n' + skipText_(plan.skip), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  plan.write.forEach(w => { plan.sheet.getRange(w.row, plan.col).setValue(w.value); });
  plan.write.forEach(w => writeLog(w.mp, w.ip, 'Успех',
    'дебиторка ' + wantDate + ': ' + formatNumber(w.value) + ' ₽'));
  plan.skip.forEach(s => writeLog(s.mp, s.ip, 'Пропуск', 'дебиторка ' + wantDate + ': ' + s.why));

  ui.alert('💰 Готово',
    'Записано ' + plan.write.length + ' значений на ' + formatNumber(sum) + ' ₽.\n\n' +
    'Строки «Ожидаем поступление на РС» не заполнялись: публичный API площадок ' +
    'статуса выплаты не отдаёт, их ставит человек.\n\n' + skipText_(plan.skip),
    ui.ButtonSet.OK);
}

function skipText_(skip) {
  if (!skip.length) return 'Пропущенных кабинетов нет.';
  return 'Не тронуто (ячейка остаётся как была):\n' +
    skip.map(s => '  стр ' + s.row + '  ' + s.ip + ' ' + s.mp + ' — ' + s.why).join('\n');
}

/** Окно «что встанет», без записи. Рисует лоадер — здесь только {html, text}. */
function previewCabinetBalances(version) {
  const wantDate = receivablesToday_();
  let plan;
  try {
    plan = planCabinetBalances_(wantDate);
  } catch (e) {
    const msg = String(e.message || e);
    return { html: pultPage_('💰 Остатки кабинетов', '<div class="warn">' + pultEsc_(msg) + '</div>', version),
             text: msg };
  }

  let sum = 0;
  plan.write.forEach(w => { sum += w.value; });

  const html = pultPage_('💰 Остатки кабинетов в дебиторку', [
    '<div class="act">Лист «' + pultEsc_(plan.sheet.getName()) + '», колонка ' +
    colLetter(plan.col) + ' (' + pultEsc_(wantDate) + '). Готово к записи: <b>' +
    plan.write.length + '</b> значений на <b>' + formatNumber(sum) + ' ₽</b>. ' +
    'Записывает пункт «4️⃣ 💰 Остатки кабинетов в дебиторку».</div>',
    '<table><tr><th>стр</th><th>ИП</th><th>МП</th><th>было</th><th>станет</th><th>источник</th></tr>',
    plan.write.map(w => '<tr><td>' + w.row + '</td><td>' + pultEsc_(w.ip) + '</td><td>' +
      w.mp + '</td><td>' + formatNumber(w.was) + '</td><td><b>' + formatNumber(w.value) +
      '</b></td><td>' + pultEsc_(w.note) + (w.share ? ' ⚠ блок «' + pultEsc_(w.share) +
      '» — пишем полный остаток' : '') + '</td></tr>').join(''),
    '</table>',
    plan.skip.length ? '<h2>Не тронуто</h2><table><tr><th>стр</th><th>ИП</th><th>МП</th><th>почему</th></tr>' +
      plan.skip.map(s => '<tr><td>' + s.row + '</td><td>' + pultEsc_(s.ip) + '</td><td>' +
        s.mp + '</td><td>' + pultEsc_(s.why) + '</td></tr>').join('') + '</table>' : '',
    '<div class="warn"><b>«Ожидаем поступление на РС» скрипт не заполняет.</b> Это деньги, ' +
    'которые площадка уже отправила, но которые ещё не дошли до счёта, и публичного метода ' +
    'для них нет ни у WB, ни у Ozon — единственный источник статуса выплаты это личный ' +
    'кабинет. Пустая ячейка честнее выдуманной.</div>'
  ].join(''), version);

  const text = ['ОСТАТКИ КАБИНЕТОВ В ДЕБИТОРКУ', wantDate, '',
    plan.write.map(w => `  стр ${w.row} ${w.ip} ${w.mp}: ${formatNumber(w.was)} → ${formatNumber(w.value)}`).join('\n'),
    '', 'Итого: ' + formatNumber(sum) + ' ₽', '', skipText_(plan.skip)].join('\n');

  return { html: html, text: text };
}
