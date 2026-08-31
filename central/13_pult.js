// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Пульт — инструкция, статус, проверка связи, миграция листов
// ════════════════════════════════════════════════════════════
//
// Стандарт «Пульт» (57.mp-core/ПУЛЬТ.md): книга — это пульт управления, а не
// «таблица со скриптом». Отсюда четыре обязательных окна:
//
//   📖 Как работать       — инструкция ИЗ КОДА, она не может отстать от версии;
//   📊 Что сейчас происходит — одно окно вместо чтения трёх листов, называет
//                            ОДНО следующее действие человека;
//   🔌 Проверка связи     — все кабинеты ОДНИМ fetchAll и БЕЗ ретраев: код
//                            ответа сам по себе диагноз, а диагностика, умеющая
//                            ждать, упирается в шестиминутный лимит и выглядит
//                            как «кнопка не работает»;
//   ⚙️ Обновить настройки — миграция листов без боевого прогона.
//
// Окна рисует ЛОАДЕР: HtmlService доступен только статическому коду проекта,
// а этот файл приезжает по сети. Поэтому здесь — {html, text}, где text обязан
// быть самодостаточным: если окно не открылось, человек читает его в alert.

/** Ссылка на нужный лист, чтобы окна не падали на чужой книге. */
function pultSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function pultEsc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Общая обёртка окна: одна типографика на все диалоги книги. */
function pultPage_(title, bodyHtml, version) {
  return '<!doctype html><meta charset="utf-8">' +
    '<style>' +
    'body{font:14px/1.55 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#202124;margin:0;padding:18px 22px}' +
    'h1{font-size:18px;margin:0 0 4px}h2{font-size:15px;margin:18px 0 6px}' +
    '.v{color:#5f6368;font-size:12px;margin-bottom:14px}' +
    'ol,ul{margin:6px 0 6px 20px;padding:0}li{margin:3px 0}' +
    'table{border-collapse:collapse;margin:6px 0;width:100%}' +
    'th,td{border:1px solid #dadce0;padding:5px 8px;text-align:left;font-size:13px}' +
    'th{background:#f1f3f4;font-weight:600}' +
    '.warn{background:#fce8e6;border-left:4px solid #d93025;padding:9px 12px;margin:12px 0;border-radius:3px}' +
    '.ok{background:#e6f4ea;border-left:4px solid #1e8e3e;padding:9px 12px;margin:12px 0;border-radius:3px}' +
    '.act{background:#e8f0fe;border-left:4px solid #1a73e8;padding:9px 12px;margin:12px 0;border-radius:3px}' +
    'code{background:#f1f3f4;padding:1px 4px;border-radius:3px;font-size:12px}' +
    '</style>' +
    '<h1>' + pultEsc_(title) + '</h1>' +
    '<div class="v">' + pultEsc_(version || '') + '</div>' + bodyHtml;
}

// ─── 📖 Как работать ─────────────────────────────────────

/**
 * Инструкция живёт в коде и обновляется тем же заходом, что и функциональность
 * (ПУЛЬТ §4). Файл на Диске через две версии врёт, а эта — не может.
 */
function panelHelp(version) {
  const steps = [
    ['1️⃣ Обновить остатки — все кабинеты',
     'Обходит все кабинеты со значением «Да» в колонке «Активен» листа «Панель ' +
     'управления»: Ozon и Wildberries. Каждый кабинет пишет свой лист, дата ' +
     'прогона — в колонке A. Повторный запуск в тот же день переписывает только ' +
     'сегодняшний блок, прошлые дни не трогает.'],
    ['2️⃣ Обновить остатки МойСклад',
     'Отдельная кнопка: МойСклад отвечает медленнее маркетплейсов, и держать его ' +
     'в общем прогоне значит рисковать шестиминутным лимитом.'],
    ['3️⃣ Кредиты в баланс',
     'Переносит остаток основного долга из листа «Кредиты Import» в строки ' +
     'управленческого баланса. Работает только если лист заполнен и заполнено ' +
     'соответствие названий кредитов строкам баланса.']
  ];

  const cols = [
    ['Доступно к продаже', 'free_to_sell из stock_on_warehouses. ВНИМАНИЕ: уже включает товар, на который создана заявка на вывоз со склада.'],
    ['Доступно (аналитика)', 'available из analytics/stocks — «чистое» доступное, без вывозимого.'],
    ['Возврат продавцу', 'весь блок кабинета «Вывоз со склада Ozon»: готовим к вывозу + готово + возвращается.'],
    ['Доставляем покупателям', 'товар уехал к покупателю, но ещё не выкуплен. Считается по отправлениям FBO, ни одна ручка остатков его не отдаёт.'],
    ['Всего у OZON', 'старая колонка: доступно + готовим + резерв. Оставлена ради формул соседних листов.'],
    ['Всего вверено OZON', 'ИТОГ, который идёт в баланс: всё, что физически у маркетплейса. «Заявлено к поставке» сюда не входит — этот товар ещё на нашем складе.']
  ];

  const html = pultPage_('📖 Как работать', [
    '<p>Книга собирает остатки Ozon, Wildberries и МойСклада по всем кабинетам ' +
    'и складывает их историей по дням. Из последнего дня берутся запасы для ' +
    'управленческого баланса.</p>',
    '<div class="warn"><b>Единственное правило, которое нельзя нарушать:</b> ключи ' +
    'кабинетов живут только на листе «Панель управления». Не переносить их в код, ' +
    'не выкладывать в переписку, не хранить копией в других книгах.</div>',
    '<h2>Порядок работы</h2><ol>',
    steps.map(s => '<li><b>' + pultEsc_(s[0]) + '</b><br>' + pultEsc_(s[1]) + '</li>').join(''),
    '</ol>',
    '<h2>Что заполняет человек</h2><ul>' +
    '<li>лист «Панель управления»: строки кабинетов — ключи, имя листа, «Активен»;</li>' +
    '<li>колонка «Себестоимость, ₽» на листах остатков Ozon — её ставит отдельный ' +
    'инструмент, скрипт книги её не трогает;</li>' +
    '<li>лист «Кредиты Import» — выгрузка графика платежей.</li></ul>',
    '<h2>Колонки листа остатков Ozon, которые чаще всего понимают неправильно</h2>',
    '<table><tr><th>Колонка</th><th>Что это на самом деле</th></tr>',
    cols.map(c => '<tr><td>' + pultEsc_(c[0]) + '</td><td>' + pultEsc_(c[1]) + '</td></tr>').join(''),
    '</table>',
    '<h2>Если что-то пошло не так</h2><ol>' +
    '<li>Откройте «📊 Что сейчас происходит» — окно называет одно действие.</li>' +
    '<li>Ключ не принят — «🔌 Проверка связи» покажет, какой именно кабинет и с ' +
    'каким кодом отказал. 403 у Ozon — ключ истёк, перевыпустить в кабинете.</li>' +
    '<li>Лист не пополнился — посмотрите лист «Лог», последние строки: там причина ' +
    'словами, а не «HTTP 404».</li>' +
    '<li>После обновления кода — «⚙️ Обновить настройки таблицы»: она досоздаёт ' +
    'новые колонки без боевого прогона.</li></ol>'
  ].join(''), version);

  const text = [
    'КАК РАБОТАТЬ',
    version || '',
    '',
    '1) Обновить остатки — все кабинеты: обходит активные кабинеты Ozon и WB,',
    '   пишет по листу на кабинет, дата прогона в колонке A. Повтор за тот же',
    '   день переписывает только сегодняшний блок.',
    '2) Обновить остатки МойСклад — отдельной кнопкой (он медленнее).',
    '3) Кредиты в баланс — из листа «Кредиты Import» в строки баланса.',
    '',
    'ПРАВИЛО: ключи кабинетов живут только на листе «Панель управления».',
    '',
    'Колонки Ozon, которые чаще всего понимают неправильно:',
    ' • «Доступно к продаже» уже включает товар «готовим к вывозу»;',
    ' • «Возврат продавцу» — это весь блок «Вывоз со склада Ozon»;',
    ' • «Доставляем покупателям» — уехало к покупателю, ещё не выкуплено;',
    ' • «Всего вверено OZON» — итог, который идёт в баланс.',
    '',
    'Если что-то пошло не так: «Что сейчас происходит» → «Проверка связи» → лист «Лог».'
  ].join('\n');

  return { html: html, text: text };
}

// ─── 📊 Что сейчас происходит ────────────────────────────

/**
 * Одно окно вместо чтения статусов в трёх листах (ПУЛЬТ §2). Заканчивается
 * ОДНИМ действием: либо задача человеку, либо «делать ничего не нужно».
 */
function panelStatus(version) {
  const tz    = Session.getScriptTimeZone();
  const panel = pultSheet_(SHEET_PANEL);
  const rows  = [];
  let   stale = [], broken = [], noKey = [];

  loadCabinets().forEach(cab => {
    let lastRun = '', status = '', count = '';
    if (panel) {
      const v = panel.getRange(cab.rowIndex, PANEL.COLS.LAST_RUN, 1, 3).getValues()[0];
      lastRun = v[0] instanceof Date ? Utilities.formatDate(v[0], tz, 'dd.MM.yyyy HH:mm') : String(v[0] || '');
      status  = String(v[1] || '');
      count   = String(v[2] || '');
    }
    const sheet = pultSheet_(cab.sheetName);
    const day   = sheet ? panelLastDay_(sheet, tz) : '';
    const keyOk = cab.mp === 'OZON' ? !!(cab.clientId && cab.apiKey) : !!cab.token;

    if (cab.active && !keyOk)          noKey.push(cab.mp + ' ' + cab.id);
    else if (cab.active && status === '❌') broken.push(cab.mp + ' ' + cab.id);
    else if (cab.active && day && day !== Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy')) {
      stale.push(cab.mp + ' ' + cab.id + ' (' + day + ')');
    }

    rows.push([cab.mp, cab.id, cab.active ? 'да' : 'нет',
               keyOk ? 'есть' : '—', lastRun, status, count, day || '—']);
  });

  let action;
  if (noKey.length)       action = 'Впишите ключи на листе «Панель управления»: ' + noKey.join(', ') + '.';
  else if (broken.length) action = 'Разберите отказ: ' + broken.join(', ') + '. Начните с «🔌 Проверка связи».';
  else if (stale.length)  action = 'Данные не за сегодня: ' + stale.join(', ') + '. Нажмите «1️⃣ Обновить остатки».';
  else                    action = 'Ничего делать не нужно — все активные кабинеты собраны сегодня.';

  const log = pultSheet_(SHEET_LOG);
  const tail = [];
  if (log && log.getLastRow() > 1) {
    const n = Math.min(6, log.getLastRow() - 1);
    log.getRange(log.getLastRow() - n + 1, 1, n, 5).getValues().forEach(r => {
      const d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd.MM HH:mm') : String(r[0] || '');
      tail.push(d + '  ' + r[1] + ' ' + r[2] + '  ' + r[4]);
    });
  }

  const html = pultPage_('📊 Что сейчас происходит', [
    '<div class="' + (noKey.length || broken.length ? 'warn' : (stale.length ? 'act' : 'ok')) + '">' +
    '<b>Что сейчас нужно от вас:</b><br>' + pultEsc_(action) + '</div>',
    '<table><tr><th>МП</th><th>ИП</th><th>Активен</th><th>Ключ</th>' +
    '<th>Последний запуск</th><th>Итог</th><th>Строк</th><th>Свежий день в листе</th></tr>',
    rows.map(r => '<tr>' + r.map(c => '<td>' + pultEsc_(c) + '</td>').join('') + '</tr>').join(''),
    '</table>',
    tail.length ? '<h2>Последние строки «Лога»</h2><ul>' +
      tail.map(t => '<li><code>' + pultEsc_(t) + '</code></li>').join('') + '</ul>' : ''
  ].join(''), version);

  const text = ['ЧТО СЕЙЧАС ПРОИСХОДИТ', version || '', '',
    'Нужно от вас: ' + action, '',
    rows.map(r => r.join('  |  ')).join('\n'),
    tail.length ? '\nЛог:\n' + tail.join('\n') : ''].join('\n');

  return { html: html, text: text };
}

/** Самый свежий день в листе остатков — по колонке A, снизу вверх. */
function panelLastDay_(sheet, tz) {
  const last = sheet.getLastRow();
  if (last < 2) return '';
  const from = Math.max(2, last - 200);
  const col  = sheet.getRange(from, 1, last - from + 1, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    const v = col[i][0];
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd.MM.yyyy');
    const m = String(v || '').match(/^(\d{2}\.\d{2}\.\d{4})/);
    if (m) return m[1];
  }
  return '';
}

// ─── 🔌 Проверка связи ───────────────────────────────────

/**
 * Все кабинеты ОДНИМ `fetchAll` и БЕЗ единого ретрая (ПУЛЬТ §2).
 *
 * Клиент сбора честно отрабатывает бэкофф, и дюжина кабинетов × паузы перекрывает
 * шестиминутный лимит Apps Script — пункт меню умирает молча и выглядит как
 * «кнопка не работает». Здесь код ответа сам становится диагнозом.
 *
 * Модуль только ЧИТАЕТ у маркетплейсов, поэтому право записи не проверяется:
 * писать наружу книге нечем и незачем.
 */
function panelCheckConnection() {
  const cabs = loadCabinets().filter(c => c.active);
  const reqs = [], meta = [];

  cabs.forEach(cab => {
    if (cab.mp === 'OZON') {
      if (!cab.clientId || !cab.apiKey) { meta.push({ cab: cab, skip: 'ключи пустые' }); return; }
      const o = ozonOptions_(cab, { limit: 1, offset: 0, warehouse_type: 'ALL' });
      o.url = OZON_URL_STOCK_ON_WAREHOUSES;
      reqs.push(o); meta.push({ cab: cab, i: reqs.length - 1 });
    } else if (cab.mp === 'WB') {
      if (!cab.token) { meta.push({ cab: cab, skip: 'токен пустой' }); return; }
      reqs.push({ url: WB_REMAINS_URL + WB_REMAINS_PARAMS, method: 'get',
                  headers: { Authorization: cab.token }, muteHttpExceptions: true });
      meta.push({ cab: cab, i: reqs.length - 1 });
    } else {
      meta.push({ cab: cab, skip: 'неизвестный маркетплейс' });
    }
  });

  const resp = reqs.length ? UrlFetchApp.fetchAll(reqs) : [];
  const lines = meta.map(m => {
    const who = m.cab.mp + ' ' + m.cab.id;
    if (m.skip) return '⚠️ ' + who + ': ' + m.skip;
    const r = resp[m.i];
    const code = r.getResponseCode();
    if (code >= 200 && code < 300) return '✅ ' + who + ': связь есть';
    return '❌ ' + who + ': ' + code + ' — ' + panelDiagnose_(m.cab.mp, code, r);
  });

  const bad = lines.filter(l => l.charAt(0) !== '✅').length;
  const msg = lines.join('\n') + '\n\n' +
    (bad ? 'Отказов: ' + bad + '. Ключ перевыпускается в личном кабинете маркетплейса; '
         + 'у Ozon нужны права на аналитику и FBO, у WB — раздел «Аналитика».'
         : 'Все активные кабинеты отвечают.');
  SpreadsheetApp.getUi().alert('🔌 Проверка связи', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return lines;
}

/** Код ответа сам по себе диагноз — переводим его на человеческий. */
function panelDiagnose_(mp, code, res) {
  if (code === 401) return 'токен не принят: отозван или истёк';
  if (code === 403) {
    return mp === 'OZON'
      ? 'ключ истёк или у него нет нужной категории прав — перевыпустить'
      : 'у токена нет раздела «Аналитика» — перевыпустить';
  }
  if (code === 429) return 'кабинет под лимитером — это не поломка ключа, повторить позже';
  if (code >= 500)  return 'маркетплейс отвечает ошибкой на своей стороне';
  return extractApiError_(res);
}

// ─── ⚙️ Обновить настройки таблицы ───────────────────────

/**
 * Миграция служебных листов без боевого прогона (ПУЛЬТ §2).
 *
 * После обновления кода человек не должен «прогонять сбор, чтобы появились
 * колонки». Здесь — вставка новых колонок листов остатков Ozon: содержимое
 * старых дней сдвигается вместе с шапкой, ничего не теряется. Прошлые дни
 * остаются с прежним «Всего вверено OZON» — это то, что было собрано тогда,
 * задним числом мы историю не переписываем.
 */
function upgradeSheets() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const done = [], skip = [];

  loadCabinets().filter(c => c.mp === 'OZON').forEach(cab => {
    const sheet = ss.getSheetByName(cab.sheetName);
    if (!sheet || sheet.getLastRow() === 0) { skip.push(cab.id + ': листа ещё нет'); return; }

    const width = Math.max(sheet.getLastColumn(), OZON_HEADERS.length);
    const head  = sheet.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim());

    if (head.indexOf('Доставляем покупателям') >= 0) { skip.push(cab.id + ': уже обновлён'); return; }

    const at = head.indexOf('Кластер');
    if (at < 0) { skip.push(cab.id + ': чужая раскладка, тронуть нельзя'); return; }

    sheet.insertColumnBefore(at + 1);
    sheet.getRange(1, at + 1).setValue('Доставляем покупателям').setFontWeight('bold');
    done.push(cab.id);
  });

  const msg = (done.length ? 'Обновлены листы: ' + done.join(', ') + '\n' : '') +
              (skip.length ? 'Пропущены:\n  ' + skip.join('\n  ') : '') +
              (!done.length && !skip.length ? 'Кабинетов Ozon в панели нет.' : '');
  SpreadsheetApp.getUi().alert('⚙️ Обновление настроек таблицы', msg,
    SpreadsheetApp.getUi().ButtonSet.OK);
  return { done: done, skip: skip };
}
