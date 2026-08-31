/**
 * Тесты центрального кода в Node на заглушках (ПУЛЬТ §11).
 *
 * Скрипт книги не запустить ни `clasp run`, ни из-под сервисного аккаунта: первый
 * запуск требует авторизации живого пользователя. Поэтому подменяем
 * SpreadsheetApp / UrlFetchApp / Utilities и грузим собранный central.js в vm.
 *
 * Ключевая деталь: FakeSheet ПРОВЕРЯЕТ размерность setValues. Именно на
 * несовпадении числа колонок Apps Script падает в бою, а стаб без проверки
 * такую ошибку пропустит.
 *
 * Запросы наружу копятся в массив и сверяются как полезная нагрузка: проверяется,
 * ЧТО ушло бы в Ozon, а не «код не упал».
 *
 *   node tests/run.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CENTRAL = path.join(__dirname, '..', 'central', 'build', 'central.js');
if (!fs.existsSync(CENTRAL)) {
  console.error('Нет central/build/central.js — сначала node tools/build.js');
  process.exit(1);
}
const CODE = fs.readFileSync(CENTRAL, 'utf8');

let passed = 0, failed = 0;
const fails = [];

function ok(name, cond, detail) {
  if (cond) { passed++; return; }
  failed++;
  fails.push(name + (detail ? ' — ' + detail : ''));
}
function eq(name, got, want) {
  ok(name, JSON.stringify(got) === JSON.stringify(want),
     'получили ' + JSON.stringify(got) + ', ждали ' + JSON.stringify(want));
}

// ─── заглушки среды ─────────────────────────────────────────────────────────
class FakeRange {
  constructor(sheet, row, col, nr, nc) {
    Object.assign(this, { sheet, row, col, nr, nc });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const row = [];
      for (let c = 0; c < this.nc; c++) {
        const src = this.sheet.data[this.row - 1 + r] || [];
        row.push(src[this.col - 1 + c] === undefined ? '' : src[this.col - 1 + c]);
      }
      out.push(row);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  // Книга ищет и колонку даты, и строки баланса по ОТОБРАЖАЕМОМУ тексту: в ячейке
  // шапки лежит серийная дата, а сравнивать надо с «31.08.2026».
  getDisplayValues() { return this.getValues().map(r => r.map(v => String(v == null ? '' : v))); }
  setValues(values) {
    if (values.length !== this.nr) {
      throw new Error(`setValues: строк ${values.length}, а диапазон на ${this.nr}`);
    }
    values.forEach(row => {
      if (row.length !== this.nc) {
        throw new Error(`setValues: колонок ${row.length}, а диапазон на ${this.nc}`);
      }
    });
    if (this.row + this.nr - 1 > this.sheet.maxRows) {
      throw new Error('setValues: за границей сетки листа');
    }
    values.forEach((row, r) => {
      const target = this.sheet.data[this.row - 1 + r] || (this.sheet.data[this.row - 1 + r] = []);
      row.forEach((v, c) => { target[this.col - 1 + c] = v; });
    });
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
}

class FakeSheet {
  constructor(name, data) {
    this.name = name;
    this.data = data || [];
    this.maxRows = Math.max(1000, this.data.length + 100);
    this.frozen = 0;
    this.cleared = 0;
    this.copies = 0;
  }
  getName() { return this.name; }
  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) {
      if ((this.data[i] || []).some(v => v !== '' && v !== undefined && v !== null)) return i + 1;
    }
    return 0;
  }
  getLastColumn() {
    return this.data.reduce((m, r) => Math.max(m, (r || []).length), 0);
  }
  getMaxRows() { return this.maxRows; }
  getRange(row, col, nr, nc) { return new FakeRange(this, row, col, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  setFrozenRows(n) { this.frozen = n; return this; }
  insertRowsAfter(after, n) { this.maxRows += n; return this; }
  deleteRows(from, count) { this.data.splice(from - 1, count); return this; }
  deleteRow(i) { return this.deleteRows(i, 1); }
  insertColumnBefore(col) {
    this.data.forEach(r => r && r.splice(col - 1, 0, ''));
    return this;
  }
  appendRow(row) { this.data.push(row.slice()); return this; }
  clear() { this.data = []; this.cleared++; return this; }
  copyTo() { this.copies++; return { setName: () => ({}) }; }
}

class FakeSpreadsheet {
  constructor(sheets) { this.sheets = sheets || {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { return (this.sheets[n] = new FakeSheet(n)); }
}

function makeEnv(opts) {
  opts = opts || {};
  const requests = [];
  const alerts = [];
  const ss = new FakeSpreadsheet(opts.sheets);

  function response(body, code, headers) {
    return {
      getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)),
      getResponseCode: () => (code === undefined ? 200 : code),
      getHeaders: () => headers || {},
    };
  }

  const routes = opts.routes || {};
  function serve(url, options) {
    requests.push({ url, options, payload: options && options.payload ? JSON.parse(options.payload) : null });
    for (const key of Object.keys(routes)) {
      if (url.indexOf(key) >= 0) {
        const r = routes[key];
        const out = typeof r === 'function'
          ? r(requests[requests.length - 1], requests.length)
          : r;
        if (out && out.__code) return response(out.body, out.__code, out.headers);
        return response(out === undefined ? {} : out);
      }
    }
    return response({});
  }

  const ctx = {
    console,
    // Date берём ИЗ ЭТОГО реалма: у vm свой глобальный Date, и `value instanceof
    // Date` в коде книги ломался бы на датах, созданных тестом. В Apps Script
    // реалм один, так что подмена здесь возвращает поведение к боевому.
    Date,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({
        alert: (...a) => { alerts.push(a.map(String).join(' | ')); return 'ok'; },
        prompt: () => ({ getSelectedButton: () => 'CANCEL', getResponseText: () => '' }),
        ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
        Button: { OK: 'OK', YES: 'YES' },
        createMenu: () => { const m = new Proxy({}, { get: () => () => m }); return m; },
        showModalDialog: () => {},
      }),
    },
    UrlFetchApp: {
      fetch: (url, options) => serve(url, options),
      fetchAll: reqs => reqs.map(r => serve(r.url, r)),
    },
    DriveApp: new Proxy({}, { get: () => () => ({}) }),
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => {} }) }) }), deleteTrigger: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: {
      sleep: () => {},
      getUuid: () => 'uuid',
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0');
        if (fmt === 'yyyy-MM-dd') return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
        if (fmt === 'dd.MM.yyyy') return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
        return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
      },
    },
    Session: { getScriptTimeZone: () => 'Europe/Moscow' },
    Logger: { log: () => {} },
  };
  vm.createContext(ctx);
  vm.runInContext(CODE, ctx, { filename: 'central.js' });
  return { ctx, ss, requests, alerts, call: (name, ...args) => vm.runInContext(name, ctx).apply(null, args) };
}

const CAB = { mp: 'OZON', id: 'КАБ', clientId: 'cid', apiKey: 'key',
              sheetName: 'Остатки по кластерам КАБ', rowIndex: 9, active: true };

function ozonRoutes(warehouses, analytics, postings) {
  let whDone = false, anDone = false;
  return {
    '/v2/analytics/stock_on_warehouses': () => {
      if (whDone) return { result: { rows: [] } };
      whDone = true;
      return { result: { rows: warehouses } };
    },
    '/v4/product/info/stocks': () => ({
      items: [{ stocks: (analytics || []).map(a => ({ sku: a.sku })) }], cursor: '',
    }),
    '/v1/analytics/stocks': () => {
      if (anDone) return { items: [] };
      anDone = true;
      return { items: analytics };
    },
    '/v2/posting/fbo/list': req => {
      const st = req.payload.filter.status;
      const list = (postings || []).filter(p => p.status === st);
      return { result: req.payload.offset ? [] : list };
    },
  };
}

// ─── 1. Нормализация имени склада ───────────────────────────────────────────
{
  const e = makeEnv({});
  const k = n => e.call('ozonKey_', 1, n);
  eq('ozonKey_: регистр и разделитель', k('Санкт_Петербург_РФЦ'), k('САНКТ-ПЕТЕРБУРГ_РФЦ'));
  eq('ozonKey_: Ё и Е', k('ОРЁЛ_РФЦ'), k('ОРЕЛ_РФЦ'));
  eq('ozonKey_: пробелы и точки', k('Ростов на Дону РФЦ'), k('РОСТОВ-НА-ДОНУ_РФЦ'));
  ok('ozonKey_: разные склады не склеиваются', k('КАЗАНЬ_РФЦ') !== k('КАЗАНЬ_2_РФЦ'));
  ok('ozonKey_: разные SKU не склеиваются',
     e.call('ozonKey_', 1, 'КАЗАНЬ') !== e.call('ozonKey_', 2, 'КАЗАНЬ'));
}

// ─── 2. Один склад под двумя именами — одна строка, без задвоения ───────────
{
  const sheet = new FakeSheet('Остатки по кластерам КАБ', null);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes(
      [{ sku: 1502995678, item_code: 'ART', item_name: 'Товар', warehouse_name: 'Санкт_Петербург_РФЦ',
         free_to_sell_amount: 27, promised_amount: 0, reserved_amount: 0 }],
      [{ sku: 1502995678, offer_id: 'ART', name: 'Товар', warehouse_name: 'САНКТ-ПЕТЕРБУРГ_РФЦ',
         cluster_name: 'Санкт-Петербург', available_stock_count: 27 }],
      []),
  });
  const n = e.call('fetchOzon', e.ss, CAB);
  const H = vm.runInContext('OZON_HEADERS', e.ctx);
  eq('склад под двумя именами: одна строка', n, 1);
  eq('склад под двумя именами: вверено 27, а не 54', sheet.data[1][H.indexOf('Всего вверено OZON')], 27);
  eq('склад под двумя именами: имя из аналитики', sheet.data[1][4], 'САНКТ-ПЕТЕРБУРГ_РФЦ');
  eq('ширина строки равна шапке', sheet.data[1].length, H.length);
  eq('шапка получила «Доставляем покупателям»', H.indexOf('Доставляем покупателям') >= 0, true);
  eq('«Всего вверено OZON» — последняя колонка шапки', H[H.length - 1], 'Всего вверено OZON');
}

// ─── 3. «Готовим к вывозу» не считается дважды ──────────────────────────────
{
  const sheet = new FakeSheet('Остатки по кластерам КАБ', null);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes(
      [{ sku: 1952753971, item_code: 'A', item_name: 'T', warehouse_name: 'ЯРОСЛАВЛЬ_РФЦ',
         free_to_sell_amount: 18, promised_amount: 0, reserved_amount: 0 }],
      [{ sku: 1952753971, offer_id: 'A', name: 'T', warehouse_name: 'ЯРОСЛАВЛЬ_РФЦ',
         available_stock_count: 1, return_to_seller_stock_count: 17, other_stock_count: 3 }],
      []),
  });
  e.call('fetchOzon', e.ss, CAB);
  const H = vm.runInContext('OZON_HEADERS', e.ctx);
  const row = sheet.data[1];
  eq('вывоз не задвоен: 18 + 3, а не 38', row[H.indexOf('Всего вверено OZON')], 21);
  eq('«Возврат продавцу» в листе остался как есть', row[H.indexOf('Возврат продавцу')], 17);
  eq('«Всего у OZON» сохранил прежний смысл', row[H.indexOf('Всего у OZON')], 18);
}

// ─── 4. Аналитика без складской строки: основа = available + return_to_seller ─
{
  const sheet = new FakeSheet('Остатки по кластерам КАБ', null);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes([], [{ sku: 7, offer_id: 'A', name: 'T', warehouse_name: 'ТРАНЗИТ',
                              available_stock_count: 4, return_to_seller_stock_count: 6 }], []),
  });
  e.call('fetchOzon', e.ss, CAB);
  const H = vm.runInContext('OZON_HEADERS', e.ctx);
  eq('строка только из аналитики: 4 + 6', sheet.data[1][H.indexOf('Всего вверено OZON')], 10);
}

// ─── 5. «Доставляем покупателям» ────────────────────────────────────────────
{
  const sheet = new FakeSheet('Остатки по кластерам КАБ', null);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes(
      [{ sku: 5, item_code: 'A', item_name: 'T', warehouse_name: 'ПЕРМЬ_РФЦ',
         free_to_sell_amount: 2, promised_amount: 0, reserved_amount: 3 }],
      [{ sku: 5, offer_id: 'A', name: 'T', warehouse_name: 'ПЕРМЬ_РФЦ', available_stock_count: 2 }],
      [{ status: 'delivering', analytics_data: { warehouse_name: 'ПЕРМЬ_РФЦ' },
         products: [{ sku: 5, quantity: 6 }] },
       { status: 'awaiting_deliver', analytics_data: { warehouse_name: 'Пермь_РФЦ' },
         products: [{ sku: 5, quantity: 1 }] },
       { status: 'awaiting_packaging', analytics_data: { warehouse_name: 'ПЕРМЬ_РФЦ' },
         products: [{ sku: 5, quantity: 3 }] },
       { status: 'delivered', analytics_data: { warehouse_name: 'ПЕРМЬ_РФЦ' },
         products: [{ sku: 5, quantity: 99 }] }]),
  });
  e.call('fetchOzon', e.ss, CAB);
  const H = vm.runInContext('OZON_HEADERS', e.ctx);
  const row = sheet.data[1];
  eq('доставляем = delivering + awaiting_deliver', row[H.indexOf('Доставляем покупателям')], 7);
  eq('вверено = 2 склад + 3 резерв + 7 доставка', row[H.indexOf('Всего вверено OZON')], 12);

  const posts = e.requests.filter(r => r.url.indexOf('posting/fbo/list') >= 0);
  const statuses = posts.map(r => r.payload.filter.status);
  ok('запрошены только два статуса',
     statuses.indexOf('delivering') >= 0 && statuses.indexOf('awaiting_deliver') >= 0 &&
     statuses.indexOf('awaiting_packaging') < 0, statuses.join(','));
  eq('лимит страницы отправлений в допустимых пределах',
     posts.every(r => r.payload.limit > 0 && r.payload.limit <= 100), true);
  ok('склад отправления запрошен', posts.every(r => r.payload.with.analytics_data === true));
  ok('окно просмотра — месяцы, а не дни', (() => {
    const p = posts[0].payload.filter;
    return (new Date(p.to) - new Date(p.since)) / 86400000 > 90;
  })());
}

// ─── 6. Повторный прогон дня и рост сетки ───────────────────────────────────
{
  const H0 = ['Дата и время', 'Артикул', 'SKU', 'Название товара', 'Склад',
    'Доступно к продаже', 'Готовим к продаже', 'Зарезервировано', 'Всего у OZON',
    'Доступно (аналитика)', 'В пути на склад', 'Заявлено к поставке',
    'Возврат от покупателя', 'Возврат продавцу', 'Брак на складе', 'Брак в пути',
    'Ждёт документов', 'Прочее', 'Доставляем покупателям', 'Кластер', 'Всего вверено OZON'];
  const today = new Date();
  const yesterday = new Date(today.getTime() - 86400000);
  const filler = n => Array.from({ length: 21 }, (_, i) => (i === 0 ? n : ''));
  const sheet = new FakeSheet('Остатки по кластерам КАБ',
    [H0.slice(), filler(yesterday), filler(today), filler(today)]);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes([{ sku: 9, item_code: 'A', item_name: 'T', warehouse_name: 'W',
                          free_to_sell_amount: 1, promised_amount: 0, reserved_amount: 0 }],
                       [{ sku: 9, offer_id: 'A', name: 'T', warehouse_name: 'W', available_stock_count: 1 }], []),
  });
  e.call('fetchOzon', e.ss, CAB);
  eq('вчерашний день на месте', sheet.data[1][0], yesterday);
  eq('сегодняшний блок переписан, а не удвоен', sheet.data.length, 3);

  const small = new FakeSheet('Остатки по кластерам КАБ', [H0.slice()]);
  small.maxRows = 1;
  const e2 = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': small },
    routes: ozonRoutes([{ sku: 9, item_code: 'A', item_name: 'T', warehouse_name: 'W',
                          free_to_sell_amount: 1, promised_amount: 0, reserved_amount: 0 }],
                       [{ sku: 9, offer_id: 'A', name: 'T', warehouse_name: 'W', available_stock_count: 1 }], []),
  });
  e2.call('fetchOzon', e2.ss, CAB);
  ok('сетка досоздана, день не потерян', small.maxRows > 1 && small.data.length === 2);
}

// ─── 7. Чужая раскладка листа: копия и новая шапка ──────────────────────────
{
  const sheet = new FakeSheet('Остатки по кластерам КАБ',
    [['Date', 'Article', 'SKU'], ['x', 'y', 'z']]);
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': sheet },
    routes: ozonRoutes([{ sku: 9, item_code: 'A', item_name: 'T', warehouse_name: 'W',
                          free_to_sell_amount: 1, promised_amount: 0, reserved_amount: 0 }],
                       [{ sku: 9, offer_id: 'A', name: 'T', warehouse_name: 'W', available_stock_count: 1 }], []),
  });
  e.call('fetchOzon', e.ss, CAB);
  eq('старый лист отложен копией', sheet.copies, 1);
  eq('лист начат заново', sheet.cleared, 1);
  eq('новая шапка на месте', sheet.data[0][18], 'Доставляем покупателям');
}

// ─── 8. Пустой ответ Ozon — это сбой, а не «остатков нет» ───────────────────
{
  const e = makeEnv({
    sheets: { 'Остатки по кластерам КАБ': new FakeSheet('Остатки по кластерам КАБ', null) },
    routes: ozonRoutes([], [], []),
  });
  let msg = '';
  try { e.call('fetchOzon', e.ss, CAB); } catch (err) { msg = err.message; }
  ok('ноль строк — ошибка, а не «✅ +0»', /0 строк/.test(msg), msg);
}

// ─── 9. Низкие остатки: колонки по заголовку и только свежий день ───────────
{
  const H0 = ['Дата и время', 'Артикул', 'SKU', 'Название товара', 'Склад',
    'Доступно к продаже', 'Готовим к продаже', 'Зарезервировано', 'Всего у OZON',
    'Доступно (аналитика)', 'В пути на склад', 'Заявлено к поставке',
    'Возврат от покупателя', 'Возврат продавцу', 'Брак на складе', 'Брак в пути',
    'Ждёт документов', 'Прочее', 'Доставляем покупателям', 'Кластер', 'Всего вверено OZON'];
  const mk = (d, name, qty) => {
    const r = new Array(21).fill('');
    r[0] = d; r[3] = name; r[4] = 'СКЛАД'; r[20] = qty;
    return r;
  };
  const today = new Date(), yest = new Date(Date.now() - 86400000);
  const sheet = new FakeSheet('Остатки по кластерам КАБ',
    [H0.slice(), mk(yest, 'Старая позиция', 1), mk(today, 'Заканчивается', 2), mk(today, 'Много', 500)]);
  const panel = new FakeSheet('Панель управления', [
    [], [], [], [], [], [], [], [],
    ['OZON', 'КАБ', 'cid', 'key', '', 'Остатки по кластерам КАБ', 'Да', '', '', ''],
  ]);
  const e = makeEnv({ sheets: { 'Остатки по кластерам КАБ': sheet, 'Панель управления': panel },
                      routes: {} });
  const n = e.call('checkLowStock');
  eq('низкие остатки: найдена одна позиция', n, 1);

  // Чужая раскладка: скрипт обязан промолчать, а не выдать цифры из чужих колонок
  const alien = new FakeSheet('Остатки по кластерам КАБ',
    [['Date', 'Article', 'Qty'], ['x', 'y', 1]]);
  const e2 = makeEnv({ sheets: { 'Остатки по кластерам КАБ': alien, 'Панель управления': panel } });
  eq('чужая раскладка — молчим, а не врём', e2.call('checkLowStock'), 0);
}

// ─── 10. Проверка связи: один fetchAll, без ретраев, диагноз словами ────────
{
  const panel = new FakeSheet('Панель управления', [
    [], [], [], [], [], [], [], [],
    ['OZON', 'A', 'cid', 'key', '', 'Остатки по кластерам A', 'Да', '', '', ''],
    ['OZON', 'B', '', '', '', 'Остатки по кластерам B', 'Да', '', '', ''],
    ['WB', 'C', '', '', 'tok', 'Остатки ВБ C', 'Да', '', '', ''],
    ['WB', 'D', '', '', 'tok', 'Остатки ВБ D', 'Нет', '', '', ''],
  ]);
  const e = makeEnv({
    sheets: { 'Панель управления': panel },
    routes: {
      '/v2/analytics/stock_on_warehouses': { __code: 403, body: { message: 'Api-key has expired' } },
      'warehouse_remains': { __code: 200, body: {} },
    },
  });
  const lines = e.call('panelCheckConnection');
  eq('проверены только активные кабинеты', lines.length, 3);
  ok('кабинет без ключей отмечен отдельно', /⚠️ OZON B/.test(lines[1]), lines[1]);
  ok('403 переведён на человеческий', /перевыпустить/.test(lines[0]), lines[0]);
  ok('живой кабинет — «связь есть»', /✅ WB C/.test(lines[2]), lines[2]);
  eq('запросов ровно по числу кабинетов с ключами — ретраев нет', e.requests.length, 2);
}

// ─── 11. Инструкция и статус: {html, text}, оба самодостаточны ──────────────
{
  const panel = new FakeSheet('Панель управления', [
    [], [], [], [], [], [], [], [],
    ['OZON', 'A', 'cid', 'key', '', 'Остатки по кластерам A', 'Да', '', '', ''],
  ]);
  const e = makeEnv({ sheets: { 'Панель управления': panel } });
  const help = e.call('panelHelp', '/* CENTRAL v3.5.0 */');
  ok('help отдаёт html', /<h1>/.test(help.html));
  ok('help отдаёт запасной текст', help.text.length > 300);
  ok('в help видна версия', help.html.indexOf('v3.5.0') >= 0);
  const st = e.call('panelStatus', '/* CENTRAL v3.5.0 */');
  ok('статус называет одно действие', /Что сейчас нужно от вас/.test(st.html));
  ok('в статусе есть кабинет', /\bA\b/.test(st.text));
  ok('статус зовёт вписать ключи, когда их нет', (() => {
    const p2 = new FakeSheet('Панель управления', [
      [], [], [], [], [], [], [], [],
      ['OZON', 'Z', '', '', '', 'Остатки по кластерам Z', 'Да', '', '', ''],
    ]);
    const e2 = makeEnv({ sheets: { 'Панель управления': p2 } });
    return /Впишите ключи/.test(e2.call('panelStatus', 'v').text);
  })());
}

// ─── 12. Миграция листов: вставляет колонку один раз ────────────────────────
{
  const old = ['Дата и время', 'Артикул', 'SKU', 'Название товара', 'Склад',
    'Доступно к продаже', 'Готовим к продаже', 'Зарезервировано', 'Всего у OZON',
    'Доступно (аналитика)', 'В пути на склад', 'Заявлено к поставке',
    'Возврат от покупателя', 'Возврат продавцу', 'Брак на складе', 'Брак в пути',
    'Ждёт документов', 'Прочее', 'Кластер', 'Всего вверено OZON'];
  const sheet = new FakeSheet('Остатки по кластерам A', [old.slice(), old.map((_, i) => i)]);
  const panel = new FakeSheet('Панель управления', [
    [], [], [], [], [], [], [], [],
    ['OZON', 'A', 'cid', 'key', '', 'Остатки по кластерам A', 'Да', '', '', ''],
  ]);
  const e = makeEnv({ sheets: { 'Остатки по кластерам A': sheet, 'Панель управления': panel } });
  const r1 = e.call('upgradeSheets');
  eq('миграция обновила лист', r1.done, ['A']);
  eq('колонка встала перед «Кластером»', sheet.data[0][18], 'Доставляем покупателям');
  eq('«Всего вверено OZON» уехало вправо', sheet.data[0][20], 'Всего вверено OZON');
  eq('данные старого дня сдвинулись вместе с шапкой', sheet.data[1][20], 19);
  const r2 = e.call('upgradeSheets');
  eq('повторная миграция ничего не делает', r2.done, []);
}

// ─── 13. Дебиторка: остатки кабинетов в баланс ──────────────────────────────
{
  const balance = new FakeSheet('Управленческий баланс Авто-тест', [
    ['Месяц —', 'Начало', '31.08.2026'],
    ['Дебиторская задолженность', '', ''],
    ['ИП1', '', ''],
    ['Остаток на балансе кабинета ВБ', '', 5],
    ['Ожидаем поступление на РС от ВБ', '', 111],
    ['Остаток на балансе кабинета ОЗОН', '', 0],
    ['Ожидаем поступление на РС от ОЗОН', '', 222],
    ['ИП3', '', ''],
    ['Остаток на балансе кабинета ВБ', '', 7],
    ['Остаток на балансе кабинета ОЗОН', '', 0],
    ['ИП11 50%', '', ''],
    ['Остаток на балансе кабинета ВБ', '', 0],
  ]);
  const panel = new FakeSheet('Панель управления', [
    [], [], [], [], [], [], [], [],
    ['OZON', 'ИП1', 'cid', 'key', '', 'Остатки по кластерам ИП1', 'Да', '', '', ''],
    ['WB',   'ИП1', '', '', 'tok', 'Остатки ВБ ИП1', 'Да', '', '', ''],
    ['WB',   'ИП3', '', '', 'tok3', 'Остатки ВБ ИП3', 'Да', '', '', ''],
    ['OZON', 'ИП3', 'cid3', 'key3', '', 'Остатки по кластерам ИП3', 'Нет', '', '', ''],
    ['WB',   'ИП11', '', '', 'tok11', 'Остатки ВБ ИП11', 'Да', '', '', ''],
  ]);
  const e = makeEnv({
    sheets: { 'Управленческий баланс Авто-тест': balance, 'Панель управления': panel },
    routes: {
      'finance-api.wildberries.ru': req => (req.options.headers.Authorization === 'tok'
        ? { currency: 'RUB', current: 1312834.31, for_withdraw: 0.23 }
        : { __code: 403, body: { detail: 'scope is not allowed' } }),
      '/v1/finance/cash-flow-statement/list': {
        result: { details: [
          { period: { begin: '2026-08-24T00:00:00Z', end: '2026-08-30T00:00:00Z' }, end_balance_amount: 1083200.38 },
          { period: { begin: '2026-08-31T00:00:00Z', end: '2026-08-31T00:00:00Z' }, end_balance_amount: 749319.12 },
        ] } },
    },
  });

  const plan = e.call('planCabinetBalances_', '31.08.2026');
  eq('дебиторка: колонка найдена по дате', plan.col, 3);
  eq('дебиторка: заполняемых строк', plan.write.length, 2);
  const wb = plan.write.filter(w => w.mp === 'WB')[0];
  const oz = plan.write.filter(w => w.mp === 'OZON')[0];
  eq('WB: берётся current', wb.value, 1312834.31);
  eq('WB: строка своего блока', [wb.row, wb.ip], [4, 'ИП1']);
  eq('Ozon: берётся САМЫЙ СВЕЖИЙ период, а не первый', oz.value, 749319.12);
  eq('Ozon: строка своего блока', [oz.row, oz.ip], [6, 'ИП1']);
  eq('прежнее значение прочитано', wb.was, 5);

  const why = {};
  plan.skip.forEach(s => { why[s.ip + '|' + s.mp] = s.why; });
  ok('403 у WB объяснён правами, а не кодом', /категории «Финансы»/.test(why['ИП3|WB']), why['ИП3|WB']);
  ok('выключенный кабинет пропущен', /выключен/.test(why['ИП3|OZON']), why['ИП3|OZON']);
  eq('строк «Ожидаем поступление» в плане нет',
     plan.write.concat(plan.skip).filter(x => x.row === 5 || x.row === 7).length, 0);
  eq('ИП11 распознан, доля помечена', plan.skip.filter(s => s.ip === 'ИП11').length, 1);

  // ключевое: пропущенная ячейка НЕ обнуляется
  const before = balance.data[8][2];
  e.call('syncCabinetBalances');
  eq('без подтверждения ничего не записано (prompt отменён)', balance.data[3][2], 5);
  eq('ячейка кабинета без прав не тронута', balance.data[8][2], before);

  const prev = e.call('previewCabinetBalances', 'v3.6.0');
  ok('предпросмотр отдаёт html', /<table>/.test(prev.html));
  ok('в предпросмотре объяснено, почему «Ожидаем поступление» пусто',
     /Ожидаем поступление/.test(prev.html));
  ok('в запасном тексте есть итог', /Итого/.test(prev.text));

  const e2 = makeEnv({ sheets: { 'Панель управления': panel } });
  let msg = '';
  try { e2.call('planCabinetBalances_', '01.01.2030'); } catch (err) { msg = err.message; }
  ok('нет колонки с датой — понятная ошибка', /нет колонки 01\.01\.2030/.test(msg), msg);
}

// ─── итог ───────────────────────────────────────────────────────────────────
console.log(`Тесты: ${passed}/${passed + failed}`);
if (failed) {
  console.log('\nНе прошли:');
  fails.forEach(f => console.log('  ✗ ' + f));
  process.exit(1);
}
