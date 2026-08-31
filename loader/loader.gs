/**
 * ОСТАТКИ И БАЛАНС — ЗАГРУЗЧИК v1 (лежит в книге «Баланс Маркеты»).
 *
 * Стандарт «Пульт» (57.mp-core/ПУЛЬТ.md): в книге — только этот файл, вся логика
 * одним файлом в репозитории. Правка логики = один пуш, книга на новой версии со
 * следующего клика; лоадер трогается ТОЛЬКО ради нового пункта меню, потому что
 * run_() резолвит функции центрального кода по имени.
 *
 * Настройка книги — лист «Панель управления»:
 *   B2 / B3 — токен Telegram-бота и адреса чатов через «;»
 *   B5      — токен МойСклада
 *   B6      — GitHub-токен (fine-grained PAT, право Contents: Read-only)
 *   строки 9+ — кабинеты: маркетплейс, ИП, ключи, имя листа, «Активен»
 *
 * v1: перевод книги на «Пульт». До этого все 14 модулей лежали в самой книге,
 *     и любая правка требовала clasp push в чужой боевой скрипт.
 */

var GH_OWNER  = 'DmitrySumsky';
var GH_REPO   = 'mp-stock-tracker';
var GH_FILE   = 'central/build/central.js';
var GH_BRANCH = 'master';

/** Ячейка листа «Панель управления» с токеном GitHub. */
var GH_TOKEN_CELL = 'B6';
var PANEL_SHEET_NAME = 'Панель управления';

/* ---------- дальше ничего менять не нужно ---------- */

/** НЕ ВЫЗЫВАЕТСЯ. Скоупы OAuth выдаются по СТАТИЧЕСКОМУ анализу кода, а загруженный
 *  по сети центральный код анализатор не видит: эта функция «показывает» все сервисы,
 *  включая операции ЗАПИСИ (иначе Google выдаст «только чтение»). */
function __scopes_() {
  SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.create('x');
  DriveApp.getRootFolder();
  DriveApp.getFileById('x');
  DriveApp.createFolder('x');
  DriveApp.createFile('x', 'x');
  ScriptApp.getProjectTriggers();
  ScriptApp.newTrigger('x');
  PropertiesService.getScriptProperties();
  CacheService.getScriptCache();
  LockService.getScriptLock();
  UrlFetchApp.fetch('https://example.com');
  HtmlService.createHtmlOutput('x');
  Utilities.getUuid();
  Session.getScriptTimeZone();
}

/** GitHub-токен: «Панель управления» B6 > Script Properties (GITHUB_TOKEN). */
function ghToken_() {
  try {
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PANEL_SHEET_NAME);
    if (s) { var t = String(s.getRange(GH_TOKEN_CELL).getValue()).trim(); if (t) return t; }
  } catch (e) {}
  var p = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  return p ? String(p).trim() : '';
}

/** Кэш кода. У CacheService потолок 100 КБ НА ЗНАЧЕНИЕ, а центральный файл уже 117 КБ
 *  и будет расти — целиком он не кладётся, и `put` молча отказывает. Поэтому режем на
 *  куски и храним их число отдельным ключом: потерялся хоть один кусок — считаем, что
 *  кэша нет, и идём в сеть.
 *
 *  Размер куска — в СИМВОЛАХ, а потолок Google — в БАЙТАХ. Комментарии и тексты окон
 *  здесь кириллические, то есть по два байта на символ: 30 000 символов дают максимум
 *  60 КБ и в потолок укладываются с запасом даже на сплошной кириллице. */
var CACHE_CHUNK = 30000;
var CACHE_TTL   = 1800;

function cacheGetCode_() {
  try {
    var c = CacheService.getScriptCache();
    var n = Number(c.get('central_parts') || 0);
    if (!n) return '';
    var keys = [];
    for (var i = 0; i < n; i++) keys.push('central_' + i);
    var got = c.getAll(keys);
    var out = '';
    for (var j = 0; j < n; j++) {
      if (got['central_' + j] === undefined || got['central_' + j] === null) return '';
      out += got['central_' + j];
    }
    return out;
  } catch (e) { return ''; }
}

function cachePutCode_(src) {
  try {
    var c = CacheService.getScriptCache(), map = {}, n = 0;
    for (var i = 0; i < src.length; i += CACHE_CHUNK) {
      map['central_' + n] = src.substring(i, i + CACHE_CHUNK);
      n++;
    }
    map.central_parts = String(n);
    c.putAll(map, CACHE_TTL);
  } catch (e) { /* не закэшировалось — просто сходим в сеть в следующий раз */ }
}

/** Центральный код: сначала по токену (у токена своя квота), иначе — публично.
 *  Без токена GitHub даёт 60 запросов в час на IP, а IP у Apps Script общий на
 *  всех — поэтому анонимный путь только запасной, а не основной. */
function fetchCode_() {
  var cached = cacheGetCode_();
  if (cached) return cached;

  var tok = ghToken_();
  var url, headers;
  if (tok) {
    url = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO +
          '/contents/' + GH_FILE + '?ref=' + GH_BRANCH;
    headers = { 'Authorization': 'Bearer ' + tok,
                'Accept': 'application/vnd.github.raw+json',
                'X-GitHub-Api-Version': '2022-11-28' };
  } else {
    url = 'https://raw.githubusercontent.com/' + GH_OWNER + '/' + GH_REPO + '/' +
          GH_BRANCH + '/' + GH_FILE;
    headers = {};
  }

  var resp = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
  var code = resp.getResponseCode(), src = resp.getContentText();

  if (code === 401) throw new Error('GitHub: токен не принят (401) — истёк или неверный. ' +
    'Обновите «Панель управления» ' + GH_TOKEN_CELL + '.');
  if (code === 403) throw new Error('GitHub: нет доступа (403). Если ячейка ' + GH_TOKEN_CELL +
    ' пуста — это исчерпанный анонимный лимит, впишите токен с правом Contents: Read-only.');
  if (code === 404) throw new Error('GitHub: файл не найден (404) — проверьте репозиторий, файл и ветку.');
  if (code !== 200) throw new Error('GitHub: ошибка ' + code + ': ' + String(src).slice(0, 200));
  if (!src || src.indexOf('function') < 0) throw new Error('В файле кода не код.');

  // Полчаса кэша: за один сеанс работы человек жмёт несколько пунктов подряд,
  // и каждый из них не должен ходить в GitHub заново.
  cachePutCode_(src);
  return src;
}

/** Сбросить кэш кода — нужен сразу после выкладки новой версии. */
function dropCodeCache_() {
  try {
    var c = CacheService.getScriptCache();
    var n = Number(c.get('central_parts') || 0);
    var keys = ['central_parts'];
    for (var i = 0; i < n; i++) keys.push('central_' + i);
    c.removeAll(keys);
  } catch (e) {}
}

/** Резолв функции центрального кода ПО ИМЕНИ: новая функция не требует правки
 *  лоадера, правки требует только новая КНОПКА. */
function run_(name, args) {
  var body = fetchCode_() + '\n;return (typeof ' + name + ' === "function" ? ' + name + ' : null);';
  var f = (new Function(body))();
  if (!f) throw new Error('В центральном коде нет функции ' + name + ' — обновите файл кода.');
  return f.apply(null, args || []);
}

/** HTML-диалог. HtmlService работает только из СТАТИЧЕСКОГО кода проекта, поэтому
 *  окно рисует лоадер, а центральный код отдаёт {html, text}. Не открылось —
 *  тот же текст через alert: инструкция обязана дойти до человека в любом случае. */
function showHtml_(res, title, w, h) {
  var ui = SpreadsheetApp.getUi();
  var html = res && res.html ? String(res.html) : '';
  var text = res && res.text ? String(res.text) : String(res || '');
  try {
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(w || 880).setHeight(h || 640), title);
  } catch (e) {
    ui.alert(title, text + '\n\n(Окно с оформлением не открылось: ' +
      String(e.message || e).slice(0, 200) + ')', ui.ButtonSet.OK);
  }
}

/** Первая строка центрального файла — строка версии. */
function centralVersion_() {
  try { return fetchCode_().split('\n')[0]; } catch (e) { return String(e.message || e); }
}

/** Меню — в порядке работы. 🔴 в этой книге нет: она только ЧИТАЕТ у
 *  маркетплейсов и пишет в саму себя, наружу не отправляется ничего. */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Остатки и баланс')
    .addItem('1️⃣ 🔄 Обновить остатки — все кабинеты', 'runAllSupplyFunctions')
    .addItem('2️⃣ 🏪 Обновить остатки МойСклад',        'fetchMsStock')
    .addItem('3️⃣ 💳 Кредиты в баланс…',                'syncCreditsToBalance')
    .addSeparator()
    .addItem('📖 Как работать (инструкция)', 'mHelp')
    .addItem('📊 Что сейчас происходит',     'mStatus')
    .addItem('🔌 Проверка связи',            'mCheck')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠 Ручной режим')
      .addItem('Только OZON',        'runAllOzon')
      .addItem('Только WB',          'runAllWB')
      .addItem('Выбрать кабинеты…',  'runSelectedSupplyFunctions')
      .addItem('Один кабинет…',      'runSingleFromMenu')
      .addItem('Добавить кабинет',   'addNewCabinet'))
    .addSubMenu(ui.createMenu('📚 Диагностика и справочники')
      .addItem('Обновить дашборд',          'updateDashboard')
      .addItem('Экспорт дашборда в CSV',    'exportDashboardCSV')
      .addItem('Проверить низкие остатки',  'checkLowStock')
      .addItem('Предпросмотр кредитов',     'previewCreditsSync')
      .addItem('Кредиты в конкретный столбец…', 'syncCreditsToSpecificColumn'))
    .addSubMenu(ui.createMenu('⏰ Автозапуск')
      .addItem('Настроить обновление остатков…',   'setupTriggers')
      .addItem('МойСклад: включить ежедневно',     'createMsDailyTrigger')
      .addItem('МойСклад: отключить ежедневно',    'removeMsDailyTrigger')
      .addItem('Удалить ВСЕ триггеры',             'removeTriggers'))
    .addSubMenu(ui.createMenu('🧹 Обслуживание')
      .addItem('Очистка старых данных…',       'cleanupOldData')
      .addItem('Обновить панель управления',   'patchPanel')
      .addItem('Миграция из старых листов',    'migrateFromOldConfig'))
    .addSeparator()
    .addItem('⚙️ Обновить настройки таблицы', 'mUpgrade')
    .addItem('ℹ️ Версия кода',                'mVersion')
    .addToUi();
}

/* ── точки входа ───────────────────────────────────────────────────────────
   Имена НЕ переименовывать: на них висят пункты меню и уже установленные
   триггеры (`runAllSupplyFunctions`, `fetchMsStock`). ───────────────────── */

function runAllSupplyFunctions()      { return run_('runAllSupplyFunctions'); }
function runAllOzon()                 { return run_('runAllOzon'); }
function runAllWB()                   { return run_('runAllWB'); }
function runSelectedSupplyFunctions() { return run_('runSelectedSupplyFunctions'); }
function runSingleFromMenu()          { return run_('runSingleFromMenu'); }
function addNewCabinet()              { return run_('addNewCabinet'); }

function fetchMsStock()          { return run_('fetchMsStock'); }
function createMsDailyTrigger()  { return run_('createMsDailyTrigger'); }
function removeMsDailyTrigger()  { return run_('removeMsDailyTrigger'); }

function previewCreditsSync()          { return run_('previewCreditsSync'); }
function syncCreditsToBalance()        { return run_('syncCreditsToBalance'); }
function syncCreditsToSpecificColumn() { return run_('syncCreditsToSpecificColumn'); }

function updateDashboard()    { return run_('updateDashboard'); }
function exportDashboardCSV() { return run_('exportDashboardCSV'); }
function checkLowStock()      { return run_('checkLowStock'); }
function healthCheck()        { return run_('panelCheckConnection'); }

function setupTriggers()        { return run_('setupTriggers'); }
function removeTriggers()       { return run_('removeTriggers'); }
function cleanupOldData()       { return run_('cleanupOldData'); }
function patchPanel()           { return run_('patchPanel'); }
function migrateFromOldConfig() { return run_('migrateFromOldConfig'); }

function mCheck()   { return run_('panelCheckConnection'); }
function mUpgrade() { dropCodeCache_(); return run_('upgradeSheets'); }

function mHelp()   { showHtml_(run_('panelHelp',   [centralVersion_()]), '📖 Как работать', 900, 700); }
function mStatus() { showHtml_(run_('panelStatus', [centralVersion_()]), '📊 Что сейчас происходит', 860, 620); }

function mVersion() {
  var ui = SpreadsheetApp.getUi();
  dropCodeCache_();
  ui.alert('Версия кода', centralVersion_() + '\n\nИсточник: GitHub ' + GH_OWNER + '/' +
    GH_REPO + '/' + GH_FILE + '@' + GH_BRANCH + '\nЗагрузчик: v1', ui.ButtonSet.OK);
}
