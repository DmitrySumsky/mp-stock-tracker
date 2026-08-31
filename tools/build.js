/**
 * Сборка центрального файла из модулей — со встроенными проверками (ПУЛЬТ §11).
 *
 * Сборка, не прошедшая проверки, НЕ ОСТАВЛЯЕТ файл на диске: иначе деплой уедет
 * с последней удачной версией и никто этого не заметит. Деплой смотрит на код
 * возврата сборки, а не на последние строки вывода.
 *
 *   node tools/build.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'central');
const OUT = path.join(SRC, 'build', 'central.js');

// Публичные точки входа: их вызывает лоадер по имени. Не разрешилась хотя бы одна —
// в книге это «В центральном коде нет функции …» под нажатием человека.
const ENTRY_POINTS = [
  'runAllSupplyFunctions', 'runAllOzon', 'runAllWB', 'runSelectedSupplyFunctions',
  'runSingleFromMenu', 'addNewCabinet',
  'fetchMsStock', 'createMsDailyTrigger', 'removeMsDailyTrigger',
  'previewCreditsSync', 'syncCreditsToBalance', 'syncCreditsToSpecificColumn',
  'updateDashboard', 'exportDashboardCSV', 'checkLowStock', 'healthCheck',
  'setupTriggers', 'removeTriggers', 'cleanupOldData', 'patchPanel',
  'migrateFromOldConfig', 'initPanel',
  'panelHelp', 'panelStatus', 'panelCheckConnection', 'upgradeSheets',
];

// Скан утечек. Секрет в коде — это отказ сборки: ключи живут только на листе книги
// и восстановить их из истории репозитория уже нельзя.
const LEAKS_HARD = [
  [/\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/i, 'похоже на API-ключ (UUID)'],
  [/\beyJ[A-Za-z0-9_-]{20,}/, 'похоже на JWT-токен'],
  [/\bgh[pousr]_[A-Za-z0-9]{20,}/, 'похоже на токен GitHub'],
  [/\b\d{9,10}:[A-Za-z0-9_-]{30,}/, 'похоже на токен Telegram-бота'],
  [/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/, 'адрес почты'],
];
// А это — предупреждение, а не отказ. Обозначения кабинетов в комментариях и в
// истории версий и есть память проекта («78 шт по ИП1» объясняет, откуда взялась
// правка). Сборку они не валят, но если центральный файл поедет в ПУБЛИЧНЫЙ
// репозиторий, список ниже — то, что придётся обезличить.
const LEAKS_SOFT = [
  [/(?:^|[^\w])ИП\s?\d{1,2}(?![\w])/m, 'обозначение кабинета (ИПn)'],
];
// Строки, где перечисленное выше законно: примеры в подсказках и текстах ошибок.
const LEAK_ALLOW = [
  /например: ИП12/i,
  /OZON ИП|WB ИП/,            // шаблоны сообщений вида `${cab.mp} ${cab.id}`
];

function fail(msg) {
  console.error('СБОРКА НЕ ПРОШЛА: ' + msg);
  process.exit(1);
}

function main() {
  const files = fs.readdirSync(SRC)
    .filter(f => f.endsWith('.js') && f !== 'preamble.js')
    .sort();
  if (!files.length) fail('в central/ нет модулей');

  const preamble = fs.readFileSync(path.join(SRC, 'preamble.js'), 'utf8');
  if (!/^\/\* .+ v\d+\.\d+\.\d+ — \d{2}\.\d{2}\.\d{4} \*\/$/m.test(preamble.split('\n')[0])) {
    fail('первая строка preamble.js должна быть строкой версии вида ' +
         '"/* … vX.Y.Z — ДД.ММ.ГГГГ */" — её читает пункт «Версия кода»');
  }

  const parts = [preamble.trimEnd(), ''];
  for (const f of files) {
    parts.push('// ══════ ' + f + ' ' + '═'.repeat(Math.max(0, 46 - f.length)));
    parts.push(fs.readFileSync(path.join(SRC, f), 'utf8').trimEnd(), '');
  }
  const code = parts.join('\n') + '\n';

  // 1. синтаксис
  try {
    new vm.Script(code, { filename: 'central.js' });
  } catch (e) {
    fail('синтаксис: ' + e.message);
  }

  // 2. точки входа резолвятся — ровно так, как это делает run_() в книге
  const missing = [];
  const stub = new Proxy({}, { get: () => () => stub });
  const ctx = vm.createContext({
    SpreadsheetApp: stub, UrlFetchApp: stub, DriveApp: stub, ScriptApp: stub,
    PropertiesService: stub, CacheService: stub, LockService: stub,
    Utilities: stub, Session: stub, Logger: stub, console: console,
  });
  let resolver;
  try {
    resolver = vm.runInContext(
      '(function(){' + code + '\n;return function(n){return typeof this[n];};})()', ctx);
  } catch (e) {
    fail('код не выполняется на верхнем уровне: ' + e.message);
  }
  for (const name of ENTRY_POINTS) {
    const probe = vm.runInContext(
      '(function(){' + code + '\n;return typeof ' + name + ';})()', ctx);
    if (probe !== 'function') missing.push(name);
  }
  if (missing.length) fail('не разрешились точки входа: ' + missing.join(', '));

  // 3. дублирование имён: два определения одного имени молча затирают друг друга
  const seen = new Map();
  const dup = [];
  const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
  let m;
  while ((m = re.exec(code))) {
    if (seen.has(m[1])) dup.push(m[1]);
    seen.set(m[1], true);
  }
  if (dup.length) fail('имя объявлено дважды: ' + [...new Set(dup)].join(', '));

  // 4. утечки
  const leaks = [], soft = [];
  code.split('\n').forEach((line, i) => {
    if (LEAK_ALLOW.some(rx => rx.test(line))) return;
    LEAKS_HARD.forEach(([rx, what]) => {
      if (rx.test(line)) leaks.push(`  строка ${i + 1}: ${what} — ${line.trim().slice(0, 90)}`);
    });
    LEAKS_SOFT.forEach(([rx, what]) => {
      if (rx.test(line)) soft.push(`  строка ${i + 1}: ${what}`);
    });
  });
  if (leaks.length) fail('скан утечек:\n' + leaks.join('\n'));
  if (soft.length) {
    console.log(`Предупреждение: ${soft.length} строк с обозначениями кабинетов — ` +
                'для приватного репозитория это норма, для публичного обезличить.');
  }

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, code, 'utf8');
  console.log('Собрано: ' + path.relative(ROOT, OUT));
  console.log('  модулей: ' + files.length + ', строк: ' + code.split('\n').length +
              ', функций: ' + seen.size);
  console.log('  ' + code.split('\n')[0]);
}

main();
