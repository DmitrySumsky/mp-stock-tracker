/**
 * Деплой модуля «Пульт»: центральный код — в репозиторий, лоадер — в книгу.
 *
 * Порядок жёсткий (ПУЛЬТ §11): сборка → тесты → выкладка. Деплой смотрит на КОД
 * ВОЗВРАТА сборки и тестов, а не на последние строки их вывода: сборка, не
 * прошедшая проверки, файл на диске не оставляет, и уехала бы прошлая версия.
 *
 * Лоадер заливается PUT-ом полного состава через Apps Script API, а не `clasp push`:
 * clasp 3.3 НЕ отправляет удаление файла, если больше ничего не менялось — отвечает
 * «Script is already up to date», а файл остаётся на сервере. Здесь это принципиально:
 * в книге лежали 14 файлов прошлой схемы, и они объявляют ТЕ ЖЕ точки входа, что и
 * лоадер. Два определения одного имени в проекте Apps Script затирают друг друга, а
 * какое победит — решает порядок файлов.
 *
 *   node tools/deploy.js                 # сухой прогон: что уедет и что заменит
 *   node tools/deploy.js --push          # выложить
 *   node tools/deploy.js --push --book   # только книга (репозиторий не трогать)
 */
'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const { execFileSync, spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const MODULE = JSON.parse(fs.readFileSync(path.join(ROOT, 'module.json'), 'utf8'));
const CLASPRC = path.join(os.homedir(), '.clasprc.json');
const CLASP_JSON = path.join(ROOT, '.clasp.json');

const args = process.argv.slice(2);
const PUSH = args.includes('--push');
const ONLY_BOOK = args.includes('--book');
const ONLY_REPO = args.includes('--repo');

const TYPES = { '.js': 'SERVER_JS', '.gs': 'SERVER_JS', '.json': 'JSON', '.html': 'HTML' };

function die(msg) { console.error('ДЕПЛОЙ ОСТАНОВЛЕН: ' + msg); process.exit(1); }

function step(title, cmd, cmdArgs) {
  console.log('▸ ' + title);
  const r = spawnSync(cmd, cmdArgs, { cwd: ROOT, encoding: 'utf8', shell: process.platform === 'win32' });
  process.stdout.write((r.stdout || '').split('\n').map(l => '   ' + l).join('\n').trimEnd() + '\n');
  if (r.status !== 0) {
    process.stderr.write(r.stderr || '');
    die(title + ' — код возврата ' + r.status);
  }
}

// ─── Apps Script API ────────────────────────────────────────────────────────
function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => { data += c; });
      res.on('end', () => (res.statusCode >= 200 && res.statusCode < 300
        ? resolve(data ? JSON.parse(data) : {})
        : reject(new Error('HTTP ' + res.statusCode + ': ' + data.slice(0, 400)))));
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

/** Токен обновляется из ~/.clasprc.json тем же аккаунтом, которым авторизован clasp. */
async function scriptApiToken() {
  if (!fs.existsSync(CLASPRC)) die('нет ~/.clasprc.json — выполните `clasp login`');
  const t = JSON.parse(fs.readFileSync(CLASPRC, 'utf8')).tokens.default;
  const body = new URLSearchParams({
    client_id: t.client_id, client_secret: t.client_secret,
    refresh_token: t.refresh_token, grant_type: 'refresh_token',
  }).toString();
  const res = await request({
    hostname: 'oauth2.googleapis.com', path: '/token', method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
  }, body);
  return res.access_token;
}

function scriptApi(scriptId, token, payload) {
  const body = payload ? JSON.stringify(payload) : null;
  return request({
    hostname: 'script.googleapis.com', path: '/v1/projects/' + scriptId + '/content',
    method: payload ? 'PUT' : 'GET',
    headers: Object.assign({ Authorization: 'Bearer ' + token },
      body ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) } : {}),
  }, body);
}

/** Состав проекта книги: лоадер и манифест. `.clasp.json` в состав НЕ кладём —
 *  Apps Script API отвечает на него 400. */
function loaderFiles() {
  const out = [];
  for (const rel of [MODULE.loader, MODULE.manifest]) {
    const abs = path.join(ROOT, rel);
    if (!fs.existsSync(abs)) die('нет файла ' + rel);
    const ext = path.extname(abs);
    if (!TYPES[ext]) die('неизвестный тип файла ' + rel);
    out.push({ name: path.basename(abs, ext), type: TYPES[ext],
               source: fs.readFileSync(abs, 'utf8') });
  }
  return out;
}

/** Точки входа лоадера. Если это же имя объявляет ещё какой-то файл проекта,
 *  победит порядок файлов — а он не наш. Такой пуш останавливаем. */
function collisions(remote, local) {
  const names = new Set();
  local.forEach(f => {
    const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
    let m;
    while ((m = re.exec(f.source))) names.add(m[1]);
  });
  const localNames = new Set(local.map(f => f.name));
  const hits = [];
  remote.filter(f => !localNames.has(f.name)).forEach(f => {
    const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
    let m;
    while ((m = re.exec(f.source || ''))) {
      if (names.has(m[1])) hits.push(f.name + ' → ' + m[1]);
    }
  });
  return hits;
}

async function deployBook() {
  if (!fs.existsSync(CLASP_JSON)) die('нет .clasp.json со scriptId книги');
  const scriptId = JSON.parse(fs.readFileSync(CLASP_JSON, 'utf8')).scriptId;
  const token = await scriptApiToken();
  const remote = (await scriptApi(scriptId, token)).files || [];
  const local = loaderFiles();

  console.log('▸ Книга ' + scriptId);
  console.log('   на сервере сейчас: ' + remote.map(f => f.name).join(', '));
  console.log('   уедет: ' + local.map(f => f.name).join(', '));
  const gone = remote.filter(f => !local.some(l => l.name === f.name)).map(f => f.name);
  if (gone.length) console.log('   будет УДАЛЕНО: ' + gone.join(', '));

  const bad = collisions(remote.filter(f => !gone.includes(f.name)), local);
  if (bad.length) die('имя объявлено дважды: ' + bad.join('; '));

  if (!PUSH) { console.log('   (сухой прогон — ничего не отправлено)'); return; }
  await scriptApi(scriptId, token, { files: local });
  console.log('   ✅ состав книги заменён на лоадер');
}

function deployRepo() {
  const central = path.join(ROOT, MODULE.central);
  if (!fs.existsSync(central)) die('нет собранного центрального файла');
  const version = fs.readFileSync(central, 'utf8').split('\n')[0];

  console.log('▸ Репозиторий');
  console.log('   файл: ' + MODULE.central);
  console.log('   версия: ' + version);
  const status = execFileSync('git', ['status', '--short'], { cwd: ROOT, encoding: 'utf8' });
  console.log('   изменения:\n' + (status.trim() ? status.trimEnd().split('\n').map(l => '     ' + l).join('\n') : '     нет'));
  if (!PUSH) { console.log('   (сухой прогон — коммит и push не делались)'); return; }

  execFileSync('git', ['add', '-A'], { cwd: ROOT, stdio: 'inherit' });
  const msg = 'feat(ozon): остатки считаются верно — склад под двумя именами, ' +
              'двойной вывоз и «доставляем покупателям»\n\n' + version;
  execFileSync('git', ['commit', '-m', msg], { cwd: ROOT, stdio: 'inherit' });
  execFileSync('git', ['push'], { cwd: ROOT, stdio: 'inherit' });
  console.log('   ✅ центральный код выложен');
}

(async () => {
  step('Сборка', 'node', ['tools/build.js']);
  step('Тесты', 'node', ['tests/run.js']);
  if (!ONLY_BOOK) deployRepo();
  if (!ONLY_REPO) await deployBook();
  if (!PUSH) console.log('\nСухой прогон. Выложить: node tools/deploy.js --push');
})().catch(e => die(e.message));
