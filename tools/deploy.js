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
  // Тема коммита берётся из ПЕРВОЙ строки свежего блока истории версий: зашитый в
  // деплой текст один раз уже соврал про содержимое коммита (v3.6.0 уехала под
  // заголовком v3.5.0). Своя тема — флагом `--message`.
  const flag = args.indexOf('--message');
  const head = fs.readFileSync(central, 'utf8').split('\n').slice(1, 3).join(' ');
  const auto = (head.match(/v\d+\.\d+\.\d+:\s*([^(]+)/) || [])[1];
  const subject = flag >= 0 && args[flag + 1]
    ? args[flag + 1]
    : 'feat: ' + (auto ? auto.trim().toLowerCase() : 'обновление центрального кода');
  const msg = subject.slice(0, 100) + '\n\n' + version;
  execFileSync('git', ['commit', '-m', msg], { cwd: ROOT, stdio: 'inherit' });
  execFileSync('git', ['push'], { cwd: ROOT, stdio: 'inherit' });
  console.log('   ✅ центральный код выложен');
}

/**
 * Проверка ЖИВОГО файла — ровно то, что сделает книга при первом клике.
 *
 * «Успех» — это проверенный результат, а не «git push не упал»: файл может уехать
 * в другую ветку, отдаться из кэша CDN старым или не резолвить свежую функцию.
 * Поэтому берём его так же, как лоадер, и достаём каждую точку входа меню.
 */
function verifyLive() {
  const loader = fs.readFileSync(path.join(ROOT, MODULE.loader), 'utf8');
  const owner = (loader.match(/GH_OWNER\s*=\s*'([^']+)'/) || [])[1];
  const repo = (loader.match(/GH_REPO\s*=\s*'([^']+)'/) || [])[1];
  const file = (loader.match(/GH_FILE\s*=\s*'([^']+)'/) || [])[1];
  const branch = (loader.match(/GH_BRANCH\s*=\s*'([^']+)'/) || [])[1];

  console.log('▸ Живой файл ' + owner + '/' + repo + '@' + branch);
  return new Promise((resolve, reject) => {
    https.get({ hostname: 'raw.githubusercontent.com',
                path: '/' + owner + '/' + repo + '/' + branch + '/' + file },
      res => {
        if (res.statusCode !== 200) return reject(new Error('репозиторий отдал ' + res.statusCode +
          ' — книга получит ту же ошибку'));
        let src = '';
        res.on('data', c => { src += c; });
        res.on('end', () => {
          const names = new Set();
          const re = /run_\('([A-Za-z_$][\w$]*)'/g;
          let m;
          while ((m = re.exec(loader))) names.add(m[1]);
          const bad = [];
          for (const n of names) {
            let f = null;
            try {
              f = new Function(src + '\n;return (typeof ' + n + ' === "function" ? ' + n + ' : null);')();
            } catch (e) { bad.push(n + ' (' + e.message + ')'); continue; }
            if (!f) bad.push(n);
          }
          if (bad.length) return reject(new Error('в живом файле не резолвятся: ' + bad.join(', ')));
          console.log('   ' + src.split('\n')[0]);
          console.log('   ✅ все ' + names.size + ' точек входа резолвятся');
          resolve();
        });
      }).on('error', reject);
  });
}

(async () => {
  step('Сборка', 'node', ['tools/build.js']);
  step('Тесты', 'node', ['tests/run.js']);
  if (!ONLY_BOOK) deployRepo();
  if (PUSH && !ONLY_BOOK) await verifyLive();
  if (!ONLY_REPO) await deployBook();
  if (!PUSH) console.log('\nСухой прогон. Выложить: node tools/deploy.js --push');
})().catch(e => die(e.message));
