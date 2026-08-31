/* MP STOCK TRACKER — CENTRAL CODE v3.6.0 — 31.08.2026 */
/* v3.6.0: ДЕБИТОРКА ПО КАБИНЕТАМ ЗАПОЛНЯЛАСЬ РУКАМИ И ПОТОМУ НЕ ЗАПОЛНЯЛАСЬ
   (запрос владельца в тот же день: «нужно ещё подтянуть данные по статьям
   Дебиторская задолженность, Остаток на балансе кабинета ВБ / ОЗОН, Ожидаем
   поступление на РС»).

   ЧТО БЫЛО. В свежей колонке баланса все десять блоков ИП стояли по нулю, хотя
   площадки должны холдингу больше двух миллионов: цифры переносились из кабинетов
   вручную и на последний срез просто не попали.

   ЧТО СТАЛО. Пункт «3️⃣ 💰 Остатки кабинетов в дебиторку» (`14_receivables.js`):
   • WB — `finance-api…/v1/account/balance`, поле `current`. Требует у токена
     категорию «Финансы»: токен под «Статистику» и «Аналитику» отвечает 403
     `scope is not allowed`;
   • Ozon — `/v1/finance/cash-flow-statement/list` с `with_details`, поле
     `end_balance_amount` САМОГО СВЕЖЕГО периода. Именно оно ведёт себя как «Баланс»
     кабинета: пока неделя не закрыта, Ozon отдаёт отдельный период «сегодня…сегодня».
     Спрашивать прошедшую дату бесполезно — на середину недели возвращается период
     целиком, вместе с ещё не наступившими днями;
   • строки ищутся ПО НАЗВАНИЮ и привязываются к блоку ИП по заголовку выше — номера
     в балансе не зашиты, туда регулярно вставляют строки;
   • кабинет, чей ключ отказал, ПРОПУСКАЕТСЯ с причиной словами, а ячейка остаётся
     как была. Обнулить чужую цифру из-за протухшего токена — худшее, что тут можно
     сделать: ноль в балансе неотличим от факта;
   • сухой прогон и подтверждение С ЦИФРОЙ до записи, каждая строка — в «Лог»;
   • заголовок блока с долей («ИП11 50%») распознаётся и помечается: строка называет
     ПОЛНЫЙ остаток кабинета, а решение о доле принимает владелец, не скрипт.

   ЧЕГО СКРИПТ НЕ ДЕЛАЕТ И ПОЧЕМУ. Строки «Ожидаем поступление на РС от ВБ / ОЗОН»
   не заполняются вовсе. Это деньги, которые площадка уже отправила, но которые ещё
   не дошли до счёта, и публичного метода для них нет НИ У ОДНОЙ площадки (проверено
   31.08.2026: у WB `…/withdrawals`, `…/payments`, `…/account/withdrawal` — 404;
   у Ozon `payments` отчёта о взаиморасчётах это УЖЕ сделанные выплаты, а
   `/v1/finance/payouts` и `/v1/finance/treasury/totals` не существуют). Единственный
   источник статуса заявки — личный кабинет. Пустая ячейка честнее выдуманной, и это
   написано в инструкции, чтобы через месяц не искать заново.
   Тесты: 63/63 (+17).
*/
/* v3.5.0: ЗАПАСЫ OZON В БАЛАНСЕ БЫЛИ НЕВЕРНЫ ТРИЖДЫ, И КНИГА ПЕРЕЕХАЛА НА «ПУЛЬТ»
   (сверка блока остатков Ozon с кабинетом, 31.08.2026: «корректные ли данные даёт
   скрипт?» — нет, и ошибка не односторонняя: по ИП1 и ИП6 занижено, по ИП3, ИП5,
   ИП8 завышено, сальдо по группе почти сходилось и прятало все три дефекта).

   ЧТО БОЛЕЛО — три независимые ошибки в «Всего вверено OZON».

   • ОДИН СКЛАД ПОД ДВУМЯ ИМЕНАМИ. `stock_on_warehouses` отдаёт «Санкт_Петербург_РФЦ»,
     `analytics/stocks` — «САНКТ-ПЕТЕРБУРГ_РФЦ»: одно и то же место, разный
     разделитель. Ключ склейки приводил только к верхнему регистру, поэтому склад
     вставал в лист ДВУМЯ строками и считался дважды — 78 шт по ИП1 (Санкт-Петербург,
     Екатеринбург, Ростов-на-Дону). Теперь ключ выбрасывает всё, что не буква и не
     цифра, и приводит Ё к Е.

   • «ГОТОВИМ К ВЫВОЗУ» СЧИТАЛСЯ ДВАЖДЫ. `free_to_sell` уже содержит товар, на
     который создана заявка на вывоз со склада, а аналитика описывает его же полем
     `return_to_seller` — формула складывала обе колонки. Проверено тремя способами:
     внутри листа («Доступно» − «Аналитика» = «Возврат продавцу», 39 строк из 42
     сошлись до штуки), против кабинета (кабинетному «Доступно» равна колонка
     «Аналитика», а не «Доступно») и по смыслу («Возврат продавцу» ровно равен
     кабинетному блоку «Вывоз со склада Ozon» по всем пяти ИП). Основа стала
     max(free_to_sell, available + return_to_seller) — 116 шт двойного счёта прочь.

   • «ДОСТАВЛЯЕМ ПОКУПАТЕЛЯМ» НЕ ПОПАДАЛ В БАЛАНС ВООБЩЕ — самая большая дыра,
     270 шт по пяти кабинетам. Товар уехал к покупателю, но ещё не выкуплен: он наш
     и обязан лежать в запасах, а ни `stock_on_warehouses`, ни `analytics/stocks`
     такого поля не имеют — в кабинете это отдельная колонка отчёта «Управление
     остатками». Считаем сами по `/v2/posting/fbo/list` в статусах `delivering` и
     `awaiting_deliver`; `awaiting_packaging` НЕ берём — этот товар ещё на складе и
     уже сидит в «Зарезервировано» (боевая сверка ИП1: reserved 13 при 13 шт в
     `awaiting_packaging`). Склад берётся из `analytics_data.warehouse_name`, поэтому
     строка склеивается с остатком, а не повисает отдельной.

   ЧЕТВЁРТАЯ, НАЙДЕНА ПОПУТНО: «Проверить низкие остатки» по Ozon молчала всегда.
   Номера колонок были зашиты от раскладки до v3.2.0, и количеством считалась
   колонка с НАЗВАНИЕМ товара: `Number('Витамин Магний…')` = NaN, условие «qty > 0»
   не срабатывало никогда, а молчание выглядело как «остатки в норме». Колонки
   ищутся по заголовку, смотрится только самый свежий день листа.

   ЧТО ЕЩЁ ИЗМЕНИЛОСЬ.
   • Лист остатков Ozon получил колонку «Доставляем покупателям» ПЕРЕД «Кластером»:
     «Всего вверено OZON» осталось последней колонкой самого скрипта, а колонки
     «Себестоимость, ₽» и «Сумма, ₽» справа принадлежат отдельному инструменту.
     Миграция — пункт «⚙️ Обновить настройки таблицы», без боевого прогона.
   • КНИГА ПЕРЕВЕДЕНА НА «ПУЛЬТ» (57.mp-core/ПУЛЬТ.md): в самой книге остался только
     лоадер, все 14 модулей собираются в этот файл и живут в репозитории. Меню
     пронумеровано по порядку работы, появились «📖 Как работать» и «📊 Что сейчас
     происходит» из кода, а «Проверка API» стала «🔌 Проверка связи» — один
     `fetchAll` без ретраев вместо последовательного обхода дюжины кабинетов с
     паузами, который упирался в шестиминутный лимит и выглядел как «кнопка не
     работает». Результат проверки больше не рассылается в Telegram.
   Тесты: 46/46 (tests/run.js — Node на заглушках, FakeSheet проверяет размерность
   setValues, запросы к Ozon сверяются как полезная нагрузка).
*/

// ══════ 00_constants.js ═══════════════════════════════
// ╔══════════════════════════════════════════════════════════╗
// ║  MP Stock Tracker — Константы и конфигурация            ║
// ╚══════════════════════════════════════════════════════════╝

// ─── Служебные листы ─────────────────────────────────────
const SHEET_PANEL = 'Панель управления';
const SHEET_LOG   = 'Лог';
const SHEET_DASH  = 'Дашборд';
const SHEET_MS    = 'Остатки МС';

// ─── Структура листа «Панель управления» ─────────────────
//
//  Строка 1: 🔔 TELEGRAM         (заголовок блока)
//  Строка 2: Токен бота  → B2
//  Строка 3: Chat ID     → B3
//  Строка 4: 🏪 МОЙ СКЛАД       (заголовок блока)
//  Строка 5: API Токен   → B5
//  Строка 6: (пусто)
//  Строка 7: 📦 КАБИНЕТЫ        (заголовок блока)
//  Строка 8: заголовки таблицы  (заморозка)
//  Строка 9+: данные кабинетов
//
const PANEL = {
  TG_TOKEN_CELL:   'B2',
  TG_CHATS_CELL:   'B3',
  MS_TOKEN_CELL:   'B5',
  TABLE_START_ROW: 9,
  COLS: {
    MP: 1, IP: 2, CLIENT_ID: 3, API_KEY: 4, TOKEN: 5,
    SHEET_NAME: 6, ACTIVE: 7, LAST_RUN: 8, STATUS: 9, ROW_COUNT: 10
  }
};

// ─── Модуль: Остатки маркетплейсов ───────────────────────
const LOW_STOCK_THRESHOLD = 5;  // порог «низких остатков», шт.

// v3.4.0. Запас строк, который добавляется листу, когда снимок дня в сетку
// не помещается (см. ensureCapacity_ в 12_utils.gs). ~2000 — это месяцы
// работы даже для самого крупного кабинета, но не раздувает лимит книги.
const SHEET_ROW_HEADROOM = 2000;

// ─── Модуль: OZON — остатки ──────────────────────────────
//
// Две ручки дополняют друг друга: первая даёт свободно/готовим/резерв,
// вторая — товары в пути, возвраты, брак и кластер склада.
//
const OZON_URL_STOCK_ON_WAREHOUSES = 'https://api-seller.ozon.ru/v2/analytics/stock_on_warehouses';
const OZON_URL_ANALYTICS_STOCKS    = 'https://api-seller.ozon.ru/v1/analytics/stocks';
const OZON_URL_PRODUCT_STOCKS      = 'https://api-seller.ozon.ru/v4/product/info/stocks';
const OZON_URL_POSTING_FBO_LIST    = 'https://api-seller.ozon.ru/v2/posting/fbo/list';

const OZON_ANALYTICS_BATCH = 100;   // потолок SKU в одном запросе /v1/analytics/stocks
const OZON_PRODUCT_PAGE    = 1000;  // размер страницы при сборе номенклатуры
const OZON_POSTING_PAGE    = 100;   // потолок страницы у /v2/posting/fbo/list

// v3.5.0. «Доставляем покупателям» — товар уехал к покупателю, но не выкуплен.
// `awaiting_packaging` сюда НЕ входит: он ещё на складе и уже сидит в резерве.
const OZON_DELIVERING_STATUSES = ['delivering', 'awaiting_deliver'];
// Глубина просмотра отправлений: доставка в дальние регионы идёт неделями,
// а фильтр по статусу делает лишние дни бесплатными — отдаются только нужные.
const OZON_FBO_LOOKBACK_DAYS = 180;

const OZON_HEADERS = [
  'Дата и время', 'Артикул', 'SKU', 'Название товара', 'Склад',
  'Доступно к продаже', 'Готовим к продаже', 'Зарезервировано', 'Всего у OZON',
  'Доступно (аналитика)', 'В пути на склад', 'Заявлено к поставке',
  'Возврат от покупателя', 'Возврат продавцу', 'Брак на складе', 'Брак в пути',
  'Ждёт документов', 'Прочее', 'Доставляем покупателям', 'Кластер',
  'Всего вверено OZON'
];

// ─── Модуль: Wildberries — отчёт «Остатки на складах» ────
//
// Старый statistics-api…/v1/supplier/stocks отключён WB 20.07.2026.
// Раздел прав у нового отчёта — «Аналитика», не «Статистика».
//
const WB_REMAINS_URL    = 'https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains';
const WB_REMAINS_PARAMS = '?groupBySa=true&groupByNm=true&groupByBarcode=true' +
                          '&groupBySize=true&filterPics=0&filterVolume=0';

// Квота отчётов у WB считается ПО ЗАПРОСАМ, и опрос статуса её тоже ест.
// При исчерпании WB отвечает 200 с пустым телом (проверено 01.08.2026), поэтому
// опрашиваем редко: отчёт готовится ~8 с, первый же опрос через 15 с его застаёт.
const WB_REMAINS_POLL_MS    = 15000;   // пауза между проверками готовности
const WB_REMAINS_TIMEOUT_MS = 120000;  // потолок ожидания отчёта
const WB_REMAINS_COOLDOWN_MS = 30000;  // пауза перед единственной повторной скачкой

// Псевдосклады в ответе WB — это не склады, а агрегаты
const WB_REMAINS_TOTAL        = 'Всего находится на складах';
const WB_REMAINS_IN_WAY_TO    = 'В пути до получателей';
const WB_REMAINS_IN_WAY_FROM  = 'В пути возвраты на склад WB';

// ─── Модуль: МойСклад ────────────────────────────────────
const MS_API_BASE  = 'https://api.moysklad.ru/api/remap/1.2';
const MS_SHOW_ZERO = false;  // показывать товары с нулевым остатком

const MS_SHEET_HEADERS = [
  'Дата выгрузки', 'Название', 'Артикул', 'Код', 'Склад',
  'Остаток', 'Резерв', 'В пути', 'Доступно',
  'Себестоимость (ср.)', 'Цена продажи', 'Сумма себестоимости', 'Ед. изм.'
];

// ─── Модуль: Кредиты ─────────────────────────────────────
const SHEET_CREDITS       = 'Кредиты Import';
const SHEET_BALANCE       = 'Управленческий баланс';
const BALANCE_TOTAL_LABEL = 'Кредиты банков';

// ─── Модуль: Дебиторка — остатки кабинетов ───────────────
//
// Листы баланса перебираются по порядку: берётся первый, где есть колонка с нужной
// датой. Свежие срезы владелец ведёт в «Авто-тест», поэтому он первый.
//
const BALANCE_SHEETS = ['Управленческий баланс Авто-тест', 'Управленческий баланс'];

const ROW_BALANCE_WB    = 'Остаток на балансе кабинета ВБ';
const ROW_BALANCE_OZON  = 'Остаток на балансе кабинета ОЗОН';

// «Остаток на балансе» — это долг площадки кабинету. Ozon отдаёт его как остаток на
// конец периода в отчёте о взаиморасчётах, WB — отдельной ручкой с категорией «Финансы».
const WB_BALANCE_URL      = 'https://finance-api.wildberries.ru/api/v1/account/balance';
const OZON_URL_CASH_FLOW  = 'https://api-seller.ozon.ru/v1/finance/cash-flow-statement/list';
// Три недели назад — чтобы в ответ гарантированно попал незакрытый период «сегодня».
const OZON_CASHFLOW_LOOKBACK_DAYS = 21;

const CR_HEADERS = {
  PRINCIPAL:   'Основной долг',
  CREDIT_NAME: 'Кредит',
  PAID:        'Оплачено'
};

/**
 * Маппинг: название кредита в «Кредиты Import» → строка в «Управленческий баланс».
 *
 * Структура каждой записи:
 *   credit       — точное название кредита в столбце «Кредит» листа «Кредиты Import»
 *   balanceLabel — точное название строки в столбце A листа «Управленческий баланс»
 *   ip           — произвольная группировка для итогов в отчёте
 *
 * Замените примеры ниже на свои данные.
 */
const CREDIT_MAP = [
  // Субъект 1
  { credit: 'Субъект 1 Кредит A',  balanceLabel: 'Субъект 1 Кредит A',  ip: 'Субъект 1' },
  { credit: 'Субъект 1 Кредит B',  balanceLabel: 'Субъект 1 Кредит B',  ip: 'Субъект 1' },
  // Субъект 2
  { credit: 'Субъект 2 Кредит A',  balanceLabel: 'Субъект 2 Кредит A',  ip: 'Субъект 2' },
];

// ══════ 01_panel.js ═══════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Панель управления — инициализация и миграция
// ════════════════════════════════════════════════════════════

/**
 * Создаёт лист «Панель управления», если он ещё не существует.
 * Вызывается автоматически из onOpen().
 */
function initPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PANEL);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_PANEL, 0);

  // Блок Telegram
  sheet.getRange('A1').setValue('🔔 TELEGRAM').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A2').setValue('Токен бота');
  sheet.getRange('A3').setValue('Chat ID (через ;)');
  sheet.getRange('B2').setValue('').setNote('Вставьте токен бота от @BotFather');
  sheet.getRange('B3').setValue('').setNote('Несколько чатов через точку с запятой: -100123;-100456');

  // Блок МойСклад
  sheet.getRange('A4').setValue('🏪 МОЙ СКЛАД').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A5').setValue('API Токен');
  sheet.getRange('B5').setValue('').setNote('Токен из МойСклад: Настройки → Доступ по API');

  // Заголовок кабинетов
  sheet.getRange('A7').setValue('📦 КАБИНЕТЫ').setFontWeight('bold').setFontSize(11);

  // Заголовки таблицы кабинетов
  const headers = [
    'Маркетплейс', 'ИП', 'Client ID (OZON)', 'API Key (OZON)',
    'Token (WB)', 'Имя листа', 'Активен', 'Последний запуск', 'Статус', 'Строк'
  ];
  const headerRange = sheet.getRange(8, 1, 1, headers.length);
  headerRange.setValues([headers]).setFontWeight('bold').setBackground('#e8eaf6');
  sheet.setFrozenRows(8);

  // Валидация
  const mpRule     = SpreadsheetApp.newDataValidation().requireValueInList(['OZON', 'WB']).build();
  const activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Нет']).build();
  sheet.getRange('A9:A200').setDataValidation(mpRule);
  sheet.getRange('G9:G200').setDataValidation(activeRule);

  // Визуальное выделение автозаполняемых столбцов
  sheet.getRange('H9:J200').setBackground('#f5f5f5').setFontColor('#666666');

  // Ширина столбцов
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(6, 240);
  sheet.setColumnWidth(7, 75);
  sheet.setColumnWidth(8, 155);
  sheet.setColumnWidth(9, 80);
  sheet.setColumnWidth(10, 70);

  return sheet;
}

/**
 * Приводит существующую Панель управления к актуальной структуре.
 * Вставляет строки для блока МойСклад, если их ещё нет.
 * Безопасно вызывать повторно — проверяет текущее состояние перед изменениями.
 */
function patchPanel() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PANEL);
  if (!sheet) { initPanel(); return; }

  const a5 = sheet.getRange('A5').getValue().toString().trim();

  // Если A5 уже содержит «API Токен» — структура верная, ничего не делаем
  if (a5 === 'API Токен') {
    SpreadsheetApp.getUi().alert('✅ Панель управления уже актуальна.');
    return;
  }

  // Вставляем 2 строки после строки 4, сдвигая КАБИНЕТЫ и таблицу вниз
  sheet.insertRowsAfter(4, 2);

  // Заполняем строку 4 (заголовок блока МойСклад)
  sheet.getRange('A4').setValue('🏪 МОЙ СКЛАД').setFontWeight('bold').setFontSize(11);

  // Строка 5 — поле API Токен
  sheet.getRange('A5').setValue('API Токен');
  sheet.getRange('B5').setNote('Токен из МойСклад: Настройки → Доступ по API');

  // Строка 6 остаётся пустой (разделитель перед КАБИНЕТЫ в строке 7)

  SpreadsheetApp.getUi().alert('✅ Панель управления обновлена.\nПоле API Токен МойСклад — ячейка B5.');
}

/**
 * Переносит кабинеты из старых технических листов в Панель управления.
 *
 * Адаптируйте ozonMap и wbMap под структуру ваших старых листов:
 *   ozonMap: для каждого кабинета укажите id, ячейку Client ID и ячейку API Key
 *   wbMap:   для каждого кабинета укажите id и ячейку с токеном
 */
function migrateFromOldConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = initPanel();
  const ui    = SpreadsheetApp.getUi();

  const existingData = panel.getRange(PANEL.TABLE_START_ROW, 1).getValue();
  if (existingData) {
    const confirm = ui.alert('Миграция', 'В таблице уже есть данные. Перезаписать?', ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;
  }

  const rows = [];

  // --- OZON: замените имя листа и ячейки под свою структуру ---
  const ozonSheet = ss.getSheetByName('Технический Озон');
  if (ozonSheet) {
    const ozonMap = [
      // { id: 'Кабинет 1', cidCell: 'B1', keyCell: 'B2' },
      // { id: 'Кабинет 2', cidCell: 'B3', keyCell: 'B4' },
    ];
    ozonMap.forEach(m => {
      const cid = ozonSheet.getRange(m.cidCell).getValue().toString().trim();
      const key = ozonSheet.getRange(m.keyCell).getValue().toString().trim();
      rows.push(['OZON', m.id, cid, key, '', `Остатки OZON ${m.id}`, 'Да', '', '', '']);
    });
  }

  // --- WB: замените имя листа и ячейки под свою структуру ---
  const wbSheet = ss.getSheetByName('Технический ВБ');
  if (wbSheet) {
    const wbMap = [
      // { id: 'Кабинет 1', tCell: 'B1' },
      // { id: 'Кабинет 2', tCell: 'D1' },
    ];
    wbMap.forEach(m => {
      const token = wbSheet.getRange(m.tCell).getValue().toString().trim();
      rows.push(['WB', m.id, '', '', token, `Остатки WB ${m.id}`, 'Да', '', '', '']);
    });
  }

  if (rows.length > 0) {
    panel.getRange(PANEL.TABLE_START_ROW, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Telegram-настройки перенесите вручную
  // panel.getRange(PANEL.TG_TOKEN_CELL).setValue('<ваш токен>');
  // panel.getRange(PANEL.TG_CHATS_CELL).setValue('<ваш chat id>');

  ui.alert(
    `✅ Миграция завершена: ${rows.length} кабинетов перенесено.\n\n` +
    `Заполните Токен бота и Chat ID в Панели управления (B2, B3).\n` +
    `Старые технические листы можно удалить.`
  );
}

// ══════ 02_config.js ══════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Чтение конфигурации из Панели управления
// ════════════════════════════════════════════════════════════

function loadTelegramConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return { token: '', chatIds: [] };

  const token    = panel.getRange(PANEL.TG_TOKEN_CELL).getValue().toString().trim();
  const chatsRaw = panel.getRange(PANEL.TG_CHATS_CELL).getValue().toString().trim();
  const chatIds  = chatsRaw.split(';').map(s => s.trim()).filter(Boolean);

  return { token, chatIds };
}

function getMsToken_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return '';
  return panel.getRange(PANEL.MS_TOKEN_CELL).getValue().toString().trim();
}

function loadCabinets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(SHEET_PANEL);
  if (!panel) return [];

  const lastRow = panel.getLastRow();
  if (lastRow < PANEL.TABLE_START_ROW) return [];

  const numRows = lastRow - PANEL.TABLE_START_ROW + 1;
  const data    = panel.getRange(PANEL.TABLE_START_ROW, 1, numRows, 10).getValues();

  return data
    .map((row, i) => ({
      rowIndex:  PANEL.TABLE_START_ROW + i,
      mp:        String(row[0]).trim().toUpperCase(),
      id:        String(row[1]).trim(),
      clientId:  String(row[2]).trim(),
      apiKey:    String(row[3]).trim(),
      token:     String(row[4]).trim(),
      sheetName: String(row[5]).trim(),
      active:    String(row[6]).trim().toLowerCase() === 'да'
    }))
    .filter(c => c.mp && c.id && c.sheetName);
}

function getActiveCabinets(mp) {
  const all = loadCabinets();
  if (!mp) return all.filter(c => c.active);
  return all.filter(c => c.active && c.mp === mp.toUpperCase());
}

// ══════ 03_telegram.js ════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Telegram — мульти-чат рассылка
// ════════════════════════════════════════════════════════════

/**
 * Отправляет сообщение во все чаты, указанные в Панели управления.
 * Поддерживает Markdown-разметку.
 * @param {string} text — текст сообщения
 */
function sendTelegram(text) {
  const tg = loadTelegramConfig();
  if (!tg.token || tg.chatIds.length === 0) {
    console.warn('Telegram не настроен: проверьте Панель управления (B2, B3)');
    return;
  }

  const url = `https://api.telegram.org/bot${tg.token}/sendMessage`;

  tg.chatIds.forEach(chatId => {
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        payload: { chat_id: chatId, text: text, parse_mode: 'Markdown' },
        muteHttpExceptions: true
      });
    } catch (e) {
      console.error(`Telegram → ${chatId}: ${e.message}`);
    }
  });
}

// ══════ 04_marketplaces.js ════════════════════════════
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

// ══════ 05_ozon.js ════════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: OZON API — сбор остатков по складам
// ════════════════════════════════════════════════════════════
//
// Источников ТРИ, и они дополняют друг друга (сверено с кабинетом 31.08.2026):
//
//   /v2/analytics/stock_on_warehouses — свободно / готовим к продаже / резерв,
//       но показывает ТОЛЬКО склады, где товар физически лежит: ни товаров в пути,
//       ни возвратов, ни брака. По одному SKU это было 22 склада против 37.
//
//   /v1/analytics/stocks — всё остальное, что вверено Ozon: в пути на склад,
//       заявлено к поставке, возвраты от покупателя и продавцу, брак, ожидание
//       документов, плюс кластер склада. Требует список SKU (пачками до 100),
//       поэтому сначала собираем номенклатуру через /v4/product/info/stocks.
//
//   /v2/posting/fbo/list — «Доставляем покупателям» (v3.5.0). Товар уехал к
//       покупателю, но ещё не выкуплен: он наш и обязан лежать в запасах, а НИ
//       ОДНА ручка остатков его не отдаёт — в кабинете это отдельная колонка
//       отчёта «Управление остатками». Считаем сами по FBO-отправлениям.
//
// Раскладка листа: первые 9 колонок оставлены как были (на них могут быть завязаны
// формулы соседних листов), новые данные дописаны справа.
//
// Лист — ИСТОРИЯ, а не снимок «сейчас» (с v3.3.0): каждый прогон дописывает свои
// строки с датой в колонке A, прошлые дни остаются. Повторный прогон в тот же день
// переписывает только сегодняшний блок. Чистка старых дней — «Сервис → Очистка данных».

/**
 * Загружает остатки OZON для одного кабинета и записывает в лист.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} cab — объект кабинета из loadCabinets()
 * @returns {number} количество загруженных строк
 */
function fetchOzon(ss, cab) {
  if (!cab.clientId || !cab.apiKey) throw new Error('Client ID или API Key пустые');

  const now       = new Date();
  const onStock   = fetchOzonOnWarehouses_(cab);   // ключ: sku|склад
  const full      = fetchOzonAnalyticsStocks_(cab); // ключ: sku|склад
  const delivering = fetchOzonDelivering_(cab);     // ключ: sku|склад

  // Объединяем: строка появляется, если она есть хотя бы в одном источнике
  const keys = {};
  [onStock, full, delivering].forEach(src => {
    Object.keys(src).forEach(k => { keys[k] = true; });
  });

  // Артикул и название у «доставляем покупателям» берём по SKU: отправление знает
  // offer_id, но склад в нём может быть тот, где остатка уже нет ни по одной ручке.
  const nameBySku = {};
  Object.keys(keys).forEach(k => {
    const o = onStock[k] || {}, f = full[k] || {};
    const sku = o.sku || f.sku;
    if (sku && !nameBySku[sku] && (o.offerId || f.offerId)) {
      nameBySku[sku] = { offerId: o.offerId || f.offerId, name: o.name || f.name };
    }
  });

  const rows = Object.keys(keys).map(key => {
    const o = onStock[key]    || {};
    const f = full[key]       || {};
    const d = delivering[key] || {};
    const byName = nameBySku[o.sku || f.sku || d.sku] || {};

    const free     = o.free     || 0;
    const promised = o.promised || 0;
    const reserved = o.reserved || 0;

    // «Всего у OZON» (колонка 9) сохраняет прежний смысл, чтобы не сломать формулы
    const legacyTotal = free + promised + reserved;

    // v3.5.0. Основа складского остатка. `free_to_sell` УЖЕ содержит товар, на
    // который создана заявка на вывоз, а аналитика описывает его же отдельным
    // полем `return_to_seller`: сложение давало двойной счёт (31.08.2026 — 116 шт
    // по четырём кабинетам). Тождество «Доступно = Аналитика + Возврат продавцу»
    // проверено построчно, 39 строк из 42 сошлись до штуки.
    const onHand = Math.max(free, (f.available || 0) + (f.returnToSeller || 0));

    // «Всего вверено OZON» — то, что физически у маркетплейса: остатки на складах,
    // товары в пути, возвраты, брак и уехавшее к покупателям. «Заявлено к поставке»
    // сюда НЕ входит: это ещё не отгруженная заявка, её видно отдельной колонкой.
    const entrusted = onHand + promised + reserved +
      (f.transit || 0) + (f.returnFromCustomer || 0) +
      (f.defectStock || 0) + (f.defectTransit || 0) + (f.waitingDocs || 0) +
      (f.other || 0) + (d.qty || 0);

    return [
      now,
      o.offerId || f.offerId || byName.offerId || '',
      o.sku     || f.sku     || d.sku || '',
      o.name    || f.name    || byName.name || '',
      // имя склада из аналитики совпадает с написанием в кабинете — берём его
      f.warehouse || o.warehouse || d.warehouse || '',
      free,
      promised,
      reserved,
      legacyTotal,
      f.available         || 0,
      f.transit           || 0,
      f.requested         || 0,
      f.returnFromCustomer|| 0,
      f.returnToSeller    || 0,
      f.defectStock       || 0,
      f.defectTransit     || 0,
      f.waitingDocs       || 0,
      f.other             || 0,
      d.qty               || 0,
      f.cluster           || '',
      entrusted
    ];
  });

  rows.sort((a, b) => String(a[1]).localeCompare(String(b[1]), 'ru', { numeric: true }));

  // Пустой ответ — это не «остатков нет», а несобранный прогон: в отчёт уходило
  // бодрое «✅ +0», хотя за день в лист не легло ни строки. Показываем ошибкой.
  if (rows.length === 0) {
    throw new Error('OZON вернул 0 строк — в историю за сегодня ничего не записано');
  }

  appendOzonSnapshot_(ss, cab.sheetName, rows, OZON_HEADERS);
  return rows.length;
}

/**
 * Остатки по складам (свободно / готовим / резерв).
 * @returns {Object} карта «sku|склад» → {sku, offerId, name, warehouse, free, promised, reserved}
 */
function fetchOzonOnWarehouses_(cab) {
  const map    = {};
  let   offset = 0;
  const limit  = 1000;

  while (true) {
    const res = fetchWithRetry(OZON_URL_STOCK_ON_WAREHOUSES, ozonOptions_(cab, {
      limit: limit, offset: offset, warehouse_type: 'ALL'
    }));

    const rows = (JSON.parse(res.getContentText()).result || {}).rows || [];
    if (rows.length === 0) break;

    rows.forEach(r => {
      const key = ozonKey_(r.sku, r.warehouse_name);
      const cur = map[key] || {
        sku: String(r.sku || ''), offerId: '', name: '',
        warehouse: r.warehouse_name || '', free: 0, promised: 0, reserved: 0
      };
      cur.offerId  = cur.offerId || r.item_code || '';
      cur.name     = cur.name    || r.item_name || '';
      cur.free     += r.free_to_sell_amount || 0;
      cur.promised += r.promised_amount     || 0;
      cur.reserved += r.reserved_amount     || 0;
      map[key] = cur;
    });

    offset += rows.length;
    if (rows.length < limit) break;
  }

  return map;
}

/**
 * Полная картина остатков: в пути, возвраты, брак, ожидание документов, кластер.
 * @returns {Object} карта «sku|склад» → поля отчёта
 */
function fetchOzonAnalyticsStocks_(cab) {
  const skus = fetchOzonSkus_(cab);
  const map  = {};
  const ADD  = {
    available:          'available_stock_count',
    transit:            'transit_stock_count',
    requested:          'requested_stock_count',
    returnFromCustomer: 'return_from_customer_stock_count',
    returnToSeller:     'return_to_seller_stock_count',
    defectStock:        'stock_defect_stock_count',
    defectTransit:      'transit_defect_stock_count',
    waitingDocs:        'waiting_docs_stock_count',
    other:              'other_stock_count'
  };

  for (let i = 0; i < skus.length; i += OZON_ANALYTICS_BATCH) {
    const batch = skus.slice(i, i + OZON_ANALYTICS_BATCH);
    const res   = fetchWithRetry(OZON_URL_ANALYTICS_STOCKS, ozonOptions_(cab, { skus: batch }));
    const items = JSON.parse(res.getContentText()).items || [];

    items.forEach(it => {
      const key = ozonKey_(it.sku, it.warehouse_name);
      const cur = map[key] || {
        sku: String(it.sku || ''), offerId: '', name: '',
        warehouse: it.warehouse_name || '', cluster: ''
      };
      cur.offerId   = cur.offerId || it.offer_id || '';
      cur.name      = cur.name    || it.name     || '';
      cur.warehouse = it.warehouse_name || cur.warehouse;
      cur.cluster   = cur.cluster || it.cluster_name || '';
      Object.keys(ADD).forEach(k => { cur[k] = (cur[k] || 0) + (it[ADD[k]] || 0); });
      map[key] = cur;
    });
  }

  return map;
}

/**
 * v3.5.0. «Доставляем покупателям» — товар уехал к покупателю, но ещё не выкуплен.
 *
 * Ни `stock_on_warehouses`, ни `analytics/stocks` этой цифры не отдают: в кабинете
 * она живёт отдельной колонкой отчёта «Управление остатками». Сверка 31.08.2026
 * показала, что без неё баланс недосчитывался 270 шт по пяти кабинетам — самая
 * крупная из трёх найденных дыр.
 *
 * Берутся отправления FBO в двух статусах:
 *   delivering       — едет к покупателю;
 *   awaiting_deliver — собрано и ждёт отгрузки со склада Ozon.
 * Статус `awaiting_packaging` НЕ берётся: этот товар ещё лежит на складе и уже
 * посчитан в «Зарезервировано» (боевая сверка ИП1 31.08: reserved 13 при 13 шт
 * в `awaiting_packaging`). Сложение дало бы двойной счёт.
 *
 * Склад отправления берётся из `analytics_data.warehouse_name` — написание там
 * совпадает со складскими ручками, поэтому строка склеивается с остатком, а не
 * повисает отдельной.
 *
 * @returns {Object} карта «sku|склад» → {sku, warehouse, qty}
 */
function fetchOzonDelivering_(cab) {
  const map  = {};
  const now  = new Date();
  const from = new Date(now.getTime() - OZON_FBO_LOOKBACK_DAYS * 86400000);
  const to   = new Date(now.getTime() + 86400000);

  OZON_DELIVERING_STATUSES.forEach(status => {
    let offset = 0;
    while (true) {
      const res = fetchWithRetry(OZON_URL_POSTING_FBO_LIST, ozonOptions_(cab, {
        dir: 'DESC', limit: OZON_POSTING_PAGE, offset: offset,
        filter: { since: from.toISOString(), to: to.toISOString(), status: status },
        with: { analytics_data: true, financial_data: false }
      }));
      const list = JSON.parse(res.getContentText()).result || [];

      list.forEach(p => {
        const wh = ((p.analytics_data || {}).warehouse_name) || '';
        (p.products || []).forEach(pr => {
          const key = ozonKey_(pr.sku, wh);
          const cur = map[key] || { sku: String(pr.sku || ''), warehouse: wh, qty: 0 };
          cur.qty += pr.quantity || 0;
          map[key] = cur;
        });
      });

      if (list.length < OZON_POSTING_PAGE) break;
      offset += list.length;
    }
  });

  return map;
}

/** Собирает все SKU кабинета — отчёт по остаткам требует их списком. */
function fetchOzonSkus_(cab) {
  const seen   = {};
  const skus   = [];
  let   cursor = '';

  while (true) {
    const res  = fetchWithRetry(OZON_URL_PRODUCT_STOCKS, ozonOptions_(cab, {
      cursor: cursor, limit: OZON_PRODUCT_PAGE, filter: {}
    }));
    const body  = JSON.parse(res.getContentText());
    const items = body.items || [];

    items.forEach(it => {
      (it.stocks || []).forEach(s => {
        if (s.sku && !seen[s.sku]) { seen[s.sku] = true; skus.push(s.sku); }
      });
    });

    // Ozon отдаёт непустой курсор даже на последней (неполной) странице,
    // поэтому останавливаемся по числу строк, а не только по курсору.
    cursor = body.cursor || '';
    if (items.length < OZON_PRODUCT_PAGE || !cursor) break;
  }

  return skus;
}

/** Общие параметры запроса к Ozon Seller API. */
function ozonOptions_(cab, payload) {
  return {
    method:      'post',
    contentType: 'application/json',
    headers:     { 'client-id': cab.clientId, 'api-key': cab.apiKey },
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  };
}

/**
 * v3.5.0. Ключ склейки источников.
 *
 * Названия складов у ручек различаются не только регистром, но и разделителем:
 * `stock_on_warehouses` отдал «Санкт_Петербург_РФЦ», `analytics/stocks` — тот же
 * склад как «САНКТ-ПЕТЕРБУРГ_РФЦ». Прежний ключ приводил только к верхнему
 * регистру, поэтому один склад вставал в лист ДВУМЯ строками и считался дважды
 * (31.08.2026 — 78 шт по ИП1: Санкт-Петербург, Екатеринбург, Казань).
 * Нормализуем жёстко: регистр, Ё→Е и прочь всё, что не буква и не цифра.
 */
function ozonKey_(sku, warehouseName) {
  const wh = String(warehouseName || '')
    .toUpperCase()
    .replace(/Ё/g, 'Е')
    .replace(/[^0-9A-ZА-Я]+/g, '');
  return String(sku) + '|' + wh;
}

/**
 * v3.3.0. Дописывает снимок дня в лист OZON, СОХРАНЯЯ прошлые дни.
 *
 * До 3.3.0 здесь стоял `sheet.clear()`: каждый прогон стирал всё, что собрали
 * раньше, и остатки жили ровно до следующего запуска — истории по Ozon не было
 * вообще, тогда как листы WB её копили. Теперь поведение одинаковое: строки
 * накапливаются, дата прогона — в колонке A.
 *
 * Повторный запуск в тот же день не плодит дубли: строки за сегодняшнее число
 * сначала удаляются, потом пишется свежий снимок. То есть «последний прогон дня
 * побеждает», а прошлые дни неприкосновенны.
 */
function appendOzonSnapshot_(ss, name, rows, headers) {
  if (rows.length === 0) return;

  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    writeOzonHeaders_(sheet, headers);
  } else if (sheet.getLastRow() === 0) {
    writeOzonHeaders_(sheet, headers);
  } else if (!ozonHeaderMatches_(sheet, headers)) {
    // Лист остался от старой раскладки (у ИП4 это англоязычная шапка с датой
    // в колонке I). Дописывать в него 21 колонку нельзя — данные встанут криво.
    // Старое содержимое сохраняем отдельной копией и начинаем лист заново.
    archiveSheetCopy_(ss, sheet);
    sheet.clear();
    writeOzonHeaders_(sheet, headers);
  } else {
    deleteRowsOfDay_(sheet, rows[0][0]);
  }

  const startRow = sheet.getLastRow() + 1;
  ensureCapacity_(sheet, startRow, rows.length);
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

/** Шапка листа остатков: жирная и закреплённая — истории будет много. */
function writeOzonHeaders_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * v3.3.0. Совпадает ли шапка листа с ожидаемой.
 * Сравниваем только первые `headers.length` колонок: справа у владельца могут
 * быть свои колонки с формулами (на листах WB так и есть), они не мешают.
 */
function ozonHeaderMatches_(sheet, headers) {
  if (sheet.getLastColumn() < headers.length) return false;
  const actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  return headers.every((h, i) => String(actual[i]).trim() === h);
}

// ══════ 06_wb.js ══════════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Wildberries API — сбор остатков
// ════════════════════════════════════════════════════════════
//
// ВНИМАНИЕ (01.08.2026): старый метод statistics-api…/v1/supplier/stocks
// ОТКЛЮЧЁН маркетплейсом 20.07.2026 (заглушка PLUG-404-20260720,
// dev.wildberries.ru/release-notes?id=494). Живой токен получает от него 404.
//
// Замена — асинхронный отчёт «Остатки на складах» из раздела «Аналитика»:
//   1) GET  …/api/v1/warehouse_remains?…       → data.taskId
//   2) GET  …/warehouse_remains/tasks/{id}/status  → data.status = done
//   3) GET  …/warehouse_remains/tasks/{id}/download → массив строк
//
// Токену нужно право «Аналитика» (у старых ключей стояла только «Статистика» —
// такой ключ отвечает 401 «token scope not allowed» и требует перевыпуска в ЛК).
//
// Чего в новом отчёте НЕТ по сравнению со старым: Цена и Скидка — эти колонки
// остаются пустыми, чтобы не ломать раскладку листов с историей.

/**
 * Загружает остатки WB для одного кабинета и записывает в лист.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} cab — объект кабинета из loadCabinets()
 * @returns {number} количество загруженных строк
 */
function fetchWB(ss, cab) {
  if (!cab.token) throw new Error('Token пустой');

  const items = fetchWbRemains_(cab.token);
  const now   = new Date();

  const headers = [
    'Дата и время', 'Склад', 'Артикул продавца', 'Артикул WB', 'Баркод',
    'Количество', 'В пути к клиенту', 'В пути от клиента', 'Полное количество', 'Цена', 'Скидка'
  ];

  const rows = [];
  items.forEach(item => {
    (item.warehouses || []).forEach(w => {
      const name = String(w.warehouseName || '');
      const qty  = w.quantity || 0;

      // «Всего находится на складах» — производная сумма, её не пишем,
      // иначе любой SUM по листу удвоится.
      if (name === WB_REMAINS_TOTAL) return;

      const toClient   = name === WB_REMAINS_IN_WAY_TO   ? qty : '';
      const fromClient = name === WB_REMAINS_IN_WAY_FROM ? qty : '';
      const onStock    = (toClient === '' && fromClient === '') ? qty : 0;

      rows.push([
        now,
        name,
        item.vendorCode || '',
        item.nmId       || '',
        item.barcode    || '',
        onStock,
        toClient,
        fromClient,
        qty,
        '',   // Цена — нового отчёта не даёт
        ''    // Скидка — нового отчёта не даёт
      ]);
    });
  });

  rows.sort((a, b) => String(a[2]).localeCompare(String(b[2]), 'ru', { numeric: true }));

  writeToSheet(ss, cab.sheetName, rows, headers);
  return rows.length;
}

/**
 * Заказывает, дожидается и скачивает отчёт «Остатки на складах».
 * @param {string} token — токен кабинета WB (право «Аналитика»)
 * @returns {Array<Object>} строки отчёта
 */
function fetchWbRemains_(token) {
  const opts = { headers: { Authorization: token }, muteHttpExceptions: true };

  const created = wbJson_(fetchWithRetry(WB_REMAINS_URL + WB_REMAINS_PARAMS, opts));
  const taskId  = created.data && created.data.taskId;
  if (!taskId) throw new Error('WB не вернул taskId');

  const deadline = Date.now() + WB_REMAINS_TIMEOUT_MS;
  while (Date.now() < deadline) {
    Utilities.sleep(WB_REMAINS_POLL_MS);
    const status = wbJson_(fetchWithRetry(`${WB_REMAINS_URL}/tasks/${taskId}/status`, opts)).data.status;
    if (status === 'done')   break;
    if (status === 'purged') throw new Error('Отчёт WB удалён на стороне маркетплейса');
  }

  // Кабинет без остатков: WB отдаёт HTTP 204 без тела (проверено на ИП4 03.08.2026).
  // Это штатный ответ «строк нет», а не сбой — переспрашивать нечего.
  const dump = fetchWithRetry(`${WB_REMAINS_URL}/tasks/${taskId}/download`, opts);
  if (dump.getResponseCode() === 204) return [];

  // Пустое тело при 200 — либо кончилась квота отчётов, либо тот же пустой отчёт.
  // Различаем повтором: отчёт уже готов, ждать его заново не нужно.
  let body = wbBody_(dump);
  if (!body) {
    Utilities.sleep(WB_REMAINS_COOLDOWN_MS);
    body = wbBody_(fetchWithRetry(`${WB_REMAINS_URL}/tasks/${taskId}/download`, opts));
  }
  if (!body) return [];

  const data = JSON.parse(body);
  if (!Array.isArray(data)) throw new Error('Отчёт WB вернул не список строк');
  return data;
}

/** Тело ответа без пробелов; пустая строка, если тела нет. */
function wbBody_(res) {
  return String(res.getContentText() || '').trim();
}

/**
 * Разбирает ответ WB, отдельно ловя пустое тело.
 * Проверено 01.08.2026: WB отвечает HTTP 200 с ПУСТЫМ телом и на исчерпанной
 * квоте отчётов, и на пустом отчёте — голый JSON.parse падает на
 * «Unexpected end of input», и причина теряется.
 */
function wbJson_(res) {
  const body = wbBody_(res);
  if (!body) {
    const left = res.getHeaders()['X-Ratelimit-Remaining'];
    throw new Error('WB вернул пустой ответ (лимитер отчётов' +
      (left !== undefined ? `, остаток квоты: ${left}` : '') + ') — повторить позже');
  }
  return JSON.parse(body);
}

// ══════ 07_moysklad.js ════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: МойСклад — остатки с себестоимостью
// ════════════════════════════════════════════════════════════

/**
 * Загружает остатки из МойСклад и дописывает строки в лист «Остатки МС».
 * Токен читается из Панели управления (B5).
 */
function fetchMsStock() {
  const sheet     = getOrCreateSheet_(SHEET_MS);
  const timestamp = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd HH:mm');

  if (sheet.getLastRow() === 0) writeMsHeaders_(sheet);

  const productMap = loadMsProductDetails_();
  Logger.log(`МойСклад: загружено ${Object.keys(productMap).length} товаров`);

  const allRows = [];
  let offset    = 0;
  const limit   = 1000;
  let total     = Infinity;

  while (offset < total) {
    const data = fetchMsPage_(
      `${MS_API_BASE}/report/stock/bystore?limit=${limit}&offset=${offset}`
    );

    if (!data || !data.rows) {
      try {
        SpreadsheetApp.getUi().alert(
          'Ошибка загрузки МойСклад. Проверьте API Токен в Панели управления (B5).'
        );
      } catch (e) {
        Logger.log('Ошибка загрузки МойСклад. Проверьте API Токен.');
      }
      return;
    }

    total = data.meta.size;

    for (const item of data.rows) {
      const productHref = item.meta ? item.meta.href.split('?')[0] : '';
      const product     = productMap[productHref] || {};

      if (!item.stockByStore || item.stockByStore.length === 0) continue;

      for (const storeData of item.stockByStore) {
        const stock     = storeData.stock     || 0;
        const reserve   = storeData.reserve   || 0;
        const inTransit = storeData.inTransit || 0;
        const available = stock - reserve;

        if (!MS_SHOW_ZERO && stock === 0 && inTransit === 0) continue;

        const costPrice = product.price    || 0;
        const salePrice = product.salePrice || 0;

        allRows.push([
          timestamp,
          product.name    || '',
          product.article || '',
          product.code    || '',
          storeData.name  || 'Неизвестный склад',
          stock, reserve, inTransit, available,
          costPrice, salePrice,
          costPrice * stock,
          product.uom || ''
        ]);
      }
    }

    offset += limit;
    if (offset < total) Utilities.sleep(300);
  }

  if (allRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    ensureCapacity_(sheet, startRow, allRows.length);
    sheet.getRange(startRow, 1, allRows.length, MS_SHEET_HEADERS.length).setValues(allRows);
    sheet.getRange(startRow, 10, allRows.length, 3).setNumberFormat('#,##0.00');
  }

  const msg = `МойСклад: добавлено ${allRows.length} строк (${timestamp})`;
  Logger.log(msg);
  sendTelegram(`✅ *${msg}*`);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'МойСклад — Остатки', 5);
}

// ─── Вспомогательные функции ─────────────────────────────

/**
 * Загружает детали товаров (name, article, code, price, uom) из report/stock/all.
 * @returns {Object} карта href → данные товара
 */
function loadMsProductDetails_() {
  const map   = {};
  let offset  = 0;
  const limit = 1000;
  let total   = Infinity;

  while (offset < total) {
    const data = fetchMsPage_(
      `${MS_API_BASE}/report/stock/all?limit=${limit}&offset=${offset}&stockMode=all`
    );
    if (!data || !data.rows) break;

    total = data.meta.size;

    for (const item of data.rows) {
      const href = item.meta ? item.meta.href.split('?')[0] : '';
      if (!href) continue;

      map[href] = {
        name:      item.name    || '',
        article:   item.article || '',
        code:      item.code    || '',
        price:     (item.price     || 0) / 100,
        salePrice: (item.salePrice || 0) / 100,
        uom:       item.uom ? (item.uom.name || '') : ''
      };
    }

    offset += limit;
    if (offset < total) Utilities.sleep(300);
  }

  return map;
}

function fetchMsPage_(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method:             'get',
      headers:            getMsAuthHeaders_(),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log(
        `МойСклад API ошибка (${response.getResponseCode()}): ` +
        response.getContentText().substring(0, 500)
      );
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log(`МойСклад запрос ошибка: ${e.message}`);
    return null;
  }
}

function getMsAuthHeaders_() {
  const token = getMsToken_();
  if (!token) throw new Error(
    'Токен МойСклад не указан! Заполните поле «API Токен» в Панели управления (B5).'
  );
  return {
    'Authorization':  `Bearer ${token}`,
    'Accept-Encoding': 'gzip'
  };
}

function writeMsHeaders_(sheet) {
  const range = sheet.getRange(1, 1, 1, MS_SHEET_HEADERS.length);
  range.setValues([MS_SHEET_HEADERS]);
  range.setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

// ─── Триггеры МойСклад ───────────────────────────────────

function createMsDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fetchMsStock')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('fetchMsStock')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Остатки МойСклад будут загружаться каждый день в ~7:00.',
    'Автозапуск МС настроен', 5
  );
}

function removeMsDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fetchMsStock')
    .forEach(t => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Триггер МойСклад удалён.',
    'Автозапуск МС отключён', 5
  );
}

// ══════ 08_credits.js ═════════════════════════════════
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

// ══════ 09_dashboard.js ═══════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Дашборд и экспорт
// ════════════════════════════════════════════════════════════

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DASH);
  if (!sheet) sheet = ss.insertSheet(SHEET_DASH);

  sheet.clear();

  const headers = ['Маркетплейс', 'ИП', 'Лист', 'Строк данных', 'Последнее обновление', 'Статус'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8eaf6');
  sheet.setFrozenRows(1);

  const cabinets = loadCabinets();
  if (cabinets.length === 0) return;

  const rows = cabinets.map(c => {
    const dataSheet = ss.getSheetByName(c.sheetName);
    if (!dataSheet) return [c.mp, c.id, c.sheetName, 0, '—', '⚠️ Нет листа'];

    const lastRow  = dataSheet.getLastRow();
    const rowCount = Math.max(0, lastRow - 1);

    let lastUpdate = '—';
    if (lastRow > 1) {
      const val = dataSheet.getRange(lastRow, 1).getValue();
      if (val instanceof Date) {
        lastUpdate = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
      } else if (typeof val === 'string' && val) {
        lastUpdate = val;
      }
    }

    const status = !c.active ? '⏸ Неактивен' : (rowCount > 0 ? '✅' : '⚠️ Пусто');
    return [c.mp, c.id, c.sheetName, rowCount, lastUpdate, status];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const totalRow = rows.length + 3;
  sheet.getRange(totalRow, 1).setValue('📊 ИТОГО').setFontWeight('bold');
  sheet.getRange(totalRow, 4).setFormula(`=SUM(D2:D${rows.length + 1})`).setFontWeight('bold');
  sheet.getRange(totalRow + 1, 1).setValue('🕐 Обновлено').setFontColor('#999999');
  sheet.getRange(totalRow + 1, 2).setValue(new Date()).setFontColor('#999999');

  sheet.autoResizeColumns(1, headers.length);
}

function exportDashboardCSV() {
  updateDashboard();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DASH);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const csv  = data.map(row => row.map(cell => {
    const val = cell instanceof Date
      ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : String(cell);
    return `"${val.replace(/"/g, '""')}"`;
  }).join(',')).join('\n');

  const filename = 'stock_dashboard_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
  const file = DriveApp.createFile(filename, csv, 'text/csv');

  SpreadsheetApp.getUi().alert(`📁 CSV сохранён на Google Drive:\n${file.getUrl()}`);
}

// ══════ 10_alerts.js ══════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Оповещения — низкие остатки и проверка API
// ════════════════════════════════════════════════════════════

/**
 * Проверяет все активные кабинеты и отправляет Telegram-уведомление
 * о позициях с остатком ≤ LOW_STOCK_THRESHOLD.
 *
 * v3.5.0. Колонки ищутся ПО ЗАГОЛОВКУ, а не по номеру. До этого номера были
 * зашиты от старой раскладки листа Ozon: количеством считалась колонка 4, где
 * с v3.2.0 лежит название товара. `Number('Витамин Магний…')` = NaN, условие
 * `qty > 0` не срабатывало никогда — по Ozon оповещение молчало вообще, и
 * молчание выглядело как «остатки в норме».
 *
 * Смотрим только САМЫЙ СВЕЖИЙ день листа: лист — история, и позиция, которая
 * заканчивалась неделю назад, к сегодняшнему дню давно приехала.
 */
function checkLowStock() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cabinets = getActiveCabinets();
  const alerts   = [];

  const QTY  = { OZON: ['Всего вверено OZON', 'Всего у OZON'], WB: ['Количество'] };
  const NAME = { OZON: ['Название товара', 'Артикул'], WB: ['Артикул продавца'] };
  const WH   = { OZON: ['Склад'], WB: ['Склад'] };

  cabinets.forEach(cab => {
    const sheet = ss.getSheetByName(cab.sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    if (!QTY[cab.mp]) return;

    const data = sheet.getDataRange().getValues();
    const head = data[0].map(h => String(h).trim());
    const pick = names => {
      for (let i = 0; i < names.length; i++) {
        const j = head.indexOf(names[i]);
        if (j >= 0) return j;
      }
      return -1;
    };
    const qi = pick(QTY[cab.mp]), ni = pick(NAME[cab.mp]), wi = pick(WH[cab.mp]);
    if (qi < 0 || ni < 0) return;   // чужая раскладка — молчим, а не врём цифрами

    const today = dayKey_(data[data.length - 1][0]);
    for (let i = data.length - 1; i >= 1; i--) {
      if (dayKey_(data[i][0]) !== today) break;   // дошли до прошлого дня
      const qty = Number(data[i][qi]) || 0;
      if (qty > 0 && qty <= LOW_STOCK_THRESHOLD) {
        alerts.push(`⚠️ ${cab.mp} ${cab.id}: "${data[i][ni]}" — ${qty} шт.` +
                    (wi >= 0 ? ` (${data[i][wi]})` : ''));
      }
    }
  });

  if (alerts.length > 0) {
    const msg = [`🔔 *Низкие остатки (≤${LOW_STOCK_THRESHOLD} шт.)*`, ''];
    msg.push(...alerts.slice(0, 50));
    if (alerts.length > 50) msg.push(`\n...и ещё ${alerts.length - 50}`);
    sendTelegram(msg.join('\n'));
  } else {
    SpreadsheetApp.getUi().alert(
      `Все остатки в норме — позиций с количеством ≤ ${LOW_STOCK_THRESHOLD} нет.`
    );
  }

  return alerts.length;
}

/**
 * Старое имя пункта «Проверка API» — на нём мог остаться чужой пункт меню.
 *
 * v3.5.0. Сама проверка переехала в `panelCheckConnection` (13_pult.js) и стала
 * одним `fetchAll` без ретраев: последовательный обход дюжины кабинетов с
 * паузами упирался в шестиминутный лимит Apps Script и выглядел как «кнопка не
 * работает». Заодно ушла рассылка результата в Telegram: диагностику смотрит
 * тот, кто её нажал, а чат она засоряла.
 */
function healthCheck() {
  return panelCheckConnection();
}

// ══════ 11_triggers.js ════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Триггеры — управление автозапуском
// ════════════════════════════════════════════════════════════

/**
 * Диалог настройки периодического обновления остатков маркетплейсов.
 * Триггер МойСклад настраивается отдельно из меню «Остатки МС».
 */
function setupTriggers() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '⏰ Настройка автозапуска',
    'Как часто обновлять остатки маркетплейсов (в часах)?\nНапример: 1, 2, 4, 6, 12',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const hours = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(hours) || hours < 1 || hours > 24) {
    ui.alert('Укажите число от 1 до 24');
    return;
  }

  // Удаляем старый триггер маркетплейсов
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runAllSupplyFunctions')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runAllSupplyFunctions')
    .timeBased()
    .everyHours(hours)
    .create();

  ui.alert(
    `✅ Триггер установлен:\n` +
    `• Обновление остатков маркетплейсов — каждые ${hours} ч.\n\n` +
    `Триггер для МойСклад настраивается отдельно через «Остатки МС → Включить ежедневный автозапуск».`
  );
}

/** Удаляет ВСЕ триггеры проекта. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Все триггеры удалены.');
}

// ══════ 12_utils.js ═══════════════════════════════════
// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Утилиты — HTTP, запись в листы, форматирование
// ════════════════════════════════════════════════════════════

// ─── HTTP ────────────────────────────────────────────────

/**
 * Выполняет HTTP-запрос с повторными попытками.
 *
 * Ретраятся только те ошибки, которые лечатся повтором: сеть, 5xx и 429.
 * 4xx возвращается сразу — повтор «метода нет» или «токен протух» ничего
 * не чинит, а на 8 кабинетах съедает минуты из лимита GAS в 6 минут.
 * Текст ошибки маркетплейса берётся ИЗ ТЕЛА ответа: у WB там лежит причина
 * («token scope not allowed», «This method is deprecated» со ссылкой), и
 * без неё в отчёт уходит бесполезное «HTTP 404».
 *
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
      const res  = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      // Успех — любой 2xx. WB отвечает 204 на пустой отчёт: раньше это уходило
      // в ретраи и падало «Адрес недоступен (после 3 попыток)».
      if (code >= 200 && code < 300) return res;

      lastError = new Error(`HTTP ${code}: ${extractApiError_(res)}`);

      // 4xx, кроме 429, повтором не лечится
      if (code >= 400 && code < 500 && code !== 429) throw lastError;

      if (code === 429) {
        const wait = Number(res.getHeaders()['X-Ratelimit-Retry'] || 0);
        if (wait > 0) Utilities.sleep(Math.min(wait, 60) * 1000);
      }
    } catch (e) {
      if (e === lastError) throw e;   // осознанный отказ, а не сбой сети
      lastError = e;
    }
    if (attempt < maxRetries) Utilities.sleep(2000 * attempt);
  }

  throw new Error(`${lastError.message} (после ${maxRetries} попыток)`);
}

/** Достаёт из тела ответа причину отказа, обрезая до читаемой длины. */
function extractApiError_(res) {
  const body = String(res.getContentText() || '').trim();
  try {
    const j = JSON.parse(body);
    const parts = [j.title, j.detail, j.message, j.error, j.errorMessage]
      .filter(Boolean).map(String);
    if (parts.length) return parts.join(' — ').slice(0, 300);
  } catch (e) { /* тело не JSON — отдаём как есть */ }
  return body.slice(0, 300) || '(пустое тело ответа)';
}

// ─── Работа с листами ────────────────────────────────────

/**
 * v3.4.0. Досоздаёт строки внизу листа, если снимок в них не помещается.
 *
 * Лист остатков растёт каждый день, а сетка Google-таблицы — нет. Когда
 * свободные строки кончаются, `setValues` за границей сетки падает, и день
 * просто не загружается (на «Остатки ВБ ИП11» так потерялись 04 и 05.08.2026:
 * сетка 7725 строк была занята вся, снимок 03.08 влез частично).
 *
 * Добавляем не впритык, а с запасом `SHEET_ROW_HEADROOM` — чтобы не дёргать
 * `insertRowsAfter` каждый прогон.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow — первая строка, куда будем писать
 * @param {number} needRows — сколько строк пишем
 */
function ensureCapacity_(sheet, startRow, needRows) {
  const required = startRow + needRows - 1;
  const maxRows  = sheet.getMaxRows();
  if (required <= maxRows) return;

  sheet.insertRowsAfter(maxRows, required - maxRows + SHEET_ROW_HEADROOM);
}

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

  const lastRow  = sheet.getLastRow();
  const startRow = lastRow === 0 ? 2 : lastRow + 1;
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  ensureCapacity_(sheet, startRow, rows.length);
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
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
 * v3.3.0. Удаляет из листа строки, у которых дата в колонке A — тот же день,
 * что и `when`. Нужно, чтобы повторный прогон переписывал сегодняшний снимок,
 * а не дублировал его поверх вчерашнего.
 *
 * Строки за один день лежат подряд (их пишет один прогон), но на всякий случай
 * удаляем блоками снизу вверх — так номера строк выше не съезжают.
 *
 * @returns {number} сколько строк удалено
 */
function deleteRowsOfDay_(sheet, when) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;

  const key = dayKey_(when);
  if (!key) return 0;   // дату не распознали — лучше ничего не трогать

  const values  = sheet.getRange(2, 1, last - 1, 1).getValues();
  let   deleted = 0;
  let   runEnd  = -1;   // индекс нижней строки непрерывного блока за этот день

  for (let i = values.length - 1; i >= -1; i--) {
    const hit = i >= 0 && dayKey_(values[i][0]) === key;
    if (hit && runEnd < 0) runEnd = i;
    if (!hit && runEnd >= 0) {
      const from  = i + 1;
      const count = runEnd - from + 1;
      sheet.deleteRows(from + 2, count);   // +2: строка 1 — шапка, индексы с нуля
      deleted += count;
      runEnd = -1;
    }
  }

  return deleted;
}

/**
 * v3.3.0. День как `yyyy-MM-dd` — из Date, из «dd.MM.yyyy …» или из ISO-строки.
 * Пустая строка означает «дата не распознана».
 */
function dayKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();

  const ru = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (ru) return `${ru[3]}-${ru[2]}-${ru[1]}`;

  const iso = s.match(/^\d{4}-\d{2}-\d{2}/);
  return iso ? iso[0] : '';
}

/**
 * v3.3.0. Откладывает копию листа перед тем, как переписать его новой раскладкой.
 * Имя копии уникализируется — за день лист могут пересобрать не один раз.
 */
function archiveSheetCopy_(ss, sheet) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const base  = `${sheet.getName()} (старый формат ${stamp})`;

  let name = base;
  for (let n = 2; ss.getSheetByName(name); n++) name = `${base} #${n}`;

  return sheet.copyTo(ss).setName(name);
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

    const dateCol = 1;
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

// ══════ 13_pult.js ════════════════════════════════════
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
    ['3️⃣ Остатки кабинетов в дебиторку',
     'Спрашивает дату колонки баланса, собирает по всем кабинетам долг площадки ' +
     '(WB — «Баланс» кабинета, Ozon — остаток отчёта о взаиморасчётах), показывает ' +
     'сухой прогон и после подтверждения пишет строки «Остаток на балансе кабинета ' +
     'ВБ/ОЗОН». Кабинет, чей ключ отказал, пропускается — ячейка остаётся как была, ' +
     'а не обнуляется. Осмысленна только СЕГОДНЯШНЯЯ дата: истории по остатку ' +
     'кабинета площадки не отдают.'],
    ['4️⃣ Кредиты в баланс',
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

  const rows = [
    ['Остаток на балансе кабинета ВБ / ОЗОН', 'долг площадки кабинету. Заполняет пункт 3️⃣ по API.'],
    ['Ожидаем поступление на РС от ВБ / ОЗОН', 'деньги, которые площадка уже отправила, но которые ещё не дошли до счёта. Публичного метода для них нет НИ У ОДНОЙ площадки — ставит человек по личному кабинету.'],
    ['Дебиторская задолженность, итоги по ИП', 'формулы, считаются сами.']
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
    '<h2>Строки дебиторки в балансе</h2>',
    '<table><tr><th>Строка</th><th>Откуда берётся</th></tr>',
    rows.map(r => '<tr><td>' + pultEsc_(r[0]) + '</td><td>' + pultEsc_(r[1]) + '</td></tr>').join(''),
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
    '3) Остатки кабинетов в дебиторку — долг WB и Ozon кабинету на СЕГОДНЯ;',
    '   кабинет с отказавшим ключом пропускается, ячейка не обнуляется.',
    '   «Ожидаем поступление на РС» скрипт не заполняет: метода нет ни у одной',
    '   площадки, эти строки ставит человек по личному кабинету.',
    '4) Кредиты в баланс — из листа «Кредиты Import» в строки баланса.',
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

// ══════ 14_receivables.js ═════════════════════════════
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

