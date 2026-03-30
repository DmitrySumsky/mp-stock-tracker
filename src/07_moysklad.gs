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
