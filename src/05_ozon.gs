// ════════════════════════════════════════════════════════════
// МОДУЛЬ: OZON API — сбор остатков по складам
// ════════════════════════════════════════════════════════════

/**
 * Загружает остатки OZON для одного кабинета и записывает в лист.
 *
 * Поля из API /v2/analytics/stock_on_warehouses:
 *   free_to_sell_amount  — доступно к продаже на складах OZON
 *   promised_amount      — готовим к продаже (приняты, проходят обработку)
 *   reserved_amount      — зарезервировано: в доставке к покупателям +
 *                          возвраты в пути + перемещения между складами
 *
 * Итого у OZON = free + promised + reserved.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} cab — объект кабинета из loadCabinets()
 * @returns {number} количество загруженных строк
 */
function fetchOzon(ss, cab) {
  if (!cab.clientId || !cab.apiKey) throw new Error('Client ID или API Key пустые');

  const url  = 'https://api-seller.ozon.ru/v2/analytics/stock_on_warehouses';
  const now  = new Date();
  const headers = [
    'Дата и время', 'Артикул', 'SKU', 'Название товара', 'Склад',
    'Доступно к продаже', 'Готовим к продаже', 'Зарезервировано', 'Всего у OZON'
  ];

  let allRows = [];
  let offset  = 0;
  const limit = 1000;

  while (true) {
    const res = fetchWithRetry(url, {
      method:      'post',
      contentType: 'application/json',
      headers:     { 'client-id': cab.clientId, 'api-key': cab.apiKey },
      payload:     JSON.stringify({ limit, offset, warehouse_type: 'ALL' }),
      muteHttpExceptions: true
    });

    const data = JSON.parse(res.getContentText()).result.rows;
    if (!data || data.length === 0) break;

    const mapped = data.map(r => {
      const free     = r.free_to_sell_amount || 0;
      const promised = r.promised_amount     || 0;
      const reserved = r.reserved_amount     || 0;
      return [
        now,
        r.item_code     || '',
        r.sku           || '',
        r.item_name     || '',
        r.warehouse_name || '',
        free,
        promised,
        reserved,
        free + promised + reserved
      ];
    });

    allRows = allRows.concat(mapped);
    offset += data.length;
    if (data.length < limit) break;
  }

  // Лист перезаписывается снапшотом, поэтому пустой ответ опасен: writeOzonSheet_
  // на нуле строк выходит до clearContents(), на листе остаются вчерашние данные,
  // а в отчёт уходило бодрое «✅ +0». Такой прогон честнее показать ошибкой.
  if (allRows.length === 0) {
    throw new Error('OZON вернул 0 строк — лист не перезаписан, на нём остались данные прошлого прогона');
  }

  writeOzonSheet_(ss, cab.sheetName, allRows, headers);
  return allRows.length;
}

/**
 * Пишет данные в лист OZON: каждый запуск — снапшот текущего состояния.
 * Очищает лист перед записью, чтобы не накапливались устаревшие строки.
 */
function writeOzonSheet_(ss, name, rows, headers) {
  if (rows.length === 0) return;

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
