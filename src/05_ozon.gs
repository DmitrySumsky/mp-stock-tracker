// ════════════════════════════════════════════════════════════
// МОДУЛЬ: OZON API — сбор остатков по складам
// ════════════════════════════════════════════════════════════

/**
 * Загружает остатки OZON для одного кабинета и записывает в лист.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} cab — объект кабинета из loadCabinets()
 * @returns {number} количество загруженных строк
 */
function fetchOzon(ss, cab) {
  if (!cab.clientId || !cab.apiKey) throw new Error('Client ID или API Key пустые');

  const url  = 'https://api-seller.ozon.ru/v2/analytics/stock_on_warehouses';
  const now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');
  const headers = [
    'Offer ID', 'SKU', 'Item Name', 'Free to Sell', 'Promised Amount',
    'Reserved Amount', 'Warehouse Name', 'IDC', 'Date and Time'
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

    const mapped = data.map(r => [
      r.item_code        || '',
      r.sku              || '',
      r.item_name        || '',
      r.free_to_sell_amount  || 0,
      r.promised_amount  || 0,
      r.reserved_amount  || 0,
      r.warehouse_name   || '',
      r.idc              || '',
      now
    ]);

    allRows = allRows.concat(mapped);
    offset += data.length;
    if (data.length < limit) break;
  }

  writeToSheet(ss, cab.sheetName, allRows, headers);
  return allRows.length;
}
