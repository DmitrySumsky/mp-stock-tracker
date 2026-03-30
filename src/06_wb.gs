// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Wildberries API — сбор остатков
// ════════════════════════════════════════════════════════════

/**
 * Загружает остатки WB для одного кабинета и записывает в лист.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} cab — объект кабинета из loadCabinets()
 * @returns {number} количество загруженных строк
 */
function fetchWB(ss, cab) {
  if (!cab.token) throw new Error('Token пустой');

  const url = 'https://statistics-api.wildberries.ru/api/v1/supplier/stocks?dateFrom=2019-06-20T00:00:00';
  const res = fetchWithRetry(url, {
    headers:           { Authorization: cab.token },
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());
  const now  = new Date();

  const headers = [
    'Дата и время', 'Склад', 'Артикул продавца', 'Артикул WB', 'Баркод',
    'Количество', 'В пути к клиенту', 'В пути от клиента', 'Полное количество', 'Цена', 'Скидка'
  ];

  const rows = data
    .map(item => [
      now,
      item.warehouseName,
      item.supplierArticle,
      item.nmId,
      item.barcode,
      item.quantity,
      item.inWayToClient,
      item.inWayFromClient,
      item.quantityFull,
      item.Price,
      item.Discount
    ])
    .sort((a, b) => String(a[2]).localeCompare(String(b[2]), 'ru', { numeric: true }));

  writeToSheet(ss, cab.sheetName, rows, headers);
  return rows.length;
}
