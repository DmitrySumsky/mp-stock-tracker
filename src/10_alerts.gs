// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Оповещения — низкие остатки и проверка API
// ════════════════════════════════════════════════════════════

/**
 * Проверяет все активные кабинеты и отправляет Telegram-уведомление
 * о позициях с остатком ≤ LOW_STOCK_THRESHOLD.
 */
function checkLowStock() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cabinets = getActiveCabinets();
  const alerts   = [];

  cabinets.forEach(cab => {
    const sheet = ss.getSheetByName(cab.sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const data = sheet.getDataRange().getValues();

    if (cab.mp === 'OZON') {
      for (let i = 1; i < data.length; i++) {
        const qty = Number(data[i][3]) || 0;
        if (qty > 0 && qty <= LOW_STOCK_THRESHOLD) {
          alerts.push(`⚠️ OZON ${cab.id}: "${data[i][2]}" — ${qty} шт. (${data[i][6]})`);
        }
      }
    } else if (cab.mp === 'WB') {
      for (let i = 1; i < data.length; i++) {
        const qty = Number(data[i][5]) || 0;
        if (qty > 0 && qty <= LOW_STOCK_THRESHOLD) {
          alerts.push(`⚠️ WB ${cab.id}: "${data[i][2]}" — ${qty} шт. (${data[i][1]})`);
        }
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
 * Делает тестовые запросы ко всем активным кабинетам и выводит результат.
 */
function healthCheck() {
  const cabinets = getActiveCabinets();
  const results  = [];

  cabinets.forEach(cab => {
    try {
      if (cab.mp === 'OZON') {
        if (!cab.clientId || !cab.apiKey) {
          results.push(`⚠️ OZON ${cab.id}: ключи пустые`);
          return;
        }
        const res = UrlFetchApp.fetch(
          'https://api-seller.ozon.ru/v2/analytics/stock_on_warehouses',
          {
            method: 'post', contentType: 'application/json',
            headers: { 'client-id': cab.clientId, 'api-key': cab.apiKey },
            payload: JSON.stringify({ limit: 1, offset: 0, warehouse_type: 'ALL' }),
            muteHttpExceptions: true
          }
        );
        results.push(res.getResponseCode() === 200
          ? `✅ OZON ${cab.id}: OK`
          : `❌ OZON ${cab.id}: HTTP ${res.getResponseCode()}`);

      } else if (cab.mp === 'WB') {
        if (!cab.token) {
          results.push(`⚠️ WB ${cab.id}: токен пустой`);
          return;
        }
        const res = UrlFetchApp.fetch(
          'https://statistics-api.wildberries.ru/api/v1/supplier/stocks?dateFrom=2024-01-01',
          { headers: { Authorization: cab.token }, muteHttpExceptions: true }
        );
        results.push(res.getResponseCode() === 200
          ? `✅ WB ${cab.id}: OK`
          : `❌ WB ${cab.id}: HTTP ${res.getResponseCode()}`);
      }
    } catch (e) {
      results.push(`❌ ${cab.mp} ${cab.id}: ${e.message}`);
    }
  });

  const msg = ['🏥 *Проверка API*', '', ...results].join('\n');
  sendTelegram(msg);
  SpreadsheetApp.getUi().alert(results.join('\n'));
}
