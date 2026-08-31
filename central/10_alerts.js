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
