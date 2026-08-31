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
