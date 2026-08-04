// ════════════════════════════════════════════════════════════
// МОДУЛЬ: OZON API — сбор остатков по складам
// ════════════════════════════════════════════════════════════
//
// Источников два, и они дополняют друг друга (сверено на боевом кабинете 03.08.2026):
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

  const now     = new Date();
  const onStock = fetchOzonOnWarehouses_(cab);   // ключ: sku|склад
  const full    = fetchOzonAnalyticsStocks_(cab); // ключ: sku|склад

  // Объединяем: строка появляется, если она есть хотя бы в одном источнике
  const keys = {};
  Object.keys(onStock).forEach(k => { keys[k] = true; });
  Object.keys(full).forEach(k => { keys[k] = true; });

  const rows = Object.keys(keys).map(key => {
    const o = onStock[key] || {};
    const f = full[key]    || {};

    const free     = o.free     || 0;
    const promised = o.promised || 0;
    const reserved = o.reserved || 0;

    // «Всего у OZON» (колонка 9) сохраняет прежний смысл, чтобы не сломать формулы
    const legacyTotal = free + promised + reserved;

    // «Всего вверено OZON» — то, что физически у маркетплейса: остатки на складах,
    // товары в пути, возвраты и брак. «Заявлено к поставке» сюда НЕ входит:
    // это ещё не отгруженная заявка, её видно отдельной колонкой.
    const entrusted = Math.max(free, f.available || 0) + promised + reserved +
      (f.transit || 0) + (f.returnFromCustomer || 0) + (f.returnToSeller || 0) +
      (f.defectStock || 0) + (f.defectTransit || 0) + (f.waitingDocs || 0) + (f.other || 0);

    return [
      now,
      o.offerId || f.offerId || '',
      o.sku     || f.sku     || '',
      o.name    || f.name    || '',
      o.warehouse || f.warehouse || '',
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
      map[ozonKey_(r.sku, r.warehouse_name)] = {
        sku:       r.sku || '',
        offerId:   r.item_code || '',
        name:      r.item_name || '',
        warehouse: r.warehouse_name || '',
        free:      r.free_to_sell_amount || 0,
        promised:  r.promised_amount     || 0,
        reserved:  r.reserved_amount     || 0
      };
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

  for (let i = 0; i < skus.length; i += OZON_ANALYTICS_BATCH) {
    const batch = skus.slice(i, i + OZON_ANALYTICS_BATCH);
    const res   = fetchWithRetry(OZON_URL_ANALYTICS_STOCKS, ozonOptions_(cab, { skus: batch }));
    const items = JSON.parse(res.getContentText()).items || [];

    items.forEach(it => {
      map[ozonKey_(it.sku, it.warehouse_name)] = {
        sku:                it.sku || '',
        offerId:            it.offer_id || '',
        name:               it.name || '',
        warehouse:          it.warehouse_name || '',
        cluster:            it.cluster_name || '',
        available:          it.available_stock_count           || 0,
        transit:            it.transit_stock_count             || 0,
        requested:          it.requested_stock_count           || 0,
        returnFromCustomer: it.return_from_customer_stock_count || 0,
        returnToSeller:     it.return_to_seller_stock_count    || 0,
        defectStock:        it.stock_defect_stock_count        || 0,
        defectTransit:      it.transit_defect_stock_count      || 0,
        waitingDocs:        it.waiting_docs_stock_count        || 0,
        other:              it.other_stock_count               || 0
      };
    });
  }

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

/** Ключ склейки двух источников. Названия складов у ручек различаются регистром. */
function ozonKey_(sku, warehouseName) {
  return String(sku) + '|' + String(warehouseName || '').trim().toUpperCase();
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
    // в колонке I). Дописывать в него 20 колонок нельзя — данные встанут криво.
    // Старое содержимое сохраняем отдельной копией и начинаем лист заново.
    archiveSheetCopy_(ss, sheet);
    sheet.clear();
    writeOzonHeaders_(sheet, headers);
  } else {
    deleteRowsOfDay_(sheet, rows[0][0]);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
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
