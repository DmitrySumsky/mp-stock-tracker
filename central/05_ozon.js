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
