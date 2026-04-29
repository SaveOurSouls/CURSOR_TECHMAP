const TECHOPS_DB_APP = {
  sourceSpreadsheetId: '1W3VK9Fw71lYdw1Klcsn_za5-2EhvLoXIAKZVYOCnKcs',
  metaSheetName: '_TC_TECHOPS_META',
  dataSheetName: '_TC_TECHOPS_DB',
  metaHeaders: ['key', 'value'],
  dataHeaders: ['tabKey', 'displayText', 'normalizedSearch', 'exportJson', 'sourceSheet'],
  cacheKeyPrefix: 'techmap-techops-db-v1',
  cacheChunkSize: 80000,
  cacheTtlSeconds: 21600,
  tabs: {
    ob: {
      key: 'ob',
      label: 'БД.ОБ',
      sourceSheetName: 'БД.ОБ',
      searchPlaceholder: 'Поиск по полю "Для базы"...',
      outputLabels: ['Для базы'],
    },
    op: {
      key: 'op',
      label: 'БД.ОП',
      sourceSheetName: 'БД.ОП',
      searchPlaceholder: 'Поиск по номеру или названию операции...',
      outputLabels: [
        'Номер | Название',
        'Время Операции',
        'Время подготовки, сек',
        'Расход на настройку м; шт;',
        'Время машины, сек/оп; сек/м',
      ],
    },
    ter: {
      key: 'ter',
      label: 'БД.ТЕР',
      sourceSheetName: 'БД.ТЕР',
      searchPlaceholder: 'Поиск по комплектующей, аналогу, серии или производителю...',
      outputLabels: ['Комплектующая', 'Аналог', 'Серия разъемов', 'Производитель'],
    },
    coax: {
      key: 'coax',
      label: 'БД.КОАКС',
      sourceSheetName: 'БД.КОАКС',
      searchPlaceholder: 'Поиск по артикулам, сериям, проводу, размерам...',
      outputLabels: [
        'Артикул',
        'Тип/Серия',
        'Производитель',
        'Поставщик',
        'Провод',
        'Программа',
        'D1',
        'D2',
        'D3',
        'L1',
        'L2',
        'L3',
      ],
    },
  },
  tabOrder: ['ob', 'op', 'ter', 'coax'],
};

if (typeof TECHMAP_DATA_MODEL !== 'undefined' && TECHMAP_DATA_MODEL.techOperationsSource) {
  const techOpsModel = TECHMAP_DATA_MODEL.techOperationsSource;
  if (techOpsModel.spreadsheetId) {
    TECHOPS_DB_APP.sourceSpreadsheetId = techOpsModel.spreadsheetId;
  }
}

function refreshTechOperationsDatabase() {
  return syncTechOperationsDatabaseMenu();
}

// Backward-compatible alias used by menu/sidebar wiring.
function refreshOperationDatabase() {
  return refreshTechOperationsDatabase();
}

function syncTechOperationsDatabaseMenu() {
  const summary = syncTechOperationsDatabase();
  SpreadsheetApp.getUi().alert(
    'База техопераций обновлена.',
    `Загружено строк: ${summary.recordCount}\nИсточник: ${summary.sourceSpreadsheetId}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return summary;
}

function syncTechOperationsDatabase() {
  const ss = SpreadsheetApp.getActive();
  ensureTechOperationsInfrastructure_(ss);

  const snapshot = fetchTechOperationsSnapshotFromSource_();
  writeTechOperationsSnapshotToSheets_(snapshot);
  cacheTechOperationsSnapshot_(snapshot);
  hideTechOperationsSheets_();
  return buildTechOperationsSummary_(snapshot);
}

function getTechOperationsDatabase(forceRefresh) {
  if (forceRefresh) {
    syncTechOperationsDatabase();
  } else {
    ensureTechOperationsDatabaseReady_();
  }

  return buildTechOperationsPayload_(getTechOperationsSnapshot_());
}

// Backward-compatible alias used by the workspace sidebar.
function getOperationDatabase(forceRefresh) {
  return getTechOperationsDatabase(forceRefresh);
}

function insertTechOperationMatrix(matrix, targetCellA1) {
  if (!matrix || !matrix.length) {
    return 'Нет данных для выгрузки';
  }

  const width = Array.isArray(matrix[0]) ? matrix[0].length : 0;
  if (!width) {
    return 'Нет данных для выгрузки';
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  let startCell = null;

  if (targetCellA1) {
    try {
      startCell = sheet.getRange(targetCellA1);
    } catch (error) {
      return `ОШИБКА: Неверный адрес ячейки (${targetCellA1}).`;
    }
  } else {
    startCell = ss.getCurrentCell() || sheet.getActiveCell();
  }

  if (!startCell) {
    return 'ОШИБКА: Не выбрана стартовая ячейка.';
  }

  ensureSheetCapacity_(
    sheet,
    startCell.getRow() + matrix.length - 1,
    startCell.getColumn() + width - 1
  );

  sheet.getRange(startCell.getRow(), startCell.getColumn(), matrix.length, width).setValues(matrix);
  sheet.getRange(startCell.getRow() + matrix.length, startCell.getColumn()).activate();
  return `Успешно выгружено ${matrix.length} строк.`;
}

// Backward-compatible alias used by the workspace sidebar.
function insertOperationRows(matrix, targetCellA1) {
  return {
    message: insertTechOperationMatrix(matrix, targetCellA1),
  };
}

function ensureTechOperationsDatabaseReady_() {
  ensureTechOperationsInfrastructure_(SpreadsheetApp.getActive());
  const snapshot = getTechOperationsSnapshot_();
  if (!snapshot.records.length) {
    syncTechOperationsDatabase();
  }
}

function ensureTechOperationsInfrastructure_(ssArg) {
  const ss = ssArg || SpreadsheetApp.getActive();
  ensureTechOperationsMetaSheet_(ss);
  ensureTechOperationsDataSheet_(ss);
}

function ensureTechOperationsMetaSheet_(ss) {
  let sheet = ss.getSheetByName(TECHOPS_DB_APP.metaSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHOPS_DB_APP.metaSheetName);
  }

  ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.metaHeaders.length);
  sheet
    .getRange(1, 1, 1, TECHOPS_DB_APP.metaHeaders.length)
    .setValues([TECHOPS_DB_APP.metaHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');
  sheet.hideSheet();
  return sheet;
}

function ensureTechOperationsDataSheet_(ss) {
  let sheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHOPS_DB_APP.dataSheetName);
  }

  ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.dataHeaders.length);
  sheet
    .getRange(1, 1, 1, TECHOPS_DB_APP.dataHeaders.length)
    .setValues([TECHOPS_DB_APP.dataHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');
  sheet.hideSheet();
  return sheet;
}

function fetchTechOperationsSnapshotFromSource_() {
  const sourceSs = SpreadsheetApp.openById(TECHOPS_DB_APP.sourceSpreadsheetId);
  const records = [];
  const countsByTab = {};

  TECHOPS_DB_APP.tabOrder.forEach((tabKey) => {
    const config = TECHOPS_DB_APP.tabs[tabKey];
    countsByTab[tabKey] = 0;
    const sheet = sourceSs.getSheetByName(config.sourceSheetName);
    if (!sheet) {
      return;
    }

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) {
      return;
    }

    const headerRowIndex = detectTechOperationsHeaderRow_(values, tabKey);
    if (headerRowIndex < 0) {
      return;
    }

    const headerMap = buildTechOperationsHeaderMap_(values[headerRowIndex]);
    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const record = buildTechOperationsRecordFromRow_(tabKey, row, headerMap, config.sourceSheetName);
      if (!record || !record.displayText) {
        continue;
      }
      countsByTab[tabKey] += 1;
      records.push(record);
    }
  });

  records.sort((a, b) => {
    if (a.tabKey !== b.tabKey) {
      return TECHOPS_DB_APP.tabOrder.indexOf(a.tabKey) - TECHOPS_DB_APP.tabOrder.indexOf(b.tabKey);
    }
    return String(a.displayText || '').localeCompare(String(b.displayText || ''), undefined, {
      numeric: true,
      sensitivity: 'base',
    });
  });

  return {
    meta: {
      sourceSpreadsheetId: TECHOPS_DB_APP.sourceSpreadsheetId,
      updatedAt: new Date().toISOString(),
      recordCount: records.length,
      countsByTab,
    },
    records,
  };
}

function buildTechOperationsHeaderMap_(headersRow) {
  const map = {};
  headersRow.forEach((header, index) => {
    const normalized = normalizeTechOperationsHeader_(header);
    if (normalized) {
      map[normalized] = index;
    }
  });
  return map;
}

function detectTechOperationsHeaderRow_(values, tabKey) {
  const aliases = getTechOperationsHeaderAliasesForTab_(tabKey);
  const maxScanRows = Math.min(values.length, 12);
  let bestIndex = -1;
  let bestScore = -1;

  for (let rowIndex = 0; rowIndex < maxScanRows; rowIndex += 1) {
    const headerMap = buildTechOperationsHeaderMap_(values[rowIndex]);
    let score = 0;
    aliases.forEach((aliasGroup) => {
      const found = aliasGroup.some((alias) => {
        const key = normalizeTechOperationsHeader_(alias);
        return headerMap[key] === 0 || headerMap[key] > 0;
      });
      if (found) {
        score += 1;
      }
    });

    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }

  const minimumScore = Math.max(1, Math.min(2, aliases.length));
  return bestScore >= minimumScore ? bestIndex : -1;
}

function getTechOperationsHeaderAliasesForTab_(tabKey) {
  switch (tabKey) {
    case 'ob':
      return [['для базы', 'длябазы']];
    case 'op':
      return [
        ['номер', 'number'],
        ['название', 'name'],
        ['время операции', 'время операции, сек', 'время операции сек'],
        ['время подготовки, сек', 'время подготовки сек'],
        ['время машины, сек/оп; сек/м', 'время машины сек/оп; сек/м', 'время машины'],
        ['тип операции', 'типоперации'],
      ];
    case 'ter':
      return [
        ['комплектующая'],
        ['аналог'],
        ['серия разъемов', 'серияразъемов'],
        ['производитель', 'бренд', 'manufacturer'],
      ];
    case 'coax':
      return [
        ['артикул'],
        ['тип/серия', 'тип / серия', 'тип серия'],
        ['производитель', 'бренд', 'manufacturer'],
        ['поставщик', 'supplier'],
        ['провод'],
        ['программа'],
      ];
    default:
      return [];
  }
}

function buildTechOperationsRecordFromRow_(tabKey, row, headerMap, sourceSheet) {
  switch (tabKey) {
    case 'ob':
      return buildTechOperationsObRecord_(row, headerMap, sourceSheet);
    case 'op':
      return buildTechOperationsOpRecord_(row, headerMap, sourceSheet);
    case 'ter':
      return buildTechOperationsTerRecord_(row, headerMap, sourceSheet);
    case 'coax':
      return buildTechOperationsCoaxRecord_(row, headerMap, sourceSheet);
    default:
      return null;
  }
}

function buildTechOperationsObRecord_(row, headerMap, sourceSheet) {
  const baseValue = getTechOperationsCellByAliases_(row, headerMap, ['для базы', 'длябазы']);
  if (!baseValue) {
    return null;
  }
  return {
    tabKey: 'ob',
    displayText: baseValue,
    normalizedSearch: normalizeTechOperationsSearch_(baseValue),
    exportValues: [baseValue],
    sourceSheet,
  };
}

function buildTechOperationsOpRecord_(row, headerMap, sourceSheet) {
  const number = getTechOperationsCellByAliases_(row, headerMap, ['номер', 'number']);
  const name = getTechOperationsCellByAliases_(row, headerMap, ['название', 'name']);
  const displayText = joinTechOperationsParts_([number, name], ' | ');
  if (!displayText) {
    return null;
  }

  const values = [
    displayText,
    getTechOperationsCellByAliases_(row, headerMap, [
      'время операции',
      'время операции, сек',
      'время операции сек',
    ]),
    getTechOperationsCellByAliases_(row, headerMap, [
      'время подготовки, сек',
      'время подготовки сек',
    ]),
    getTechOperationsCellByAliases_(row, headerMap, [
      'расход на настройку м; шт;',
      'расход на настройку м;шт;',
      'расход на настройку',
    ]),
    getTechOperationsCellByAliases_(row, headerMap, [
      'время машины, сек/оп; сек/м',
      'время машины сек/оп; сек/м',
      'время машины',
    ]),
  ];

  return {
    tabKey: 'op',
    displayText,
    normalizedSearch: normalizeTechOperationsSearch_(joinTechOperationsParts_(values, ' | ')),
    exportValues: values,
    sourceSheet,
  };
}

function buildTechOperationsTerRecord_(row, headerMap, sourceSheet) {
  const values = [
    getTechOperationsCellByAliases_(row, headerMap, ['комплектующая']),
    getTechOperationsCellByAliases_(row, headerMap, ['аналог']),
    getTechOperationsCellByAliases_(row, headerMap, ['серия разъемов', 'серияразъемов']),
    getTechOperationsCellByAliases_(row, headerMap, ['производитель', 'бренд', 'manufacturer']),
  ];
  const displayText = joinTechOperationsParts_(values, ' | ');
  if (!displayText) {
    return null;
  }
  return {
    tabKey: 'ter',
    displayText,
    normalizedSearch: normalizeTechOperationsSearch_(displayText),
    exportValues: values,
    sourceSheet,
  };
}

function buildTechOperationsCoaxRecord_(row, headerMap, sourceSheet) {
  const values = [
    getTechOperationsCellByAliases_(row, headerMap, ['артикул']),
    getTechOperationsCellByAliases_(row, headerMap, ['тип/серия', 'тип / серия', 'тип серия']),
    getTechOperationsCellByAliases_(row, headerMap, ['производитель', 'бренд', 'manufacturer']),
    getTechOperationsCellByAliases_(row, headerMap, ['поставщик', 'supplier']),
    getTechOperationsCellByAliases_(row, headerMap, ['провод']),
    getTechOperationsCellByAliases_(row, headerMap, ['программа']),
    getTechOperationsCellByAliases_(row, headerMap, ['d1']),
    getTechOperationsCellByAliases_(row, headerMap, ['d2']),
    getTechOperationsCellByAliases_(row, headerMap, ['d3']),
    getTechOperationsCellByAliases_(row, headerMap, ['l1']),
    getTechOperationsCellByAliases_(row, headerMap, ['l2']),
    getTechOperationsCellByAliases_(row, headerMap, ['l3']),
  ];
  const displayText = joinTechOperationsParts_(values, ' | ');
  if (!displayText) {
    return null;
  }
  return {
    tabKey: 'coax',
    displayText,
    normalizedSearch: normalizeTechOperationsSearch_(displayText),
    exportValues: values,
    sourceSheet,
  };
}

function getTechOperationsCellByAliases_(row, headerMap, aliases) {
  for (let index = 0; index < aliases.length; index += 1) {
    const headerIndex = headerMap[normalizeTechOperationsHeader_(aliases[index])];
    if (headerIndex === 0 || headerIndex > 0) {
      return normalizeString_(row[headerIndex]);
    }
  }
  return '';
}

function normalizeTechOperationsHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeTechOperationsSearch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, '');
}

function joinTechOperationsParts_(parts, delimiter) {
  return (parts || []).filter((part) => normalizeString_(part)).join(delimiter || ' | ');
}

function writeTechOperationsSnapshotToSheets_(snapshot) {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ensureTechOperationsDataSheet_(ss);
  const metaSheet = ensureTechOperationsMetaSheet_(ss);

  dataSheet.clearContents();
  ensureSheetCapacity_(
    dataSheet,
    Math.max(snapshot.records.length + 1, 2),
    TECHOPS_DB_APP.dataHeaders.length
  );
  dataSheet
    .getRange(1, 1, 1, TECHOPS_DB_APP.dataHeaders.length)
    .setValues([TECHOPS_DB_APP.dataHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');

  if (snapshot.records.length) {
    const rows = snapshot.records.map((record) => [
      record.tabKey,
      record.displayText,
      record.normalizedSearch,
      JSON.stringify(record.exportValues || []),
      record.sourceSheet,
    ]);
    dataSheet.getRange(2, 1, rows.length, TECHOPS_DB_APP.dataHeaders.length).setValues(rows);
  }

  metaSheet.clearContents();
  ensureSheetCapacity_(metaSheet, 6, TECHOPS_DB_APP.metaHeaders.length);
  metaSheet
    .getRange(1, 1, 1, TECHOPS_DB_APP.metaHeaders.length)
    .setValues([TECHOPS_DB_APP.metaHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');

  const metaRows = [
    ['sourceSpreadsheetId', snapshot.meta.sourceSpreadsheetId],
    ['updatedAt', snapshot.meta.updatedAt],
    ['recordCount', String(snapshot.meta.recordCount || 0)],
    ['countsByTabJson', JSON.stringify(snapshot.meta.countsByTab || {})],
  ];
  metaSheet.getRange(2, 1, metaRows.length, 2).setValues(metaRows);
  hideTechOperationsSheets_();
}

function getTechOperationsSnapshot_() {
  const cached = loadTechOperationsSnapshotFromCache_();
  if (cached && cached.records && cached.records.length) {
    return cached;
  }

  const stored = loadTechOperationsSnapshotFromSheets_();
  if (stored.records.length) {
    cacheTechOperationsSnapshot_(stored);
  }
  return stored;
}

function loadTechOperationsSnapshotFromSheets_() {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ensureTechOperationsDataSheet_(ss);
  const metaSheet = ensureTechOperationsMetaSheet_(ss);

  const records = [];
  const lastRow = dataSheet.getLastRow();
  if (lastRow >= 2) {
    const values = dataSheet
      .getRange(2, 1, lastRow - 1, TECHOPS_DB_APP.dataHeaders.length)
      .getValues()
      .filter((row) => row[0] && row[1]);

    values.forEach((row) => {
      records.push({
        tabKey: row[0],
        displayText: row[1],
        normalizedSearch: row[2],
        exportValues: parseJsonArray_(row[3]),
        sourceSheet: row[4],
      });
    });
  }

  const meta = {
    sourceSpreadsheetId: TECHOPS_DB_APP.sourceSpreadsheetId,
    updatedAt: '',
    recordCount: records.length,
    countsByTab: {},
  };

  const metaLastRow = metaSheet.getLastRow();
  if (metaLastRow >= 2) {
    const metaRows = metaSheet.getRange(2, 1, metaLastRow - 1, 2).getValues();
    metaRows.forEach((row) => {
      const key = row[0];
      const value = row[1];
      if (key === 'sourceSpreadsheetId') {
        meta.sourceSpreadsheetId = value || TECHOPS_DB_APP.sourceSpreadsheetId;
      } else if (key === 'updatedAt') {
        meta.updatedAt = value || '';
      } else if (key === 'recordCount') {
        meta.recordCount = toInt_(value);
      } else if (key === 'countsByTabJson') {
        try {
          meta.countsByTab = JSON.parse(value) || {};
        } catch (error) {
          meta.countsByTab = {};
        }
      }
    });
  }

  return { meta, records };
}

function buildTechOperationsPayload_(snapshot) {
  const payload = {
    meta: buildTechOperationsSummary_(snapshot),
    tabs: {},
    dbOb: [],
    dbOp: [],
    dbTer: [],
    dbKoax: [],
  };

  const payloadKeyMap = {
    ob: 'dbOb',
    op: 'dbOp',
    ter: 'dbTer',
    coax: 'dbKoax',
  };

  TECHOPS_DB_APP.tabOrder.forEach((tabKey) => {
    const config = TECHOPS_DB_APP.tabs[tabKey];
    const items = (snapshot.records || [])
      .filter((record) => record.tabKey === tabKey)
      .map((record, index) => ({
        id: `${tabKey}-${index}`,
        displayText: record.displayText,
        label: record.displayText,
        searchText: record.normalizedSearch,
        values: record.exportValues || [],
        outputRow: record.exportValues || [],
        sourceSheet: record.sourceSheet,
      }));

    payload.tabs[tabKey] = {
      key: tabKey,
      label: config.label,
      searchPlaceholder: config.searchPlaceholder,
      outputLabels: config.outputLabels.slice(),
      items,
      count: items.length,
    };
    payload[payloadKeyMap[tabKey]] = items;
  });

  return payload;
}

function buildTechOperationsSummary_(snapshot) {
  return {
    sourceSpreadsheetId: snapshot.meta.sourceSpreadsheetId || TECHOPS_DB_APP.sourceSpreadsheetId,
    updatedAt: snapshot.meta.updatedAt || '',
    recordCount: snapshot.meta.recordCount || (snapshot.records ? snapshot.records.length : 0),
    countsByTab: snapshot.meta.countsByTab || {},
  };
}

function cacheTechOperationsSnapshot_(snapshot) {
  const cache = CacheService.getDocumentCache();
  const serialized = JSON.stringify(snapshot);
  const chunkCount = Math.ceil(serialized.length / TECHOPS_DB_APP.cacheChunkSize) || 1;
  cache.put(
    `${TECHOPS_DB_APP.cacheKeyPrefix}:count`,
    String(chunkCount),
    TECHOPS_DB_APP.cacheTtlSeconds
  );

  for (let index = 0; index < chunkCount; index += 1) {
    const chunk = serialized.slice(
      index * TECHOPS_DB_APP.cacheChunkSize,
      (index + 1) * TECHOPS_DB_APP.cacheChunkSize
    );
    cache.put(
      `${TECHOPS_DB_APP.cacheKeyPrefix}:chunk:${index}`,
      chunk,
      TECHOPS_DB_APP.cacheTtlSeconds
    );
  }
}

function loadTechOperationsSnapshotFromCache_() {
  const cache = CacheService.getDocumentCache();
  const countValue = cache.get(`${TECHOPS_DB_APP.cacheKeyPrefix}:count`);
  const count = toInt_(countValue);
  if (!count) {
    return null;
  }

  let serialized = '';
  for (let index = 0; index < count; index += 1) {
    const chunk = cache.get(`${TECHOPS_DB_APP.cacheKeyPrefix}:chunk:${index}`);
    if (chunk === null || chunk === undefined) {
      return null;
    }
    serialized += chunk;
  }

  try {
    return JSON.parse(serialized);
  } catch (error) {
    return null;
  }
}

function hideTechOperationsSheets_() {
  const ss = SpreadsheetApp.getActive();
  [TECHOPS_DB_APP.metaSheetName, TECHOPS_DB_APP.dataSheetName].forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.hideSheet();
    }
  });
}

function isTechOperationsSystemSheet_(sheetName) {
  return (
    sheetName === TECHOPS_DB_APP.metaSheetName ||
    sheetName === TECHOPS_DB_APP.dataSheetName
  );
}
