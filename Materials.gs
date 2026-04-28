const MATERIAL_DB_APP = {
  sourceSpreadsheetId: '1NExDzeG-vw3zY_ooeXoxRffIARh2wbLJNAHT3FU_Ig8',
  sourceSheetNames: ['COMPCON', 'COMPCOAX', 'COMPTERM', 'COMPWIRE', 'COMPACCESS'],
  metaSheetName: '_TC_MATERIAL_META',
  dataSheetName: '_TC_MATERIAL_DB',
  metaHeaders: ['key', 'value'],
  dataHeaders: [
    'tag',
    'article',
    'type',
    'manufacturer',
    'supplier',
    'sourceSheet',
    'normalizedTag',
    'normalizedType',
    'normalizedManufacturer',
    'normalizedSupplier',
  ],
  cacheKeyPrefix: 'techmap-material-db-v1',
  cacheChunkSize: 80000,
  cacheTtlSeconds: 21600,
};

if (typeof TECHMAP_DATA_MODEL !== 'undefined' && TECHMAP_DATA_MODEL.materialDatabase) {
  const materialModel = TECHMAP_DATA_MODEL.materialDatabase;
  if (materialModel.serviceSheets) {
    MATERIAL_DB_APP.metaSheetName = materialModel.serviceSheets.metaSheetName;
    MATERIAL_DB_APP.dataSheetName = materialModel.serviceSheets.dataSheetName;
  }
}

function showMaterialsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MaterialsSidebar')
    .setTitle('Материалы')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function refreshMaterialsDatabase() {
  return syncMaterialDatabaseMenu();
}

function syncMaterialDatabaseMenu() {
  const summary = syncMaterialDatabase();
  SpreadsheetApp.getUi().alert(
    'База материалов обновлена.',
    `Загружено позиций: ${summary.recordCount}\nИсточник: ${summary.sourceSpreadsheetId}\nЛисты: ${summary.sourceSheets.join(', ')}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncMaterialDatabase() {
  const ss = SpreadsheetApp.getActive();
  ensureMaterialInfrastructure_(ss);

  const snapshot = fetchMaterialSnapshotFromSource_();
  writeMaterialSnapshotToSheets_(snapshot);
  cacheMaterialSnapshot_(snapshot);
  hideMaterialSheets_();
  return buildMaterialSummary_(snapshot);
}

function getMaterialSearchData() {
  ensureMaterialDatabaseReady_();
  const snapshot = getMaterialSnapshot_();
  return buildMaterialSearchPayload_(snapshot);
}

function getMaterialDatabase(forceRefresh) {
  if (forceRefresh) {
    syncMaterialDatabase();
  } else {
    ensureMaterialDatabaseReady_();
  }

  const payload = buildMaterialSearchPayload_(getMaterialSnapshot_());
  return {
    items: payload.records,
    lookups: {
      types: payload.uniqueTypes,
      manufacturers: payload.uniqueManufacturers,
      suppliers: payload.uniqueSuppliers,
    },
    meta: payload.meta,
  };
}

/**
 * Совместимость со старой HTML-панелью пользователя.
 * Возвращает строку формата "tag|||tag|||tag".
 */
function getSearchData() {
  const payload = getMaterialSearchData();
  const tags = payload.records.map((item) => item.fullTag);
  if (!tags.length) {
    return "ОШИБКА: База материалов пуста.";
  }
  return tags.join('|||');
}

function insertBatchIntoCell(valuesArray) {
  if (!valuesArray || !valuesArray.length) {
    return 'Список пуст';
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const currentCell = ss.getCurrentCell() || sheet.getActiveCell();
  if (!currentCell) {
    return 'ОШИБКА: Не выбрана стартовая ячейка.';
  }

  const startRow = currentCell.getRow();
  const startColumn = currentCell.getColumn();
  if (startRow < 1) {
    return 'ОШИБКА: Не удалось определить строку вставки.';
  }

  ensureSheetCapacity_(sheet, startRow + valuesArray.length - 1, startColumn);
  const matrix = valuesArray.map((item) => [item]);
  sheet.getRange(startRow, startColumn, matrix.length, 1).setValues(matrix);
  sheet.getRange(startRow + matrix.length, startColumn).activate();
  return `Успешно добавлено ${matrix.length} поз.`;
}

function insertReplicatedData(valuesArray, targetCellA1) {
  if (!valuesArray || !valuesArray.length) {
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
    startCell = sheet.getActiveCell();
  }

  if (!startCell) {
    return 'ОШИБКА: Не выбрана стартовая ячейка.';
  }

  ensureSheetCapacity_(sheet, startCell.getRow() + valuesArray.length - 1, startCell.getColumn());
  const matrix = valuesArray.map((value) => [value]);
  sheet.getRange(startCell.getRow(), startCell.getColumn(), matrix.length, 1).setValues(matrix);
  return `Успешно выгружено ${matrix.length} строк.`;
}

function getMaterialDatabaseInfo() {
  ensureMaterialDatabaseReady_();
  const snapshot = getMaterialSnapshot_();
  return buildMaterialSummary_(snapshot);
}

function ensureMaterialDatabaseReady_() {
  ensureMaterialInfrastructure_(SpreadsheetApp.getActive());
  const snapshot = getMaterialSnapshot_();
  if (!snapshot.records.length) {
    syncMaterialDatabase();
  }
}

function ensureMaterialInfrastructure_(ssArg) {
  const ss = ssArg || SpreadsheetApp.getActive();
  ensureMaterialMetaSheet_(ss);
  ensureMaterialDataSheet_(ss);
}

function ensureMaterialMetaSheet_(ss) {
  let sheet = ss.getSheetByName(MATERIAL_DB_APP.metaSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(MATERIAL_DB_APP.metaSheetName);
  }

  ensureSheetCapacity_(sheet, 2, MATERIAL_DB_APP.metaHeaders.length);
  sheet
    .getRange(1, 1, 1, MATERIAL_DB_APP.metaHeaders.length)
    .setValues([MATERIAL_DB_APP.metaHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');
  sheet.hideSheet();
  return sheet;
}

function ensureMaterialDataSheet_(ss) {
  let sheet = ss.getSheetByName(MATERIAL_DB_APP.dataSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(MATERIAL_DB_APP.dataSheetName);
  }

  ensureSheetCapacity_(sheet, 2, MATERIAL_DB_APP.dataHeaders.length);
  sheet
    .getRange(1, 1, 1, MATERIAL_DB_APP.dataHeaders.length)
    .setValues([MATERIAL_DB_APP.dataHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');
  sheet.hideSheet();
  return sheet;
}

function fetchMaterialSnapshotFromSource_() {
  const sourceSs = SpreadsheetApp.openById(MATERIAL_DB_APP.sourceSpreadsheetId);
  const uniqueMap = {};
  const records = [];

  MATERIAL_DB_APP.sourceSheetNames.forEach((sheetName) => {
    const sheet = sourceSs.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) {
      return;
    }

    const headerMap = buildMaterialHeaderMap_(values[0]);
    for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const record = buildMaterialRecordFromRow_(row, headerMap, sheetName);
      if (!record || !record.tag) {
        continue;
      }

      const uniqueKey = record.normalizedTag;
      if (!uniqueKey || uniqueMap[uniqueKey]) {
        continue;
      }

      uniqueMap[uniqueKey] = true;
      records.push(record);
    }
  });

  records.sort((a, b) =>
    a.tag.localeCompare(b.tag, undefined, { numeric: true, sensitivity: 'base' })
  );

  return {
    meta: {
      sourceSpreadsheetId: MATERIAL_DB_APP.sourceSpreadsheetId,
      sourceSheets: MATERIAL_DB_APP.sourceSheetNames.slice(),
      updatedAt: new Date().toISOString(),
      recordCount: records.length,
    },
    records,
  };
}

function buildMaterialHeaderMap_(headersRow) {
  const map = {};
  headersRow.forEach((header, index) => {
    const normalized = normalizeMaterialHeader_(header);
    if (normalized) {
      map[normalized] = index;
    }
  });
  return map;
}

function buildMaterialRecordFromRow_(row, headerMap, sourceSheet) {
  const tag = getMaterialCellByHeader_(row, headerMap, [
    'поисковый тег',
    'поисковыйтег',
    'search tag',
    'searchtag',
  ]);

  const article = getMaterialCellByHeader_(row, headerMap, [
    'артикул',
    'обозначение',
    'pn',
    'part number',
    'partnumber',
  ]);
  const type = getMaterialCellByHeader_(row, headerMap, ['тип', 'наименование', 'type', 'description']);
  const manufacturer = getMaterialCellByHeader_(row, headerMap, [
    'производитель',
    'бренд',
    'manufacturer',
    'vendor',
  ]);
  const supplier = getMaterialCellByHeader_(row, headerMap, ['поставщик', 'supplier']);

  let finalTag = normalizeString_(tag);
  if (!finalTag) {
    finalTag = [article, type, manufacturer, supplier].filter(Boolean).join(' | ');
  }
  finalTag = normalizeString_(finalTag);
  if (!finalTag) {
    return null;
  }

  const parsed = parseMaterialTag_(finalTag);
  return {
    tag: finalTag,
    article: article || parsed.article,
    type: type || parsed.type,
    manufacturer: manufacturer || parsed.manufacturer,
    supplier: supplier || parsed.supplier,
    sourceSheet,
    normalizedTag: normalizeMaterialSearch_(finalTag),
    normalizedType: normalizeMaterialSearch_(type || parsed.type),
    normalizedManufacturer: normalizeMaterialSearch_(manufacturer || parsed.manufacturer),
    normalizedSupplier: normalizeMaterialSearch_(supplier || parsed.supplier),
  };
}

function getMaterialCellByHeader_(row, headerMap, aliases) {
  for (let index = 0; index < aliases.length; index += 1) {
    const headerIndex = headerMap[aliases[index]];
    if (headerIndex === 0 || headerIndex > 0) {
      return normalizeString_(row[headerIndex]);
    }
  }
  return '';
}

function parseMaterialTag_(tag) {
  const parts = String(tag || '')
    .split('|')
    .map((part) => normalizeString_(part));
  return {
    article: parts[0] || '',
    type: parts[1] || '',
    manufacturer: parts[2] || '',
    supplier: parts[3] || '',
  };
}

function normalizeMaterialHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeMaterialSearch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, '');
}

function writeMaterialSnapshotToSheets_(snapshot) {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ensureMaterialDataSheet_(ss);
  const metaSheet = ensureMaterialMetaSheet_(ss);

  dataSheet.clearContents();
  ensureSheetCapacity_(
    dataSheet,
    Math.max(snapshot.records.length + 1, 2),
    MATERIAL_DB_APP.dataHeaders.length
  );
  dataSheet
    .getRange(1, 1, 1, MATERIAL_DB_APP.dataHeaders.length)
    .setValues([MATERIAL_DB_APP.dataHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');

  if (snapshot.records.length) {
    const rows = snapshot.records.map((record) => [
      record.tag,
      record.article,
      record.type,
      record.manufacturer,
      record.supplier,
      record.sourceSheet,
      record.normalizedTag,
      record.normalizedType,
      record.normalizedManufacturer,
      record.normalizedSupplier,
    ]);
    dataSheet.getRange(2, 1, rows.length, MATERIAL_DB_APP.dataHeaders.length).setValues(rows);
  }

  metaSheet.clearContents();
  ensureSheetCapacity_(metaSheet, 5, MATERIAL_DB_APP.metaHeaders.length);
  metaSheet
    .getRange(1, 1, 1, MATERIAL_DB_APP.metaHeaders.length)
    .setValues([MATERIAL_DB_APP.metaHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f6fc');

  const metaRows = [
    ['sourceSpreadsheetId', snapshot.meta.sourceSpreadsheetId],
    ['sourceSheetsJson', JSON.stringify(snapshot.meta.sourceSheets || [])],
    ['updatedAt', snapshot.meta.updatedAt],
    ['recordCount', String(snapshot.meta.recordCount || 0)],
  ];
  metaSheet.getRange(2, 1, metaRows.length, 2).setValues(metaRows);
  hideMaterialSheets_();
}

function getMaterialSnapshot_() {
  const cached = loadMaterialSnapshotFromCache_();
  if (cached && cached.records && cached.records.length) {
    return cached;
  }

  const stored = loadMaterialSnapshotFromSheets_();
  if (stored.records.length) {
    cacheMaterialSnapshot_(stored);
  }
  return stored;
}

function loadMaterialSnapshotFromSheets_() {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ensureMaterialDataSheet_(ss);
  const metaSheet = ensureMaterialMetaSheet_(ss);

  const records = [];
  const lastRow = dataSheet.getLastRow();
  if (lastRow >= 2) {
    const values = dataSheet
      .getRange(2, 1, lastRow - 1, MATERIAL_DB_APP.dataHeaders.length)
      .getValues()
      .filter((row) => row[0]);

    values.forEach((row) => {
      records.push({
        tag: row[0],
        article: row[1],
        type: row[2],
        manufacturer: row[3],
        supplier: row[4],
        sourceSheet: row[5],
        normalizedTag: row[6],
        normalizedType: row[7],
        normalizedManufacturer: row[8],
        normalizedSupplier: row[9],
      });
    });
  }

  const meta = {
    sourceSpreadsheetId: MATERIAL_DB_APP.sourceSpreadsheetId,
    sourceSheets: MATERIAL_DB_APP.sourceSheetNames.slice(),
    updatedAt: '',
    recordCount: records.length,
  };

  const metaLastRow = metaSheet.getLastRow();
  if (metaLastRow >= 2) {
    const metaRows = metaSheet.getRange(2, 1, metaLastRow - 1, 2).getValues();
    metaRows.forEach((row) => {
      const key = row[0];
      const value = row[1];
      if (key === 'sourceSpreadsheetId') {
        meta.sourceSpreadsheetId = value || MATERIAL_DB_APP.sourceSpreadsheetId;
      } else if (key === 'sourceSheetsJson') {
        try {
          const parsed = JSON.parse(value);
          if (Array.isArray(parsed) && parsed.length) {
            meta.sourceSheets = parsed;
          }
        } catch (error) {
          meta.sourceSheets = MATERIAL_DB_APP.sourceSheetNames.slice();
        }
      } else if (key === 'updatedAt') {
        meta.updatedAt = value || '';
      } else if (key === 'recordCount') {
        meta.recordCount = toInt_(value);
      }
    });
  }

  return { meta, records };
}

function buildMaterialSearchPayload_(snapshot) {
  const records = snapshot.records || [];
  const uniqueTypes = new Set();
  const uniqueManufacturers = new Set();
  const uniqueSuppliers = new Set();

  const payloadRecords = records.map((record) => {
    if (record.type) {
      uniqueTypes.add(record.type);
    }
    if (record.manufacturer) {
      uniqueManufacturers.add(record.manufacturer);
    }
    if (record.supplier) {
      uniqueSuppliers.add(record.supplier);
    }

    return {
      o: record.tag,
      fullTag: record.tag,
      lowNoSpace: record.normalizedTag,
      manufNoSpace: record.normalizedManufacturer,
      typeNoSpace: record.normalizedType,
      supplierNoSpace: record.normalizedSupplier,
      sourceSheet: record.sourceSheet,
      article: record.article,
      type: record.type,
      manufacturer: record.manufacturer,
      supplier: record.supplier,
    };
  });

  return {
    records: payloadRecords,
    uniqueTypes: Array.from(uniqueTypes).sort(materialLocaleSort_),
    uniqueManufacturers: Array.from(uniqueManufacturers).sort(materialLocaleSort_),
    uniqueSuppliers: Array.from(uniqueSuppliers).sort(materialLocaleSort_),
    meta: buildMaterialSummary_(snapshot),
  };
}

function buildMaterialSummary_(snapshot) {
  return {
    sourceSpreadsheetId: snapshot.meta.sourceSpreadsheetId || MATERIAL_DB_APP.sourceSpreadsheetId,
    sourceSheets: snapshot.meta.sourceSheets || MATERIAL_DB_APP.sourceSheetNames.slice(),
    updatedAt: snapshot.meta.updatedAt || '',
    recordCount: snapshot.meta.recordCount || (snapshot.records ? snapshot.records.length : 0),
  };
}

function cacheMaterialSnapshot_(snapshot) {
  const cache = CacheService.getDocumentCache();
  const serialized = JSON.stringify(snapshot);
  const chunkCount = Math.ceil(serialized.length / MATERIAL_DB_APP.cacheChunkSize) || 1;
  cache.put(
    `${MATERIAL_DB_APP.cacheKeyPrefix}:count`,
    String(chunkCount),
    MATERIAL_DB_APP.cacheTtlSeconds
  );

  for (let index = 0; index < chunkCount; index += 1) {
    const chunk = serialized.slice(
      index * MATERIAL_DB_APP.cacheChunkSize,
      (index + 1) * MATERIAL_DB_APP.cacheChunkSize
    );
    cache.put(
      `${MATERIAL_DB_APP.cacheKeyPrefix}:chunk:${index}`,
      chunk,
      MATERIAL_DB_APP.cacheTtlSeconds
    );
  }
}

function loadMaterialSnapshotFromCache_() {
  const cache = CacheService.getDocumentCache();
  const countValue = cache.get(`${MATERIAL_DB_APP.cacheKeyPrefix}:count`);
  const count = toInt_(countValue);
  if (!count) {
    return null;
  }

  let serialized = '';
  for (let index = 0; index < count; index += 1) {
    const chunk = cache.get(`${MATERIAL_DB_APP.cacheKeyPrefix}:chunk:${index}`);
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

function hideMaterialSheets_() {
  const ss = SpreadsheetApp.getActive();
  [MATERIAL_DB_APP.metaSheetName, MATERIAL_DB_APP.dataSheetName].forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.hideSheet();
    }
  });
}

function isMaterialSystemSheet_(sheetName) {
  return (
    sheetName === MATERIAL_DB_APP.metaSheetName ||
    sheetName === MATERIAL_DB_APP.dataSheetName
  );
}

function materialLocaleSort_(a, b) {
  return String(a || '').localeCompare(String(b || ''), undefined, {
    numeric: true,
    sensitivity: 'base',
  });
}
