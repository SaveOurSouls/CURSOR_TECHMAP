const TECHOPS_DB_APP = {
  sourceSpreadsheetId: '1W3VK9Fw71lYdw1Klcsn_za5-2EhvLoXIAKZVYOCnKcs',
  metaSheetName: '_TC_TECHOPS_META',
  dataSheetName: '_TC_TECHOPS_DB',
  metaHeaders: ['key', 'value'],
  dataHeaders: ['tabKey', 'displayText', 'normalizedSearch', 'exportJson', 'sourceSheet', 'sortKey', 'extra1', 'extra2', 'extra3', 'extra4', 'extra5', 'extra6', 'extra7'],
  cacheKeyPrefix: 'techmap-techops-db-v6',
  schemaVersion: 6,
  cacheChunkSize: 80000,
  cacheTtlSeconds: 21600,
  tabs: {
    ob: {
      key: 'ob',
      label: 'БД.ОБ',
      sourceSheetName: 'БД.ОБ',
      headerRowNumber: 2,
      searchPlaceholder: 'Поиск по полю "Для базы"...',
      outputLabels: ['Для базы'],
    },
    op: {
      key: 'op',
      label: 'БД.ОП',
      sourceSheetName: 'БД.ОП',
      headerRowNumber: 2,
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
      headerRowNumber: 1,
      searchPlaceholder: 'Поиск по производителю, серии, product name...',
      outputLabels: ['Тип', 'Производитель', 'Product Name', 'Series', 'Шаг', 'Тип конт.', 'Арт. ISL', 'Арт. SAG', 'Аппликатор'],
    },
    coax: {
      key: 'coax',
      label: 'БД.КОАКС',
      sourceSheetName: 'БД.КОАКС',
      headerRowNumber: 2,
      searchPlaceholder: 'Поиск по артикулам, сериям, проводу, размерам...',
      outputLabels: ['Артикул', 'Программа', 'D1', 'D2', 'D3', 'L1', 'L2', 'L3', 'L+', 'L-'],
    },
  },
  tabOrder: ['ob', 'op', 'ter', 'coax'],
};


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
  clearTechOperationsCache_();

  const snapshot = fetchTechOperationsSnapshotFromSource_();
  writeTechOperationsSnapshotToSheets_(snapshot);
  cacheTechOperationsSnapshot_(snapshot);
  return buildTechOperationsSummary_(snapshot);
}

function clearTechOperationsCache_() {
  const cache = CacheService.getDocumentCache();
  const countValue = cache.get(`${TECHOPS_DB_APP.cacheKeyPrefix}:count`);
  const count = toInt_(countValue) || 0;
  cache.remove(`${TECHOPS_DB_APP.cacheKeyPrefix}:count`);
  for (let i = 0; i < count; i += 1) {
    cache.remove(`${TECHOPS_DB_APP.cacheKeyPrefix}:chunk:${i}`);
  }
}

function getTechOperationsDatabase(forceRefresh) {
  if (forceRefresh) {
    syncTechOperationsDatabase();
    return buildTechOperationsPayload_(loadTechOperationsSnapshotFromCache_() || loadTechOperationsSnapshotFromSheets_());
  }

  // Fast path: cache warm and schema version matches — no sheet reads at all
  const cached = loadTechOperationsSnapshotFromCache_();
  if (cached && cached.records && cached.records.length &&
      String(cached.meta && cached.meta.schemaVersion) === String(TECHOPS_DB_APP.schemaVersion)) {
    return buildTechOperationsPayload_(cached);
  }

  // Cache miss or stale schema: load from sheets
  ensureTechOperationsInfrastructure_(SpreadsheetApp.getActive());
  const stored = loadTechOperationsSnapshotFromSheets_();
  if (!stored.records.length ||
      String(stored.meta.schemaVersion) !== String(TECHOPS_DB_APP.schemaVersion)) {
    syncTechOperationsDatabase();
    return buildTechOperationsPayload_(loadTechOperationsSnapshotFromCache_() || loadTechOperationsSnapshotFromSheets_());
  }
  cacheTechOperationsSnapshot_(stored);
  return buildTechOperationsPayload_(stored);
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

  const templateRow = startCell.getRow();
  const startCol = startCell.getColumn();
  const numRows = matrix.length;
  const maxCols = sheet.getMaxColumns();

  // ── Step 1: Read template row merge structure BEFORE inserting rows ──
  // After insertRowsAfter the indices shift, so capture merges first.
  const templateMerges = sheet
    .getRange(templateRow, 1, 1, maxCols)
    .getMergedRanges()
    .filter((mr) => mr.getNumColumns() > 1)
    .map((mr) => ({ col: mr.getColumn(), numCols: mr.getNumColumns() }));

  // ── Step 2: Insert blank rows after template row ──
  sheet.insertRowsAfter(templateRow, numRows);
  SpreadsheetApp.flush();

  const templateRange = sheet.getRange(templateRow, 1, 1, maxCols);
  const newRowsRange  = sheet.getRange(templateRow + 1, 1, numRows, maxCols);

  // ── Step 3: Copy formatting (colours, borders, fonts) ──
  templateRange.copyTo(newRowsRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  SpreadsheetApp.flush();

  // ── Step 4: Apply merges to each new row explicitly ──
  // GAS doesn't propagate merged-cell structure through copyTo reliably,
  // so we re-create each merge from the saved template map.
  if (templateMerges.length) {
    for (let i = 0; i < numRows; i++) {
      const newRow = templateRow + 1 + i;
      templateMerges.forEach(({ col, numCols }) => {
        sheet.getRange(newRow, col, 1, numCols).merge();
      });
    }
    SpreadsheetApp.flush();
  }

  // ── Step 3: Write matrix data into the new rows ──
  const writeStartRow = templateRow + 1;
  const scanCols = width + 20;

  const mergedRanges = sheet
    .getRange(writeStartRow, startCol, numRows, scanCols)
    .getMergedRanges();

  if (!mergedRanges.length) {
    // Fast path: no merged cells in the write area.
    sheet.getRange(writeStartRow, startCol, numRows, width).setValues(matrix);
  } else {
    // Build merge map: absRow → { absCol → mergeWidth | 0 (ghost) }
    const mergeByRow = {};
    mergedRanges.forEach((mr) => {
      const r = mr.getRow();
      const c = mr.getColumn();
      const w = mr.getNumColumns();
      if (!mergeByRow[r]) mergeByRow[r] = {};
      mergeByRow[r][c] = w;
      for (let g = 1; g < w; g++) mergeByRow[r][c + g] = 0;
    });

    // Write row by row: buffer normal cells into segments (setValues),
    // write merge top-left cells individually (setValue), skip ghost cells.
    for (let r = 0; r < numRows; r++) {
      const absRow   = writeStartRow + r;
      const rowMerges = mergeByRow[absRow] || {};
      const rowData   = matrix[r];

      let dataIdx = 0;
      let absCol  = startCol;
      let segStart = -1;
      const segVals = [];

      const flushSeg = () => {
        if (segStart >= 0 && segVals.length) {
          sheet.getRange(absRow, segStart, 1, segVals.length).setValues([segVals.slice()]);
          segStart = -1;
          segVals.length = 0;
        }
      };

      while (dataIdx < rowData.length) {
        const mw = rowMerges[absCol];
        if (mw === 0) {
          flushSeg();
          absCol++;
        } else if (mw >= 2) {
          flushSeg();
          sheet.getRange(absRow, absCol, 1, 1).setValue(rowData[dataIdx++]);
          absCol += mw;
        } else {
          if (segStart < 0) segStart = absCol;
          segVals.push(rowData[dataIdx++]);
          absCol++;
        }
      }
      flushSeg();
    }
  }

  sheet.getRange(writeStartRow + numRows, startCol).activate();
  return `Успешно выгружено ${numRows} строк.`;
}

// Backward-compatible alias used by the workspace sidebar.
function insertOperationRows(matrix, targetCellA1) {
  return {
    message: insertTechOperationMatrix(matrix, targetCellA1),
  };
}

function writeSingleCellNames(text, targetCellA1) {
  if (text === null || text === undefined) {
    return { ok: false, message: 'Нет данных для записи.' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  let targetCell;
  if (targetCellA1) {
    try {
      targetCell = sheet.getRange(targetCellA1);
    } catch (e) {
      return { ok: false, message: 'ОШИБКА: Неверный адрес ячейки (' + targetCellA1 + ').' };
    }
  } else {
    targetCell = ss.getCurrentCell() || sheet.getActiveCell();
  }
  if (!targetCell) {
    return { ok: false, message: 'ОШИБКА: Не выбрана стартовая ячейка.' };
  }
  targetCell.setValue(text);
  return { ok: true, message: 'Записано в ' + targetCell.getA1Notation() };
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
    ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.metaHeaders.length);
    sheet
      .getRange(1, 1, 1, TECHOPS_DB_APP.metaHeaders.length)
      .setValues([TECHOPS_DB_APP.metaHeaders])
      .setFontWeight('bold')
      .setBackground('#f3f6fc');
    sheet.hideSheet();
  }
  return sheet;
}

function ensureTechOperationsDataSheet_(ss) {
  let sheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHOPS_DB_APP.dataSheetName);
    ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.dataHeaders.length);
    sheet
      .getRange(1, 1, 1, TECHOPS_DB_APP.dataHeaders.length)
      .setValues([TECHOPS_DB_APP.dataHeaders])
      .setFontWeight('bold')
      .setBackground('#f3f6fc');
    sheet.hideSheet();
    return sheet;
  }

  // Schema check: rebuild only when column count changed
  const existingCols = sheet.getLastColumn();
  if (existingCols > 0 && existingCols !== TECHOPS_DB_APP.dataHeaders.length) {
    sheet.clear();
    if (existingCols < TECHOPS_DB_APP.dataHeaders.length) {
      sheet.insertColumnsAfter(existingCols, TECHOPS_DB_APP.dataHeaders.length - existingCols);
    } else {
      sheet.deleteColumns(TECHOPS_DB_APP.dataHeaders.length + 1, existingCols - TECHOPS_DB_APP.dataHeaders.length);
    }
    ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.dataHeaders.length);
    sheet
      .getRange(1, 1, 1, TECHOPS_DB_APP.dataHeaders.length)
      .setValues([TECHOPS_DB_APP.dataHeaders])
      .setFontWeight('bold')
      .setBackground('#f3f6fc');
    sheet.hideSheet();
  }

  return sheet;
}

function fetchTechOperationsSnapshotFromSource_() {
  const sourceSs = SpreadsheetApp.openById(TECHOPS_DB_APP.sourceSpreadsheetId);
  const records = [];
  const countsByTab = {};
  const diagnosticsByTab = {};
  const columnHeadersByTab = {};

  TECHOPS_DB_APP.tabOrder.forEach((tabKey) => {
    const config = TECHOPS_DB_APP.tabs[tabKey];
    countsByTab[tabKey] = 0;
    diagnosticsByTab[tabKey] = {
      label: config.label,
      sourceSheetName: config.sourceSheetName,
      headerRowNumber: '',
      foundHeaders: [],
      matchedGroups: [],
      missingGroups: [],
      parsedRows: 0,
      sheetFound: false,
    };
    const sheet = sourceSs.getSheetByName(config.sourceSheetName);
    if (!sheet) {
      return;
    }
    diagnosticsByTab[tabKey].sheetFound = true;

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) {
      return;
    }

    const headerRowIndex = detectTechOperationsHeaderRow_(values, tabKey);
    if (headerRowIndex < 0) {
      return;
    }
    diagnosticsByTab[tabKey].headerRowNumber = headerRowIndex + 1;

    const headerMap = buildTechOperationsHeaderMap_(values[headerRowIndex]);
    const namedColumns = values[headerRowIndex]
      .map((h, i) => ({ name: String(h || '').trim(), index: i }))
      .filter(({ name }) => name !== '');
    columnHeadersByTab[tabKey] = namedColumns.map(({ name }) => name);
    const diagnostics = buildTechOperationsHeaderDiagnostics_(tabKey, headerMap, values[headerRowIndex]);
    diagnosticsByTab[tabKey].foundHeaders = diagnostics.foundHeaders;
    diagnosticsByTab[tabKey].matchedGroups = diagnostics.matchedGroups;
    diagnosticsByTab[tabKey].missingGroups = diagnostics.missingGroups;
    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const record = buildTechOperationsRecordFromRow_(tabKey, row, headerMap, config.sourceSheetName, namedColumns);
      if (!record || !record.displayText) {
        continue;
      }
      countsByTab[tabKey] += 1;
      diagnosticsByTab[tabKey].parsedRows += 1;
      records.push(record);
    }
  });

  records.sort((a, b) => {
    if (a.tabKey !== b.tabKey) {
      return TECHOPS_DB_APP.tabOrder.indexOf(a.tabKey) - TECHOPS_DB_APP.tabOrder.indexOf(b.tabKey);
    }
    const aKey = String(a.sortKey || a.displayText || '');
    const bKey = String(b.sortKey || b.displayText || '');
    return aKey.localeCompare(bKey, undefined, { numeric: true, sensitivity: 'base' });
  });

  return {
    meta: {
      sourceSpreadsheetId: TECHOPS_DB_APP.sourceSpreadsheetId,
      updatedAt: new Date().toISOString(),
      recordCount: records.length,
      schemaVersion: TECHOPS_DB_APP.schemaVersion,
      countsByTab,
      diagnosticsByTab,
      columnHeadersByTab,
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
  const config = TECHOPS_DB_APP.tabs[tabKey];
  if (config && config.headerRowNumber) {
    const explicitIndex = Number(config.headerRowNumber) - 1;
    if (explicitIndex >= 0 && explicitIndex < values.length) {
      return explicitIndex;
    }
  }

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
        ['product name', 'productname', 'комплектующая'],
        ['series', 'серия разъемов', 'серия'],
        ['производитель', 'бренд', 'manufacturer'],
        ['тип разъёма', 'тип разъема'],
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

function buildTechOperationsHeaderDiagnostics_(tabKey, headerMap, headersRow) {
  const aliasGroups = getTechOperationsHeaderAliasesForTab_(tabKey);
  const foundHeaders = (headersRow || [])
    .map((header) => normalizeString_(header))
    .filter(Boolean);
  const matchedGroups = [];
  const missingGroups = [];

  aliasGroups.forEach((group) => {
    const matchedAlias = group.find((alias) => {
      const key = normalizeTechOperationsHeader_(alias);
      return headerMap[key] === 0 || headerMap[key] > 0;
    });
    if (matchedAlias) {
      matchedGroups.push(group[0]);
    } else {
      missingGroups.push(group[0]);
    }
  });

  return {
    foundHeaders,
    matchedGroups,
    missingGroups,
  };
}

function buildTechOperationsRecordFromRow_(tabKey, row, headerMap, sourceSheet, namedColumns) {
  switch (tabKey) {
    case 'ob':
      return buildTechOperationsObRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'op':
      return buildTechOperationsOpRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'ter':
      return buildTechOperationsTerRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'coax':
      return buildTechOperationsCoaxRecord_(row, headerMap, sourceSheet, namedColumns);
    default:
      return null;
  }
}

function buildTechOperationsObRecord_(row, headerMap, sourceSheet, namedColumns) {
  const baseValue = getTechOperationsCellByAliases_(row, headerMap, ['для базы', 'длябазы']);
  if (!baseValue) {
    return null;
  }
  const obType = getTechOperationsCellByAliases_(row, headerMap, ['тип', 'type', 'категория', 'category', 'группа']);
  return {
    tabKey: 'ob',
    displayText: baseValue,
    normalizedSearch: normalizeTechOperationsSearch_(baseValue),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
    obType: obType || '',
    sortKey: obType ? `${obType} ${baseValue}` : baseValue,
  };
}

function buildTechOperationsOpRecord_(row, headerMap, sourceSheet, namedColumns) {
  const number = getTechOperationsCellByAliases_(row, headerMap, ['номер', 'number']);
  const name = getTechOperationsCellByAliases_(row, headerMap, ['название', 'name']);

  // Display as "Название | Номер" so list is sorted and shown by name first
  const displayText = joinTechOperationsParts_([name, number], ' | ');
  if (!displayText) {
    return null;
  }

  return {
    tabKey: 'op',
    displayText,
    sortKey: name || number,
    normalizedSearch: normalizeTechOperationsSearch_(number + ' ' + name),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
    opNumber: number,
    opName: name,
  };
}

function buildTechOperationsTerRecord_(row, headerMap, sourceSheet, namedColumns) {
  const manufacturer = getTechOperationsCellByAliases_(row, headerMap, ['производитель', 'бренд', 'manufacturer']);
  const series       = getTechOperationsCellByAliases_(row, headerMap, ['series', 'серия разъемов', 'серия']);
  const productName  = getTechOperationsCellByAliases_(row, headerMap, ['product name', 'productname', 'комплектующая']);
  const connType     = getTechOperationsCellByAliases_(row, headerMap, ['тип разъёма', 'тип разъема']);
  const artISL       = getTechOperationsCellByAliases_(row, headerMap, ['артикул (контакта isl)', 'артикул контакта isl']);
  const artSAG       = getTechOperationsCellByAliases_(row, headerMap, ['артикул (контакт sag)', 'артикул контакт sag']);

  const displayText = joinTechOperationsParts_([manufacturer, series, productName], ' | ');
  if (!displayText) {
    return null;
  }

  return {
    tabKey: 'ter',
    displayText,
    terManufacturer: manufacturer,
    terSeries:       series,
    terComponent:    productName,
    normalizedSearch: normalizeTechOperationsSearch_(
      [manufacturer, series, productName, connType, artISL, artSAG].join(' ')
    ),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
  };
}

function buildTechOperationsCoaxRecord_(row, headerMap, sourceSheet, namedColumns) {
  const article    = getTechOperationsCellByAliases_(row, headerMap, ['артикул']);
  const typeSeries = getTechOperationsCellByAliases_(row, headerMap, ['тип/серия', 'тип / серия', 'тип серия']);
  const mfr        = getTechOperationsCellByAliases_(row, headerMap, ['производитель', 'бренд', 'manufacturer']);
  const supplier   = getTechOperationsCellByAliases_(row, headerMap, ['поставщик', 'supplier']);
  const wire       = getTechOperationsCellByAliases_(row, headerMap, ['провод']);
  const program    = getTechOperationsCellByAliases_(row, headerMap, ['программа']);

  const displayText = joinTechOperationsParts_([typeSeries, wire, supplier], ' | ');
  if (!displayText) {
    return null;
  }

  const exportValues = (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || '');

  return {
    tabKey: 'coax',
    displayText,
    coaxWire:     wire,
    coaxType:     typeSeries,
    coaxMfr:      mfr,
    coaxArticle:  article,
    sortKey: `${wire}\u0000${typeSeries}\u0000${mfr}\u0000${article}`,
    normalizedSearch: normalizeTechOperationsSearch_(
      [article, typeSeries, mfr, supplier, wire, program].join(' ')
    ),
    exportValues,
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
      record.sortKey || '',
      record.terManufacturer || record.opNumber || record.obType || '',
      record.terSeries       || record.opName   || '',
      record.terComponent    || '',
      record.coaxWire        || '',
      record.coaxType        || '',
      record.coaxMfr         || '',
      record.coaxArticle     || '',
    ]);
    dataSheet.getRange(2, 1, rows.length, TECHOPS_DB_APP.dataHeaders.length).setValues(rows);
  }

  metaSheet.clearContents();
  ensureSheetCapacity_(metaSheet, 9, TECHOPS_DB_APP.metaHeaders.length);
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
    ['diagnosticsByTabJson', JSON.stringify(snapshot.meta.diagnosticsByTab || {})],
    ['schemaVersion', String(TECHOPS_DB_APP.schemaVersion)],
    ['columnHeadersByTabJson', JSON.stringify(snapshot.meta.columnHeadersByTab || {})],
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
  const dataSheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  const metaSheet = ss.getSheetByName(TECHOPS_DB_APP.metaSheetName);

  const records = [];
  if (dataSheet) {
    const lastRow = dataSheet.getLastRow();
    if (lastRow >= 2) {
      dataSheet
        .getRange(2, 1, lastRow - 1, TECHOPS_DB_APP.dataHeaders.length)
        .getValues()
        .filter((row) => row[0] && row[1])
        .forEach((row) => {
          records.push({
            tabKey: row[0],
            displayText: row[1],
            normalizedSearch: row[2],
            exportValues: parseJsonArray_(row[3]),
            sourceSheet: row[4],
            sortKey: row[5] || '',
            terManufacturer: row[6] || '',
            opNumber:        row[6] || '',
            obType:          row[6] || '',
            terSeries:       row[7] || '',
            opName:          row[7] || '',
            terComponent:    row[8] || '',
            coaxWire:    row[9]  || '',
            coaxType:    row[10] || '',
            coaxMfr:     row[11] || '',
            coaxArticle: row[12] || '',
          });
        });
    }
  }

  const meta = {
    sourceSpreadsheetId: TECHOPS_DB_APP.sourceSpreadsheetId,
    updatedAt: '',
    recordCount: records.length,
    schemaVersion: 0,
    countsByTab: {},
    diagnosticsByTab: {},
    columnHeadersByTab: {},
  };

  if (metaSheet) {
    const metaLastRow = metaSheet.getLastRow();
    if (metaLastRow >= 2) {
      metaSheet.getRange(2, 1, metaLastRow - 1, 2).getValues().forEach((row) => {
        const key = row[0];
        const value = row[1];
        if (key === 'sourceSpreadsheetId') {
          meta.sourceSpreadsheetId = value || TECHOPS_DB_APP.sourceSpreadsheetId;
        } else if (key === 'updatedAt') {
          meta.updatedAt = value || '';
        } else if (key === 'recordCount') {
          meta.recordCount = toInt_(value);
        } else if (key === 'schemaVersion') {
          meta.schemaVersion = toInt_(value);
        } else if (key === 'countsByTabJson') {
          try { meta.countsByTab = JSON.parse(value) || {}; } catch (e) {}
        } else if (key === 'diagnosticsByTabJson') {
          try { meta.diagnosticsByTab = JSON.parse(value) || {}; } catch (e) {}
        } else if (key === 'columnHeadersByTabJson') {
          try { meta.columnHeadersByTab = JSON.parse(value) || {}; } catch (e) {}
        }
      });
    }
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
      .filter((record) => {
        if (record.tabKey !== tabKey) return false;
        // Skip separator/placeholder rows in БД.ОБ (no letters or digits)
        if (tabKey === 'ob' && !/[а-яёa-z0-9]/i.test(record.displayText || '')) return false;
        return true;
      })
      .map((record, index) => {
        const item = {
          id: `${tabKey}-${index}`,
          displayText: record.displayText,
          label: record.displayText,
          searchText: record.normalizedSearch,
          values: record.exportValues || [],
          outputRow: record.exportValues || [],
          sourceSheet: record.sourceSheet,
          sortKey: record.sortKey || '',
        };
        if (tabKey === 'op') {
          item.opNumber = record.opNumber || '';
          item.opName   = record.opName   || record.sortKey || '';
        }
        if (tabKey === 'ob') {
          item.obType = record.obType || '';
        }
        if (tabKey === 'ter') {
          const exp = record.exportValues || [];
          item.terComponent    = record.terComponent    || exp[0] || '';
          item.terSeries       = record.terSeries       || exp[2] || '';
          item.terManufacturer = record.terManufacturer || exp[3] || '';
        }
        if (tabKey === 'coax') {
          item.coaxWire    = record.coaxWire    || '';
          item.coaxType    = record.coaxType    || '';
          item.coaxMfr     = record.coaxMfr     || '';
          item.coaxArticle = record.coaxArticle || '';
          // Leaf label shown at level 4 (Артикул); if empty use Тип+Провод
          item.label = item.coaxArticle ||
            joinTechOperationsParts_([item.coaxType, item.coaxWire], ' | ');
          item.sortKey = `${item.coaxWire}\u0000${item.coaxType}\u0000${item.coaxMfr}\u0000${item.coaxArticle}`;
        }
        return item;
      });

    payload.tabs[tabKey] = {
      key: tabKey,
      label: config.label,
      searchPlaceholder: config.searchPlaceholder,
      outputLabels: config.outputLabels.slice(),
      columnHeaders: ((snapshot.meta && snapshot.meta.columnHeadersByTab) || {})[tabKey] || [],
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
    diagnosticsByTab: snapshot.meta.diagnosticsByTab || {},
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
