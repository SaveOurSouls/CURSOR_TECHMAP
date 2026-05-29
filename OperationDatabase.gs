// ============================================================
//  OperationDatabase.gs — база техопераций
//  Зависимости: Config.gs, Utils.gs
// ============================================================

// ── Public API ───────────────────────────────────────────────

/**
 * Диагностика: показывает все заголовки столбцов в БД.ТЕР и значения L+/L-
 * для конкретного артикула. Запустить вручную из редактора скриптов.
 * Пример вызова: diagnoseTerSheet('sshl-002t-p0.2')
 */
function diagnoseTerSheet(searchArticle) {
  const srcSS  = SpreadsheetApp.openById(TECHOPS_DB_APP.sourceSpreadsheetId);
  const tab    = TECHOPS_DB_APP.tabs.ter;
  const sheet  = srcSS.getSheetByName(tab.sourceSheetName);
  if (!sheet) { Logger.log('Sheet not found: ' + tab.sourceSheetName); return; }

  const lastCol   = sheet.getLastColumn();
  const headerRow = sheet.getRange(tab.headerRowNumber, 1, 1, lastCol).getValues()[0];

  Logger.log('=== БД.ТЕР HEADERS ===');
  headerRow.forEach((h, i) => {
    if (String(h).trim()) Logger.log(`col ${i+1} (${columnLetter_(i)}): "${h}"  →  "${String(h).toLowerCase().replace(/\s+/g,' ').trim()}"`);
  });

  if (!searchArticle) return;

  Logger.log('\n=== ПОИСК: "' + searchArticle + '" (по всем колонкам) ===');
  const data   = sheet.getRange(tab.headerRowNumber + 1, 1, Math.min(sheet.getLastRow() - tab.headerRowNumber, 2000), lastCol).getValues();
  const artLow = searchArticle.toLowerCase().trim();

  // Find row where ANY cell exactly matches the article
  let found = null;
  let foundRowNum = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].some(c => String(c || '').trim().toLowerCase() === artLow)) {
      found = data[i]; foundRowNum = i + tab.headerRowNumber + 1; break;
    }
  }

  if (found) {
    Logger.log('Строка найдена — row ' + foundRowNum);
    // Find L+ and L- column indices
    const lpIdx = headerRow.findIndex(h => String(h).trim() === 'L+');
    const lmIdx = headerRow.findIndex(h => String(h).trim() === 'L-');
    Logger.log(`L+ (col ${lpIdx+1}): "${found[lpIdx]}"`);
    Logger.log(`L- (col ${lmIdx+1}): "${found[lmIdx]}"`);
    Logger.log('--- все значения ---');
    headerRow.forEach((h, i) => { if (String(h).trim()) Logger.log(`  [${columnLetter_(i)}] ${h}: "${found[i]}"`); });
  } else {
    Logger.log('Артикул не найден в первых 2000 строках');
    Logger.log('Первые 3 значения в кол.G: ' + data.slice(0,3).map(r => r[6]).join(' | '));
  }
}

/** Проверяет что реально хранится в _TC_TECHOPS_DB для конкретного артикула TER. */
function diagnoseDbRecord(searchArticle) {
  const ss        = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  const metaSheet = ss.getSheetByName(TECHOPS_DB_APP.metaSheetName);

  if (!dataSheet) { Logger.log('DB-лист не найден: ' + TECHOPS_DB_APP.dataSheetName); return; }

  // Schema version from meta
  if (metaSheet) {
    const meta = metaSheet.getDataRange().getValues();
    const sv = meta.find(r => r[0] === 'schemaVersion');
    Logger.log('schemaVersion в DB: ' + (sv ? sv[1] : 'не найдена'));
    Logger.log('schemaVersion в коде: ' + TECHOPS_DB_APP.schemaVersion);
  }

  const data   = dataSheet.getDataRange().getValues();
  const artLow = (searchArticle || '').toLowerCase().trim();
  Logger.log('Всего строк в DB: ' + (data.length - 1));

  let terCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== 'ter') continue;
    terCount++;
    if (String(data[i][10] || '').toLowerCase().trim() === artLow) {
      Logger.log('Найдена TER-запись (строка ' + (i + 1) + '):');
      Logger.log('  row[6]  terManufacturer : "' + data[i][6]  + '"');
      Logger.log('  row[10] terArticle      : "' + data[i][10] + '"');
      Logger.log('  row[11] terLPlus        : "' + data[i][11] + '"');
      Logger.log('  row[12] terLMinus       : "' + data[i][12] + '"');
      return;
    }
  }
  Logger.log('TER-записей всего: ' + terCount);
  Logger.log('Артикул "' + searchArticle + '" в DB не найден');
}

function columnLetter_(idx) {
  let s = ''; let n = idx;
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
  return s;
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
  getTechOpsCache_().clear();

  const snapshot = fetchTechOperationsSnapshotFromSource_();
  writeTechOperationsSnapshotToSheets_(snapshot);
  getTechOpsCache_().save(snapshot);
  return buildTechOperationsSummary_(snapshot);
}

function getTechOperationsDatabase(forceRefresh) {
  if (forceRefresh) {
    syncTechOperationsDatabase();
    return buildTechOperationsPayload_(getTechOpsCache_().load() || loadTechOperationsSnapshotFromSheets_());
  }

  const cached = getTechOpsCache_().load();
  if (cached && cached.records && cached.records.length &&
      String(cached.meta && cached.meta.schemaVersion) === String(TECHOPS_DB_APP.schemaVersion)) {
    return buildTechOperationsPayload_(cached);
  }

  ensureTechOperationsInfrastructure_(SpreadsheetApp.getActive());
  const stored = loadTechOperationsSnapshotFromSheets_();
  if (!stored.records.length ||
      String(stored.meta.schemaVersion) !== String(TECHOPS_DB_APP.schemaVersion)) {
    syncTechOperationsDatabase();
    return buildTechOperationsPayload_(getTechOpsCache_().load() || loadTechOperationsSnapshotFromSheets_());
  }
  getTechOpsCache_().save(stored);
  return buildTechOperationsPayload_(stored);
}

/** Backward-compatible alias used by the workspace sidebar. */
function getOperationDatabase(forceRefresh) {
  return getTechOperationsDatabase(forceRefresh);
}

function insertTechOperationMatrix(matrix, targetCellA1) {
  if (!matrix || !matrix.length) return 'Нет данных для выгрузки';

  const width = Array.isArray(matrix[0]) ? matrix[0].length : 0;
  if (!width) return 'Нет данных для выгрузки';

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

  if (!startCell) return 'ОШИБКА: Не выбрана стартовая ячейка.';

  const startRow = startCell.getRow();
  const startCol = startCell.getColumn();
  const numRows  = matrix.length;
  const writeEnd = startCol + width - 1;

  ensureSheetCapacity_(sheet, startRow + numRows - 1, writeEnd + 20);

  // Сканируем от col 1 — getMergedRanges() может пропустить мёрджи, чья верхняя
  // левая ячейка левее startCol но тело перекрывает зону записи.
  const scanWidth = writeEnd + 20;
  const mergedRanges = sheet
    .getRange(startRow, 1, numRows, scanWidth)
    .getMergedRanges();

  const writeAreaHasMerges = mergedRanges.some((mr) => {
    const mStart = mr.getColumn();
    const mEnd   = mStart + mr.getNumColumns() - 1;
    return mEnd >= startCol && mStart <= writeEnd;
  });

  if (!writeAreaHasMerges) {
    sheet.getRange(startRow, startCol, numRows, width).setValues(matrix);
    sheet.getRange(startRow + numRows, startCol).activate();
    return `Успешно выгружено ${numRows} строк.`;
  }

  const mergeByRow = {};
  mergedRanges.forEach((mr) => {
    const r = mr.getRow();
    const c = mr.getColumn();
    const w = mr.getNumColumns();
    if (!mergeByRow[r]) mergeByRow[r] = {};
    mergeByRow[r][c] = w;
    for (let g = 1; g < w; g++) mergeByRow[r][c + g] = 0;
  });

  const allSegs = matrix.map((rowData, ri) => {
    const absRow = startRow + ri;
    const rowMerges = mergeByRow[absRow] || {};
    const segs = [];
    let dataIdx = 0, absCol = startCol, segStart = -1, segVals = [];

    const flushNormal = () => {
      if (segStart >= 0 && segVals.length) {
        segs.push({ type: 'n', col: segStart, len: segVals.length, vals: segVals.slice() });
        segStart = -1; segVals = [];
      }
    };

    while (dataIdx < rowData.length) {
      const mw = rowMerges[absCol];
      if (mw === 0) {
        flushNormal(); absCol++;
      } else if (mw >= 2) {
        flushNormal();
        segs.push({ type: 'm', col: absCol, val: rowData[dataIdx++] });
        absCol += mw;
      } else {
        if (segStart < 0) segStart = absCol;
        segVals.push(rowData[dataIdx++]);
        absCol++;
      }
    }
    flushNormal();
    return segs;
  });

  const tmpl = allSegs[0];
  const uniform = numRows === 1 || allSegs.every((segs) =>
    segs.length === tmpl.length &&
    segs.every((s, i) => s.type === tmpl[i].type && s.col === tmpl[i].col &&
      (s.type === 'n' ? s.len === tmpl[i].len : true))
  );

  if (uniform) {
    tmpl.forEach((seg, si) => {
      if (seg.type === 'n') {
        sheet.getRange(startRow, seg.col, numRows, seg.len)
          .setValues(allSegs.map((rowSegs) => rowSegs[si].vals));
      } else {
        allSegs.forEach((rowSegs, ri) => {
          sheet.getRange(startRow + ri, seg.col).setValue(rowSegs[si].val);
        });
      }
    });
  } else {
    allSegs.forEach((segs, ri) => {
      const absRow = startRow + ri;
      segs.forEach((seg) => {
        if (seg.type === 'n') {
          sheet.getRange(absRow, seg.col, 1, seg.len).setValues([seg.vals]);
        } else {
          sheet.getRange(absRow, seg.col).setValue(seg.val);
        }
      });
    });
  }

  sheet.getRange(startRow + numRows, startCol).activate();
  return `Успешно выгружено ${numRows} строк.`;
}

/** Backward-compatible alias used by the workspace sidebar. */
function insertOperationRows(matrix, targetCellA1) {
  return { message: insertTechOperationMatrix(matrix, targetCellA1) };
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
  if (!targetCell) return { ok: false, message: 'ОШИБКА: Не выбрана стартовая ячейка.' };
  targetCell.setValue(text);
  return { ok: true, message: 'Записано в ' + targetCell.getA1Notation() };
}

// ── Cache ────────────────────────────────────────────────────

function getTechOpsCache_() {
  return ChunkCache_(TECHOPS_DB_APP.cacheKeyPrefix, TECHOPS_DB_APP.cacheChunkSize, TECHOPS_DB_APP.cacheTtlSeconds);
}

// ── Infrastructure ───────────────────────────────────────────

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

// ── Snapshot fetch from source ───────────────────────────────

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
    if (!sheet) return;
    diagnosticsByTab[tabKey].sheetFound = true;

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return;

    const headerRowIndex = detectTechOperationsHeaderRow_(values, tabKey);
    if (headerRowIndex < 0) return;
    diagnosticsByTab[tabKey].headerRowNumber = headerRowIndex + 1;

    const headerMap = buildTechOperationsHeaderMap_(values[headerRowIndex]);
    const namedColumns = values[headerRowIndex]
      .map((h, i) => ({ name: String(h || '').trim(), index: i }))
      .filter(({ name }) => name !== '');
    columnHeadersByTab[tabKey] = namedColumns.map(({ name }) => name);
    const diagnostics = buildTechOperationsHeaderDiagnostics_(tabKey, headerMap, values[headerRowIndex]);
    diagnosticsByTab[tabKey].foundHeaders  = diagnostics.foundHeaders;
    diagnosticsByTab[tabKey].matchedGroups = diagnostics.matchedGroups;
    diagnosticsByTab[tabKey].missingGroups = diagnostics.missingGroups;

    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const record = buildTechOperationsRecordFromRow_(tabKey, row, headerMap, config.sourceSheetName, namedColumns);
      if (!record || !record.displayText) continue;
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

// ── Header parsing ───────────────────────────────────────────

function buildTechOperationsHeaderMap_(headersRow) {
  const map = {};
  headersRow.forEach((header, index) => {
    const normalized = normalizeHeader_(header);
    if (normalized) map[normalized] = index;
  });
  return map;
}

function detectTechOperationsHeaderRow_(values, tabKey) {
  const config = TECHOPS_DB_APP.tabs[tabKey];
  if (config && config.headerRowNumber) {
    const explicitIndex = Number(config.headerRowNumber) - 1;
    if (explicitIndex >= 0 && explicitIndex < values.length) return explicitIndex;
  }

  const aliases = getTechOperationsHeaderAliasesForTab_(tabKey);
  const maxScanRows = Math.min(values.length, 12);
  let bestIndex = -1;
  let bestScore = -1;

  for (let rowIndex = 0; rowIndex < maxScanRows; rowIndex += 1) {
    const headerMap = buildTechOperationsHeaderMap_(values[rowIndex]);
    let score = 0;
    aliases.forEach((aliasGroup) => {
      if (aliasGroup.some((alias) => {
        const key = normalizeHeader_(alias);
        return headerMap[key] === 0 || headerMap[key] > 0;
      })) score += 1;
    });
    if (score > bestScore) { bestScore = score; bestIndex = rowIndex; }
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
  const foundHeaders = (headersRow || []).map((h) => normalizeString_(h)).filter(Boolean);
  const matchedGroups = [];
  const missingGroups = [];

  aliasGroups.forEach((group) => {
    const matched = group.some((alias) => {
      const key = normalizeHeader_(alias);
      return headerMap[key] === 0 || headerMap[key] > 0;
    });
    if (matched) matchedGroups.push(group[0]);
    else missingGroups.push(group[0]);
  });

  return { foundHeaders, matchedGroups, missingGroups };
}

// ── Record builders ──────────────────────────────────────────

function buildTechOperationsRecordFromRow_(tabKey, row, headerMap, sourceSheet, namedColumns) {
  switch (tabKey) {
    case 'ob':   return buildTechOperationsObRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'op':   return buildTechOperationsOpRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'ter':  return buildTechOperationsTerRecord_(row, headerMap, sourceSheet, namedColumns);
    case 'coax': return buildTechOperationsCoaxRecord_(row, headerMap, sourceSheet, namedColumns);
    default:     return null;
  }
}

function buildTechOperationsObRecord_(row, headerMap, sourceSheet, namedColumns) {
  const baseValue = getTechOperationsCellByAliases_(row, headerMap, ['для базы', 'длябазы']);
  if (!baseValue) return null;
  const obType = getTechOperationsCellByAliases_(row, headerMap, ['тип', 'type', 'категория', 'category', 'группа']);
  return {
    tabKey: 'ob',
    displayText: baseValue,
    normalizedSearch: normalizeSearch_(baseValue),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
    obType: obType || '',
    sortKey: obType ? `${obType} ${baseValue}` : baseValue,
  };
}

function buildTechOperationsOpRecord_(row, headerMap, sourceSheet, namedColumns) {
  const number = getTechOperationsCellByAliases_(row, headerMap, ['номер', 'number']);
  const name   = getTechOperationsCellByAliases_(row, headerMap, ['название', 'name']);
  const tOp    = getTechOperationsCellByAliases_(row, headerMap, ['время операции', 'время операции, сек', 'время операции сек']);
  const tPrep  = getTechOperationsCellByAliases_(row, headerMap, ['время подготовки, сек', 'время подготовки сек', 'время подготовки']);
  const tMach  = getTechOperationsCellByAliases_(row, headerMap, ['время машины, сек/оп; сек/м', 'время машины сек/оп; сек/м', 'время машины']);

  const displayText = joinTechOperationsParts_([name, number], ' | ');
  if (!displayText) return null;

  return {
    tabKey: 'op',
    displayText,
    sortKey: name || number,
    normalizedSearch: normalizeSearch_(number + ' ' + name),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
    opNumber: number,
    opName:   name,
    tOp:      tOp   || '',
    tPrep:    tPrep || '',
    tMachine: tMach || '',
  };
}

function buildTechOperationsTerRecord_(row, headerMap, sourceSheet, namedColumns) {
  const manufacturer = getTechOperationsCellByAliases_(row, headerMap, ['производитель', 'бренд', 'manufacturer']);
  const series       = getTechOperationsCellByAliases_(row, headerMap, ['series', 'серия разъемов', 'серия']);
  const productName  = getTechOperationsCellByAliases_(row, headerMap, ['product name', 'productname', 'комплектующая']);
  const connType     = getTechOperationsCellByAliases_(row, headerMap, ['тип разъёма', 'тип разъема']);
  const terType      = getTechOperationsCellByAliases_(row, headerMap, ['тип контакта', 'тип конт.', 'тип конт']);
  const artISL       = getTechOperationsCellByAliases_(row, headerMap, ['артикул (контакта isl)', 'артикул контакта isl']);
  const artSAG       = getTechOperationsCellByAliases_(row, headerMap, ['артикул (контакт sag)', 'артикул контакт sag']);
  const terArticle   = getTechOperationsCellByAliases_(row, headerMap, ['артикул контакта (reel)', 'артикул контакта', 'артикул']);
  const lPlus  = getTechOperationsCellByAliases_(row, headerMap, ['l+', 'l+ в мм', 'l+(мм)', 'l +'])
              || getTechOperationsCellByHeaderRegex_(row, headerMap, /^l\s*\+/);
  const lMinus = getTechOperationsCellByAliases_(row, headerMap, ['l-', 'l−', 'l–', 'l—', 'l- в мм', 'l-(мм)', 'l −', 'l -'])
              || getTechOperationsCellByHeaderRegex_(row, headerMap, /^l\s*[-−–—]/);
  const applicator   = getTechOperationsCellByAliases_(row, headerMap, ['аппликатор', 'applicator', 'applikator']);
  const crimpHeight  = getTechOperationsCellByAliases_(row, headerMap, [
    'высота обжима проводника , мм', 'высота обжима проводника, мм',
    'высота обжима проводника', 'crimp height conductor', 'crimp height',
  ]);
  const pullForceMin = getTechOperationsCellByAliases_(row, headerMap, [
    'усилие обрыва контакта от, n', 'усилие обрыва контакта от n',
    'усилие обрыва от, n', 'усилие обрыва от', 'pull force min', 'pull test min', 'pull-off force min',
  ]);
  const pullForceMax = getTechOperationsCellByAliases_(row, headerMap, [
    'усилие обрыва контакта до, n', 'усилие обрыва контакта до n',
    'усилие обрыва до, n', 'усилие обрыва до', 'pull force max', 'pull test max', 'pull-off force max',
  ]);

  const displayText = joinTechOperationsParts_([manufacturer, series, productName], ' | ');
  if (!displayText) return null;

  return {
    tabKey: 'ter',
    displayText,
    terManufacturer:  manufacturer,
    terSeries:        series,
    terComponent:     productName,
    terType,
    terArticle,
    terLPlus:         lPlus        || '',
    terLMinus:        lMinus       || '',
    terApplicator:    applicator   || '',
    terCrimpHeight:   crimpHeight  || '',
    terPullForceMin:  pullForceMin || '',
    terPullForceMax:  pullForceMax || '',
    normalizedSearch: normalizeSearch_(
      [manufacturer, series, productName, connType, terType, artISL, artSAG].join(' ')
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
  if (!displayText) return null;

  return {
    tabKey: 'coax',
    displayText,
    coaxWire:    wire,
    coaxType:    typeSeries,
    coaxMfr:     mfr,
    coaxArticle: article,
    sortKey: `${wire} ${typeSeries} ${mfr} ${article}`,
    normalizedSearch: normalizeSearch_(
      [article, typeSeries, mfr, supplier, wire, program].join(' ')
    ),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
  };
}

function getTechOperationsCellByHeaderRegex_(row, headerMap, pattern) {
  for (const [key, idx] of Object.entries(headerMap)) {
    if (pattern.test(key)) return normalizeString_(row[idx]);
  }
  return '';
}

function getTechOperationsCellByAliases_(row, headerMap, aliases) {
  for (let index = 0; index < aliases.length; index += 1) {
    const headerIndex = headerMap[normalizeHeader_(aliases[index])];
    if (headerIndex === 0 || headerIndex > 0) {
      return normalizeString_(row[headerIndex]);
    }
  }
  return '';
}

function joinTechOperationsParts_(parts, delimiter) {
  return (parts || []).filter((part) => normalizeString_(part)).join(delimiter || ' | ');
}

// ── Snapshot persistence ─────────────────────────────────────

/**
 * Десериализует запись из строки _TC_OP_DATA с учётом типа (tabKey).
 * Устраняет мультиплексирование col 6-10, где раньше один столбец
 * использовался для нескольких полей разных вкладок одновременно.
 */
function parseTechOpsRow_(row) {
  const tabKey = row[0];
  const base = {
    tabKey,
    displayText:      row[1],
    normalizedSearch: row[2],
    exportValues:     parseJsonArray_(row[3]),
    sourceSheet:      row[4],
    sortKey:          row[5] || '',
  };
  // Extra columns [6..16] are written by toSnapshotRow_() — positions match exactly.
  if (tabKey === 'op')   return { ...base, opNumber: row[6] || '', opName: row[7] || '', tOp: row[8] || '', tPrep: row[9] || '', tMachine: row[10] || '' };
  if (tabKey === 'ter')  return { ...base, terManufacturer: row[6] || '', terSeries: row[7] || '', terComponent: row[8] || '', terType: row[9] || '', terArticle: row[10] || '', terLPlus: row[11] || '', terLMinus: row[12] || '', terApplicator: row[13] || '', terCrimpHeight: row[14] || '', terPullForceMin: row[15] || '', terPullForceMax: row[16] || '' };
  if (tabKey === 'coax') return { ...base, coaxWire: row[9] || '', coaxType: row[10] || '', coaxMfr: row[11] || '', coaxArticle: row[12] || '' };
  if (tabKey === 'ob')   return { ...base, obType: row[6] || '' };
  return base;
}

// Returns an array of 11 extra-column values for a snapshot record (col indices 6-16).
// Keeps each tabKey's fields at fixed positions and makes parseTechOpsRow_ the single
// source of truth for the layout — changing one here changes both read and write paths.
function toSnapshotRow_(record) {
  switch (record.tabKey) {
    case 'op':   return [record.opNumber || '', record.opName || '', record.tOp || '', record.tPrep || '', record.tMachine || '', '', '', '', '', '', ''];
    case 'ter':  return [record.terManufacturer || '', record.terSeries || '', record.terComponent || '', record.terType || '', record.terArticle || '', record.terLPlus || '', record.terLMinus || '', record.terApplicator || '', record.terCrimpHeight || '', record.terPullForceMin || '', record.terPullForceMax || ''];
    case 'coax': return ['', '', '', record.coaxWire || '', record.coaxType || '', record.coaxMfr || '', record.coaxArticle || '', '', '', '', ''];
    case 'ob':   return [record.obType || '', '', '', '', '', '', '', '', '', '', ''];
    default:     return ['', '', '', '', '', '', '', '', '', '', ''];
  }
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
      ...toSnapshotRow_(record),
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
        .forEach((row) => records.push(parseTechOpsRow_(row)));
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
        if      (key === 'sourceSpreadsheetId') meta.sourceSpreadsheetId = value || TECHOPS_DB_APP.sourceSpreadsheetId;
        else if (key === 'updatedAt')           meta.updatedAt = value || '';
        else if (key === 'recordCount')         meta.recordCount = toInt_(value);
        else if (key === 'schemaVersion')       meta.schemaVersion = toInt_(value);
        else if (key === 'countsByTabJson')     try { meta.countsByTab = JSON.parse(value) || {}; } catch (e) {}
        else if (key === 'diagnosticsByTabJson') try { meta.diagnosticsByTab = JSON.parse(value) || {}; } catch (e) {}
        else if (key === 'columnHeadersByTabJson') try { meta.columnHeadersByTab = JSON.parse(value) || {}; } catch (e) {}
      });
    }
  }

  return { meta, records };
}

// ── Payload builder ──────────────────────────────────────────

function buildTechOperationsPayload_(snapshot) {
  const payload = { meta: buildTechOperationsSummary_(snapshot), tabs: {}, dbOb: [], dbOp: [], dbTer: [], dbKoax: [] };

  const payloadKeyMap = { ob: 'dbOb', op: 'dbOp', ter: 'dbTer', coax: 'dbKoax' };

  TECHOPS_DB_APP.tabOrder.forEach((tabKey) => {
    const config = TECHOPS_DB_APP.tabs[tabKey];
    const items = (snapshot.records || [])
      .filter((record) => {
        if (record.tabKey !== tabKey) return false;
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
          item.tOp      = record.tOp      || '';
          item.tPrep    = record.tPrep    || '';
          item.tMachine = record.tMachine || '';
        }
        if (tabKey === 'ob') {
          item.obType = record.obType || '';
        }
        if (tabKey === 'ter') {
          const exp = record.exportValues || [];
          item.terComponent    = record.terComponent    || exp[0] || '';
          item.terSeries       = record.terSeries       || exp[2] || '';
          item.terManufacturer = record.terManufacturer || exp[3] || '';
          item.terType         = record.terType         || '';
          item.terArticle      = record.terArticle      || '';
          item.terLPlus        = record.terLPlus        || '';
          item.terLMinus       = record.terLMinus       || '';
          item.terApplicator   = record.terApplicator   || '';
          item.terCrimpHeight  = record.terCrimpHeight  || '';
          item.terPullForceMin = record.terPullForceMin || '';
          item.terPullForceMax = record.terPullForceMax || '';
        }
        if (tabKey === 'coax') {
          item.coaxWire    = record.coaxWire    || '';
          item.coaxType    = record.coaxType    || '';
          item.coaxMfr     = record.coaxMfr     || '';
          item.coaxArticle = record.coaxArticle || '';
          item.label = item.coaxArticle ||
            joinTechOperationsParts_([item.coaxType, item.coaxWire], ' | ');
          item.sortKey = `${item.coaxWire} ${item.coaxType} ${item.coaxMfr} ${item.coaxArticle}`;
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
    updatedAt:    snapshot.meta.updatedAt || '',
    recordCount:  snapshot.meta.recordCount || (snapshot.records ? snapshot.records.length : 0),
    countsByTab:  snapshot.meta.countsByTab || {},
    diagnosticsByTab: snapshot.meta.diagnosticsByTab || {},
  };
}

function hideTechOperationsSheets_() {
  const ss = SpreadsheetApp.getActive();
  [TECHOPS_DB_APP.metaSheetName, TECHOPS_DB_APP.dataSheetName].forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) sheet.hideSheet();
  });
}

/** Возвращает snapshot из кеша или листов; используется AssemblyGenerator. */
function getTechOperationsSnapshot_() {
  const cached = getTechOpsCache_().load();
  if (cached && cached.records && cached.records.length) return cached;
  const stored = loadTechOperationsSnapshotFromSheets_();
  if (stored.records.length) getTechOpsCache_().save(stored);
  return stored;
}
