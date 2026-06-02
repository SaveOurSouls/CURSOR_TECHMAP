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
  const srcSS  = SpreadsheetApp.openById(getSourceSpreadsheetId_());
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

/**
 * Записывает данные терминала (L+, L−, шаг, аппликатор) в исходную БД.ТЕР
 * по артикулу и затем пересинхронизирует снимок.
 * fieldsJson = JSON.stringify({lPlus, lMinus, step, applicator})
 */
function saveTerDataToSourceDb(article, fieldsJson) {
  return withDocumentLock_(function() { return saveTerDataToSourceDbImpl_(article, fieldsJson); });
}

function saveTerDataToSourceDbImpl_(article, fieldsJson) {
  const fields = safeJsonParse_(fieldsJson, null);
  if (!fields || typeof fields !== 'object') {
    return { ok: false, message: 'Некорректные данные формы (битый JSON).' };
  }
  const srcSS   = SpreadsheetApp.openById(getSourceSpreadsheetId_());
  const tab     = TECHOPS_DB_APP.tabs.ter;
  const sheet   = srcSS.getSheetByName(tab.sourceSheetName);
  if (!sheet) throw new Error('Лист ' + tab.sourceSheetName + ' не найден');

  const data    = sheet.getDataRange().getValues();
  const headers = data[tab.headerRowNumber - 1].map(h => normalizeHeader_(h));

  const colIdx = aliases => {
    for (const a of aliases) {
      const i = headers.indexOf(normalizeHeader_(a));
      if (i >= 0) return i;
    }
    return headers.findIndex(h => aliases.some(a => h.startsWith(normalizeHeader_(a).split(' ')[0]) && h.includes(normalizeHeader_(a).split(' ').pop())));
  };

  const artCol  = colIdx(['артикул контакта (reel)', 'артикул контакта', 'артикул']);
  const lpCol   = colIdx(['l+', 'l+ в мм', 'l+(мм)']);
  const lmCol   = colIdx(['l-', 'l−', 'l– в мм', 'l-(мм)']);
  const stepCol = headers.findIndex(h => /^шаг/.test(h));
  const applCol = colIdx(['аппликатор', 'applicator']);

  if (artCol < 0) throw new Error('Колонка артикула не найдена в ' + tab.sourceSheetName);

  // 3-проходный матч ПО ВСЕМ строкам с приоритетом точного: exact → normalized
  // → без префикса-производителя. Точное совпадение всегда побеждает неточное,
  // даже если неточное встречается в более ранней строке (защита от записи не туда).
  const normArt = s => String(s).toLowerCase().replace(/[\s\-\.\/\(\)_]/g, '');
  const artNormExact = String(article).toLowerCase().trim();
  const artNorm      = normArt(article);
  const words        = String(article).trim().split(/\s+/);
  const artNormNoMfr = words.length > 1 ? normArt(words.slice(1).join(' ')) : '';

  let rowIdx = -1;
  let matchType = '';
  for (let i = tab.headerRowNumber; i < data.length; i++) {
    if (String(data[i][artCol] || '').toLowerCase().trim() === artNormExact) { rowIdx = i; matchType = 'exact'; break; }
  }
  if (rowIdx < 0) {
    for (let i = tab.headerRowNumber; i < data.length; i++) {
      if (normArt(data[i][artCol]) === artNorm) { rowIdx = i; matchType = 'normalized'; break; }
    }
  }
  if (rowIdx < 0 && artNormNoMfr.length >= 3) {
    for (let i = tab.headerRowNumber; i < data.length; i++) {
      if (normArt(data[i][artCol]) === artNormNoMfr) { rowIdx = i; matchType = 'noMfr'; break; }
    }
  }

  if (rowIdx < 0) {
    if (!fields._forceAdd) return { ok: false, notFound: true };
    const newRow = new Array(headers.length).fill('');
    newRow[artCol] = article;
    // Технические числа нормализуем к запятой; аппликатор — текст, как есть.
    if (fields.lPlus      !== undefined && lpCol   >= 0) newRow[lpCol]   = normalizeTechnicalDecimal_(fields.lPlus);
    if (fields.lMinus     !== undefined && lmCol   >= 0) newRow[lmCol]   = normalizeTechnicalDecimal_(fields.lMinus);
    if (fields.step       !== undefined && stepCol >= 0) newRow[stepCol] = normalizeTechnicalDecimal_(fields.step);
    if (fields.applicator !== undefined && applCol >= 0) newRow[applCol] = fields.applicator;
    sheet.appendRow(newRow);
    syncTechOperationsDatabase();
    return { ok: true, added: true };
  }

  // S3: совпадение НЕточное (по нормализации / без префикса-производителя) —
  // не пишем молча в потенциально чужую строку. Возвращаем найденный артикул
  // на подтверждение; клиент переспросит и повторит вызов с _confirmInexact=true.
  if (matchType !== 'exact' && !fields._confirmInexact) {
    return {
      ok: false,
      needsConfirm: true,
      matchedArticle:   String(data[rowIdx][artCol] || ''),
      requestedArticle: String(article || ''),
    };
  }

  // Точечная запись по колонкам (не вся строка): источник может содержать
  // формулы в соседних ячейках — батч-перезапись строки затёрла бы их статикой.
  // Технические числа нормализуем к запятой; аппликатор — текст, как есть.
  const sheetRow = rowIdx + 1;
  if (fields.lPlus      !== undefined && lpCol   >= 0) sheet.getRange(sheetRow, lpCol   + 1).setValue(normalizeTechnicalDecimal_(fields.lPlus));
  if (fields.lMinus     !== undefined && lmCol   >= 0) sheet.getRange(sheetRow, lmCol   + 1).setValue(normalizeTechnicalDecimal_(fields.lMinus));
  if (fields.step       !== undefined && stepCol >= 0) sheet.getRange(sheetRow, stepCol + 1).setValue(normalizeTechnicalDecimal_(fields.step));
  if (fields.applicator !== undefined && applCol >= 0) sheet.getRange(sheetRow, applCol + 1).setValue(fields.applicator);

  syncTechOperationsDatabase();
  return { ok: true, added: false };
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
  return withDocumentLock_(function() {
    const ss = SpreadsheetApp.getActive();
    ensureTechOperationsInfrastructure_(ss);
    getTechOpsCache_().clear();

    const snapshot = fetchTechOperationsSnapshotFromSource_();
    writeTechOperationsSnapshotToSheets_(snapshot);
    getTechOpsCache_().save(snapshot);
    return buildTechOperationsSummary_(snapshot);
  });
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
    writeSheetHeader_(sheet, TECHOPS_DB_APP.metaHeaders);
    sheet.hideSheet();
  }
  return sheet;
}

function ensureTechOperationsDataSheet_(ss) {
  let sheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHOPS_DB_APP.dataSheetName);
    ensureSheetCapacity_(sheet, 2, TECHOPS_DB_APP.dataHeaders.length);
    writeSheetHeader_(sheet, TECHOPS_DB_APP.dataHeaders);
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
    writeSheetHeader_(sheet, TECHOPS_DB_APP.dataHeaders);
    sheet.hideSheet();
  }

  return sheet;
}

// ── Snapshot fetch from source ───────────────────────────────

function fetchTechOperationsSnapshotFromSource_() {
  const sourceSs = SpreadsheetApp.openById(getSourceSpreadsheetId_());
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

    // Колонки полей резолвим один раз на лист (не построчно) — P3.
    const resolvedColumns = DB_SCHEMA[tabKey] ? resolveSchemaColumns_(DB_SCHEMA[tabKey], headerMap) : {};

    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const record = buildRecordFromSchema_(tabKey, row, resolvedColumns, config.sourceSheetName, namedColumns);
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
      sourceSpreadsheetId: getSourceSpreadsheetId_(),
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
      if (aliasGroup.some((alias) => hasHeader_(headerMap, normalizeHeader_(alias)))) score += 1;
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
    const matched = group.some((alias) => hasHeader_(headerMap, normalizeHeader_(alias)));
    if (matched) matchedGroups.push(group[0]);
    else missingGroups.push(group[0]);
  });

  return { foundHeaders, matchedGroups, missingGroups };
}

// ── Record builders (декларативная схема) ───────────────────
// ЕДИНЫЙ источник извлечения полей всех вкладок. Каждое поле описано алиасами
// заголовков (+ опц. regex-фолбэк), флагом store (писать ли в запись) и
// производными display/search/sortKey. Один buildRecordFromSchema_ заменяет
// четыре ручных билдера (Ob/Op/Ter/Coax). Форма записи сохранена 1:1 — снапшот
// и parseTechOpsRow_↔toSnapshotRow_ совместимы, schemaVersion не меняется.
// Колонки резолвятся ОДИН раз на лист (resolveSchemaColumns_) → не нормализуем
// алиасы построчно (узкое место P3).
const DB_SCHEMA = {
  ob: {
    fields: [
      { key: 'baseValue', aliases: ['для базы', 'длябазы'] },
      { key: 'obType',    aliases: ['тип', 'type', 'категория', 'category', 'группа'], store: true },
    ],
    display: v => v.baseValue,
    search:  v => v.baseValue,
    sortKey: v => (v.obType ? `${v.obType} ${v.baseValue}` : v.baseValue),
  },
  op: {
    fields: [
      { key: 'opNumber', aliases: ['номер', 'number'], store: true },
      { key: 'opName',   aliases: ['название', 'name'], store: true },
      { key: 'tOp',      aliases: ['время операции', 'время операции, сек', 'время операции сек'], store: true },
      { key: 'tPrep',    aliases: ['время подготовки, сек', 'время подготовки сек', 'время подготовки'], store: true },
      { key: 'tMachine', aliases: ['время машины, сек/оп; сек/м', 'время машины сек/оп; сек/м', 'время машины'], store: true },
    ],
    display: v => joinTechOperationsParts_([v.opName, v.opNumber], ' | '),
    sortKey: v => v.opName || v.opNumber,
    search:  v => v.opNumber + ' ' + v.opName,
  },
  ter: {
    fields: [
      { key: 'terManufacturer', aliases: ['производитель', 'бренд', 'manufacturer'], store: true },
      { key: 'terSeries',       aliases: ['series', 'серия разъемов', 'серия'], store: true },
      { key: 'terComponent',    aliases: ['product name', 'productname', 'комплектующая'], store: true },
      { key: 'connType',        aliases: ['тип разъёма', 'тип разъема'] },
      { key: 'terType',         aliases: ['тип контакта', 'тип конт.', 'тип конт'], store: true },
      { key: 'artISL',          aliases: ['артикул (контакта isl)', 'артикул контакта isl'] },
      { key: 'artSAG',          aliases: ['артикул (контакт sag)', 'артикул контакт sag'] },
      { key: 'terArticle',      aliases: ['артикул контакта (reel)', 'артикул контакта', 'артикул'], store: true },
      { key: 'terLPlus',        aliases: ['l+', 'l+ в мм', 'l+(мм)', 'l +'], regex: /^l\s*\+/, store: true },
      { key: 'terLMinus',       aliases: ['l-', 'l−', 'l–', 'l—', 'l- в мм', 'l-(мм)', 'l −', 'l -'], regex: /^l\s*[-−–—]/, store: true },
      { key: 'terStep',         aliases: ['шаг разъема', 'шаг разъёма', 'шаг', 'pitch', 'step', 'шаг ленты', 'шаг контакта'], regex: /^шаг/, store: true },
      { key: 'terApplicator',   aliases: ['аппликатор', 'applicator', 'applikator'], store: true },
      { key: 'terCrimpHeight',  aliases: ['высота обжима проводника , мм', 'высота обжима проводника, мм', 'высота обжима проводника', 'crimp height conductor', 'crimp height'], store: true },
      { key: 'terPullForceMin', aliases: ['усилие обрыва контакта от, n', 'усилие обрыва контакта от n', 'усилие обрыва от, n', 'усилие обрыва от', 'pull force min', 'pull test min', 'pull-off force min'], store: true },
      { key: 'terPullForceMax', aliases: ['усилие обрыва контакта до, n', 'усилие обрыва контакта до n', 'усилие обрыва до, n', 'усилие обрыва до', 'pull force max', 'pull test max', 'pull-off force max'], store: true },
    ],
    display: v => joinTechOperationsParts_([v.terManufacturer, v.terSeries, v.terComponent], ' | '),
    search:  v => [v.terManufacturer, v.terSeries, v.terComponent, v.connType, v.terType, v.artISL, v.artSAG].join(' '),
    // sortKey отсутствует — как в исходном ter-билдере (sortKey → '' в снапшоте).
  },
  coax: {
    fields: [
      { key: 'coaxArticle', aliases: ['артикул'], store: true },
      { key: 'coaxType',    aliases: ['тип/серия', 'тип / серия', 'тип серия'], store: true },
      { key: 'coaxMfr',     aliases: ['производитель', 'бренд', 'manufacturer'], store: true },
      { key: 'supplier',    aliases: ['поставщик', 'supplier'] },
      { key: 'coaxWire',    aliases: ['провод'], store: true },
      { key: 'program',     aliases: ['программа'] },
    ],
    display: v => joinTechOperationsParts_([v.coaxType, v.coaxWire, v.supplier], ' | '),
    sortKey: v => `${v.coaxWire} ${v.coaxType} ${v.coaxMfr} ${v.coaxArticle}`,
    search:  v => [v.coaxArticle, v.coaxType, v.coaxMfr, v.supplier, v.coaxWire, v.program].join(' '),
  },
};

// Резолвит индекс колонки для каждого поля схемы ОДИН раз на лист:
// алиас (точное совпадение нормализованного заголовка) → regex-фолбэк → -1.
function resolveSchemaColumns_(schema, headerMap) {
  const resolved = {};
  schema.fields.forEach((f) => {
    let idx = -1;
    for (let i = 0; i < f.aliases.length; i += 1) {
      const key = normalizeHeader_(f.aliases[i]);
      if (hasHeader_(headerMap, key)) { idx = headerMap[key]; break; }
    }
    if (idx < 0 && f.regex) {
      for (const k of Object.keys(headerMap)) {
        if (f.regex.test(k)) { idx = headerMap[k]; break; }
      }
    }
    resolved[f.key] = idx;
  });
  return resolved;
}

function readSchemaCell_(row, idx) {
  return idx >= 0 ? normalizeString_(row[idx]) : '';
}

// Строит запись по схеме вкладки. Возвращает null если displayText пуст
// (как ранние return null в исходных билдерах).
function buildRecordFromSchema_(tabKey, row, resolved, sourceSheet, namedColumns) {
  const schema = DB_SCHEMA[tabKey];
  if (!schema) return null;

  const v = {};
  schema.fields.forEach((f) => { v[f.key] = readSchemaCell_(row, resolved[f.key]); });

  const displayText = schema.display(v);
  if (!displayText) return null;

  const rec = {
    tabKey,
    displayText,
    normalizedSearch: normalizeSearch_(schema.search(v)),
    exportValues: (namedColumns || []).map(({ index }) => normalizeString_(row[index]) || ''),
    sourceSheet,
  };
  if (schema.sortKey) rec.sortKey = schema.sortKey(v);
  schema.fields.forEach((f) => { if (f.store) rec[f.key] = v[f.key] || ''; });
  return rec;
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
// ── Единый словарь extra-колонок снапшота ────────────────────
// Какие семантические поля записи лежат в extra-колонках (начиная с индекса
// TECHOPS_EXTRA_BASE) для каждой вкладки. Позиция в массиве = смещение колонки.
// Пустая строка '' = неиспользуемая колонка (заполняется пустым значением).
// ЕДИНЫЙ источник истины: и запись (toSnapshotRow_), и чтение (parseTechOpsRow_)
// выводятся отсюда → перестановка/добавление колонки правится в ОДНОМ месте,
// рассинхрон чтения и записи невозможен. См. таблицу TER в CLAUDE.md.
const TECHOPS_EXTRA_BASE = 6; // exportJson/sortKey занимают 0..5, extra-колонки — с 6
const TECHOPS_EXTRA_FIELDS = {
  op:   ['opNumber', 'opName', 'tOp', 'tPrep', 'tMachine'],
  ter:  ['terManufacturer', 'terSeries', 'terComponent', 'terType', 'terArticle',
         'terLPlus', 'terLMinus', 'terApplicator', 'terCrimpHeight',
         'terPullForceMin', 'terPullForceMax', 'terStep'],
  coax: ['', '', '', 'coaxWire', 'coaxType', 'coaxMfr', 'coaxArticle'],
  ob:   ['obType'],
};

function parseTechOpsRow_(row) {
  const tabKey = row[0];
  const rec = {
    tabKey,
    displayText:      row[1],
    normalizedSearch: row[2],
    exportValues:     parseJsonArray_(row[3]),
    sourceSheet:      row[4],
    sortKey:          row[5] || '',
  };
  (TECHOPS_EXTRA_FIELDS[tabKey] || []).forEach((field, i) => {
    if (field) rec[field] = row[TECHOPS_EXTRA_BASE + i] || '';
  });
  return rec;
}

// Возвращает массив extra-колонок снапшота для записи (индексы 6..17).
// Симметричен parseTechOpsRow_ через общий словарь TECHOPS_EXTRA_FIELDS.
function toSnapshotRow_(record) {
  const count = TECHOPS_DB_APP.dataHeaders.length - TECHOPS_EXTRA_BASE; // = 12
  const out = new Array(count).fill('');
  (TECHOPS_EXTRA_FIELDS[record.tabKey] || []).forEach((field, i) => {
    if (field && i < count) out[i] = record[field] || '';
  });
  return out;
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
  writeSheetHeader_(dataSheet, TECHOPS_DB_APP.dataHeaders);

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
  writeSheetHeader_(metaSheet, TECHOPS_DB_APP.metaHeaders);

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
      // getDisplayValues (не getValues): снапшот хранит строки из getDisplayValues
      // источника. getValues превратил бы числовые поля (шаг, L+, L-) обратно в
      // числа, и клиентский step.replace(...) падал бы с TypeError.
      dataSheet
        .getRange(2, 1, lastRow - 1, TECHOPS_DB_APP.dataHeaders.length)
        .getDisplayValues()
        .filter((row) => row[0] && row[1])
        .forEach((row) => records.push(parseTechOpsRow_(row)));
    }
  }

  const meta = {
    sourceSpreadsheetId: getSourceSpreadsheetId_(),
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
        if      (key === 'sourceSpreadsheetId') meta.sourceSpreadsheetId = value || getSourceSpreadsheetId_();
        else if (key === 'updatedAt')           meta.updatedAt = value || '';
        else if (key === 'recordCount')            meta.recordCount        = toInt_(value);
        else if (key === 'schemaVersion')          meta.schemaVersion      = toInt_(value);
        else if (key === 'countsByTabJson')        meta.countsByTab        = safeJsonParse_(value, {});
        else if (key === 'diagnosticsByTabJson')   meta.diagnosticsByTab   = safeJsonParse_(value, {});
        else if (key === 'columnHeadersByTabJson') meta.columnHeadersByTab = safeJsonParse_(value, {});
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
          item.terStep         = record.terStep         || '';
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
    sourceSpreadsheetId: snapshot.meta.sourceSpreadsheetId || getSourceSpreadsheetId_(),
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

/**
 * Гарантирует, что снапшот техопераций готов в кеше, БЕЗ форс-ресинка.
 * Тяжёлый syncTechOperationsDatabase() (чтение внешней таблицы) вызывается
 * только если кеша/листов нет или сменилась схема. Быстрый путь открытия
 * генератора и сайдбара — данные обновляются вручную (кнопка/меню).
 */
function ensureTechOperationsSnapshotReady_() {
  const cached = getTechOpsCache_().load();
  if (cached && cached.records && cached.records.length &&
      String(cached.meta && cached.meta.schemaVersion) === String(TECHOPS_DB_APP.schemaVersion)) {
    return;
  }
  ensureTechOperationsInfrastructure_(SpreadsheetApp.getActive());
  const stored = loadTechOperationsSnapshotFromSheets_();
  if (stored.records.length &&
      String(stored.meta.schemaVersion) === String(TECHOPS_DB_APP.schemaVersion)) {
    getTechOpsCache_().save(stored);
    return;
  }
  syncTechOperationsDatabase(); // пусто или сменилась схема — единственный случай синка
}

/** Возвращает snapshot из кеша или листов; используется AssemblyGenerator. */
function getTechOperationsSnapshot_() {
  const cached = getTechOpsCache_().load();
  if (cached && cached.records && cached.records.length) return cached;
  const stored = loadTechOperationsSnapshotFromSheets_();
  if (stored.records.length) getTechOpsCache_().save(stored);
  return stored;
}
