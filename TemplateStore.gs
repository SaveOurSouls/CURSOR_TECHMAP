// Internal implementation: catalog/store management, copy/paste helpers, image handling.
// All functions here are private (suffixed _) and called from Code.gs entry points.

/**
 * Диапазон для сохранения: из данных диалога (на момент открытия), иначе текущее выделение.
 * Повторный getActiveWorkingRange_ при открытом модальном окне нестабилен и может давать сбои API.
 */
function resolveTemplateSourceRange_(formData) {
  const sel = formData && formData.selection;
  if (sel && sel.sheetName && sel.rangeA1) {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sel.sheetName);
    if (!sheet || isSystemSheet_(sheet.getName())) {
      throw new Error(
        'Неверный лист выделения. Закройте окно, выделите шаблон на рабочем листе и откройте «Сохранить шаблон» снова.'
      );
    }
    const range = sheet.getRange(sel.rangeA1);
    if (
      sel.height &&
      sel.width &&
      (range.getNumRows() !== sel.height || range.getNumColumns() !== sel.width)
    ) {
      throw new Error(
        'Выделение изменилось с момента открытия окна. Закройте диалог и откройте «Сохранить шаблон» заново.'
      );
    }
    validateTemplateRange_(range);
    return range;
  }
  return getActiveWorkingRange_();
}

function ensureInfrastructure_(ssArg) {
  const ss = ssArg || SpreadsheetApp.getActive();
  ensureCatalogSheet_(ss);
  ensureStoreSheet_(ss);
  removeCanvasSheet_();
}

function ensureCatalogSheet_(ss) {
  let sheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHMAP_APP.librarySheetName);
    sheet
      .getRange(1, 1, 1, TECHMAP_APP.catalogHeaders.length)
      .setValues([TECHMAP_APP.catalogHeaders])
      .setFontWeight('bold')
      .setBackground('#d9e2f3');
    sheet.hideSheet();
  }
  return sheet;
}

// Lightweight read-only check: returns the catalog sheet if it exists, null otherwise.
function getCatalogSheetIfExists_(ss) {
  return (ss || SpreadsheetApp.getActive()).getSheetByName(TECHMAP_APP.librarySheetName);
}

function ensureStoreSheet_(ss) {
  let sheet = ss.getSheetByName(TECHMAP_APP.storeSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHMAP_APP.storeSheetName);
    sheet.hideSheet();
  }
  return sheet;
}

function readCatalog_() {
  const sheet = getCatalogSheetIfExists_(SpreadsheetApp.getActive());
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, TECHMAP_APP.catalogHeaders.length)
    .getValues()
    .filter((row) => row[0])
    .map((row) => ({
      id: String(row[0]),
      title: row[1],
      category: row[2],
      description: row[3],
      storeRow: toInt_(row[4]),
      storeColumn: toInt_(row[5]),
      height: toInt_(row[6]),
      width: toInt_(row[7]),
      sourceSheet: row[8],
      sourceRange: row[9],
      updatedAt: row[10],
      rowHeights: parseJsonArray_(row[11]),
      columnWidths: parseJsonArray_(row[12]),
      imagesJson: row[13] || '[]',
    }));
}

function upsertCatalogRecord_(catalogSheet, record) {
  const lastRow = catalogSheet.getLastRow();
  const ids = lastRow > 1
    ? catalogSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
    : [];
  const existingIndex = ids.indexOf(record.id);
  const rowValues = [[
    record.id,
    record.title,
    record.category,
    record.description,
    record.storeRow,
    record.storeColumn,
    record.height,
    record.width,
    record.sourceSheet,
    record.sourceRange,
    record.updatedAt,
    record.rowHeightsJson,
    record.columnWidthsJson,
    record.imagesJson || '[]',
  ]];

  const targetRow = existingIndex >= 0 ? existingIndex + 2 : lastRow + 1;
  catalogSheet.getRange(targetRow, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues(rowValues);
}

function getTemplateById_(templateId) {
  const template = readCatalog_().find((item) => item.id === templateId);
  if (!template) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }
  return template;
}

function getActiveWorkingRange_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  if (!sheet || isSystemSheet_(sheet.getName())) {
    throw new Error('Выберите диапазон на рабочем листе, а не на служебном.');
  }

  const range = sheet.getActiveRange();
  if (!range) {
    throw new Error('Не выбран диапазон.');
  }

  validateTemplateRange_(range);
  return range;
}

function validateTemplateRange_(range) {
  const startRow = range.getRow();
  const endRow = startRow + range.getNumRows() - 1;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getNumColumns() - 1;

  const invalidMerge = range.getMergedRanges().find((mergedRange) => {
    const mergedStartRow = mergedRange.getRow();
    const mergedEndRow = mergedStartRow + mergedRange.getNumRows() - 1;
    const mergedStartColumn = mergedRange.getColumn();
    const mergedEndColumn = mergedStartColumn + mergedRange.getNumColumns() - 1;

    return (
      mergedStartRow < startRow ||
      mergedEndRow > endRow ||
      mergedStartColumn < startColumn ||
      mergedEndColumn > endColumn
    );
  });

  if (invalidMerge) {
    throw new Error(
      `Диапазон содержит неполностью захваченное объединение ячеек (${invalidMerge.getA1Notation()}). Выделите шаблон целиком.`
    );
  }
}

function allocateStoreLocation_(range, existingTemplate, catalog) {
  if (
    existingTemplate &&
    existingTemplate.height === range.getNumRows() &&
    existingTemplate.width === range.getNumColumns()
  ) {
    return { row: existingTemplate.storeRow, column: existingTemplate.storeColumn };
  }

  // Append after the last live block. No compaction — cleared slots from deletions are
  // simply skipped. Store grows slowly over time but save stays fast.
  const nextRow = (catalog || []).reduce((maxRow, item) => {
    if (existingTemplate && item.id === existingTemplate.id) return maxRow;
    return Math.max(maxRow, item.storeRow + item.height - 1);
  }, 0) + 1;

  return { row: nextRow, column: 1 };
}

/**
 * Временно показывает скрытый лист на время callback.
 * Служба «Таблицы» часто возвращает «Ошибка службы: Таблицы» для операций
 * clear, copyTo и deleteRows по полностью скрытому листу (в т.ч. _TC_STORE).
 */
function runWithSheetVisible_(sheet, fn) {
  if (!sheet) {
    return fn();
  }
  const wasHidden = sheet.isSheetHidden();
  if (wasHidden) {
    sheet.showSheet();
  }
  try {
    return fn();
  } finally {
    if (wasHidden) {
      sheet.hideSheet();
    }
  }
}

/**
 * Physically rewrites _TC_STORE so it contains only live template blocks
 * packed together from row 1, and updates storeRow in the catalog.
 */
function compactifyStore_(catalog) {
  const ss = SpreadsheetApp.getActive();
  const savedSheet = ss.getActiveSheet();

  const storeSheet = ensureStoreSheet_(ss);
  const catalogSheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!catalogSheet) {
    return;
  }

  runWithSheetVisible_(storeSheet, () => {
    const live = (catalog || [])
      .filter((item) => item.storeRow > 0 && item.height > 0 && item.width > 0)
      .sort((a, b) => a.storeRow - b.storeRow);

    if (!live.length) {
      const maxRow = storeSheet.getLastRow();
      if (maxRow > 0) {
        storeSheet.deleteRows(1, maxRow);
      }
      return;
    }

    let alreadyCompact = true;
    let cursor = 1;
    for (const item of live) {
      if (item.storeRow !== cursor) {
        alreadyCompact = false;
        break;
      }
      cursor += item.height;
    }
    if (alreadyCompact) {
      return;
    }

    const tempName = '_TC_COMPACT_TMP';
    let tempSheet = ss.getSheetByName(tempName);
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    tempSheet = ss.insertSheet(tempName);
    tempSheet.hideSheet();

    let writeRow = 1;
    const newStoreRows = {};

    live.forEach((item) => {
      const srcRange = storeSheet.getRange(item.storeRow, item.storeColumn, item.height, item.width);
      ensureSheetCapacity_(tempSheet, writeRow + item.height - 1, item.width);
      const destRange = tempSheet.getRange(writeRow, 1, item.height, item.width);
      destRange.breakApart();
      copyRangePreservingFormulas_(srcRange, destRange);
      newStoreRows[item.id] = writeRow;
      writeRow += item.height;
    });

    const storeLastRow = storeSheet.getLastRow();
    if (storeLastRow > 0) {
      storeSheet.deleteRows(1, storeLastRow);
    }
    if (writeRow > 1) {
      const cols = live.reduce((max, item) => Math.max(max, item.width || 0), 0) || 20;
      ensureSheetCapacity_(storeSheet, writeRow - 1, cols);
      const compactedRange = tempSheet.getRange(1, 1, writeRow - 1, cols);
      copyRangePreservingFormulas_(compactedRange, storeSheet.getRange(1, 1, writeRow - 1, cols));
    }

    ss.deleteSheet(tempSheet);

    const catalogLastRow = catalogSheet.getLastRow();
    if (catalogLastRow < 2) {
      if (!isSystemSheet_(savedSheet.getName())) {
        ss.setActiveSheet(savedSheet);
      }
      return;
    }
    const catalogIds = catalogSheet.getRange(2, 1, catalogLastRow - 1, 1).getValues();
    catalogIds.forEach((row, idx) => {
      const id = String(row[0]).trim();
      if (newStoreRows[id] !== undefined) {
        catalogSheet.getRange(idx + 2, 5).setValue(newStoreRows[id]);
      }
    });

    if (!isSystemSheet_(savedSheet.getName())) {
      ss.setActiveSheet(savedSheet);
    }
  });
}

function writeRangeToStore_(sourceRange, storeRow, storeColumn) {
  const storeSheet = ensureStoreSheet_(SpreadsheetApp.getActive());
  const height = sourceRange.getNumRows();
  const width = sourceRange.getNumColumns();
  ensureSheetCapacity_(storeSheet, storeRow + height - 1, storeColumn + width - 1);

  runWithSheetVisible_(storeSheet, () => {
    const targetRange = storeSheet.getRange(storeRow, storeColumn, height, width);
    targetRange.breakApart();
    clearStoreSlotForWrite_(targetRange);
    SpreadsheetApp.flush();
    copyRangePreservingFormulas_(sourceRange, targetRange);
    targetRange.getCell(1, 1).setNote('techmap-template-store');
  });
}

/** Очистка слота _TC_STORE перед записью; полный clear() по скрытому листу иногда падает в API. */
function clearStoreSlotForWrite_(targetRange) {
  try {
    targetRange.clear({ contentsOnly: false });
  } catch (e) {
    try {
      targetRange.clearDataValidations();
    } catch (e2) {}
    try {
      targetRange.clearFormat();
    } catch (e3) {}
    try {
      targetRange.clearContent();
    } catch (e4) {}
    try {
      targetRange.removeCheckboxes();
    } catch (e5) {}
  }
}

/**
 * Один setValues на большой сетке иногда даёт «Ошибка службы: Таблицы» — режем на полосы.
 */
function setRangeValuesChunked_(targetRange, values) {
  if (!values || !values.length) {
    return;
  }
  const numRows = values.length;
  const numCols = values[0].length;
  const maxChunk = 25;
  if (numRows <= maxChunk) {
    try {
      targetRange.setValues(values);
    } catch (e) {
      Utilities.sleep(120);
      try { SpreadsheetApp.flush(); } catch (fe) {}
      setRangeValuesRowByRow_(targetRange, values);
    }
    return;
  }
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();
  const sheet = targetRange.getSheet();
  for (let r = 0; r < numRows; r += maxChunk) {
    const h = Math.min(maxChunk, numRows - r);
    const slice = values.slice(r, r + h);
    try {
      sheet.getRange(startRow + r, startCol, h, numCols).setValues(slice);
      SpreadsheetApp.flush();
    } catch (e) {
      setRangeValuesRowByRow_(sheet.getRange(startRow + r, startCol, h, numCols), slice);
    }
  }
}

/**
 * Резервная запись значений построчно/поячеечно, когда setValues на блоке падает.
 */
function setRangeValuesRowByRow_(targetRange, values) {
  const sheet = targetRange.getSheet();
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();
  const numCols = values[0] ? values[0].length : 0;
  for (let r = 0; r < values.length; r += 1) {
    const rowValues = [values[r]];
    try {
      sheet.getRange(startRow + r, startCol, 1, numCols).setValues(rowValues);
    } catch (e) {
      for (let c = 0; c < numCols; c += 1) {
        try {
          sheet.getRange(startRow + r, startCol + c).setValue(values[r][c]);
        } catch (e2) {}
      }
    }
    if (r % 20 === 0) {
      try { SpreadsheetApp.flush(); } catch (fe) {}
      Utilities.sleep(20);
    }
  }
}

function applySourceFormulasCellwise_(targetRange, formulas) {
  if (!formulas || !formulas.length) return;
  const sheet = targetRange.getSheet();
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();
  // Apply only cells that actually have formulas — setFormulas on a block writes empty
  // strings for non-formula cells in the same row, which erases plain text values.
  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const f = row[c];
      if (!f) continue;
      try { sheet.getRange(startRow + r, startCol + c).setFormula(f); } catch (e) {}
    }
  }
}

/**
 * Copies a range using Sheets API v4 batchUpdate/copyPaste.
 * Requires "Google Sheets API" advanced service enabled in Extensions > Apps Script > Services.
 * Returns true on success, false if the service is unavailable or the call fails.
 */
function tryCopyRangeViaSheetsApi_(sourceRange, targetRange) {
  try {
    if (typeof Sheets === 'undefined') return false;
    var ssId = SpreadsheetApp.getActive().getId();
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        copyPaste: {
          source: {
            sheetId: sourceRange.getSheet().getSheetId(),
            startRowIndex: sourceRange.getRow() - 1,
            endRowIndex: sourceRange.getRow() + sourceRange.getNumRows() - 1,
            startColumnIndex: sourceRange.getColumn() - 1,
            endColumnIndex: sourceRange.getColumn() + sourceRange.getNumColumns() - 1,
          },
          destination: {
            sheetId: targetRange.getSheet().getSheetId(),
            startRowIndex: targetRange.getRow() - 1,
            endRowIndex: targetRange.getRow() + targetRange.getNumRows() - 1,
            startColumnIndex: targetRange.getColumn() - 1,
            endColumnIndex: targetRange.getColumn() + targetRange.getNumColumns() - 1,
          },
          pasteType: 'PASTE_NORMAL',
          pasteOrientation: 'NORMAL',
        }
      }]
    }, ssId);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Gets the blob of an over-grid image. Tries img.getBlob() first (fast path), then falls
 * back to downloading via the Sheets API sourceUri if getBlob() throws (a known GAS bug
 * for images inserted via Insert > Image in newer Sheets versions).
 * Returns null if both approaches fail.
 */
/**
 * Gets the blob of an over-grid image.
 * New GAS API (2024+) uses getUrl() instead of getBlob() — tries both.
 * Downloads via UrlFetchApp with OAuth when only a URL is available.
 */
function getOverGridImageBlob_(img, sourceSheet) {
  // Legacy path: getBlob() (older GAS / image types)
  try {
    if (typeof img.getBlob === 'function') {
      var blob = img.getBlob();
      if (blob) return blob;
    }
  } catch (e) {}

  // New path: getUrl() + UrlFetchApp (newer GAS image type)
  try {
    if (typeof img.getUrl === 'function') {
      var url = img.getUrl();
      if (url) {
        var resp = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true,
        });
        if (resp.getResponseCode() === 200) return resp.getBlob();
      }
    }
  } catch (e) {}

  return null;
}

/**
 * Explicitly copies in-cell image values via Sheets API.
 * PASTE_NORMAL (both copyTo and copyPaste) does not reliably transfer in-cell images.
 * Reads userEnteredValue.imageValue from each source cell and writes to destination
 * with a targeted updateCells request so no other cell content is disturbed.
 * Requires "Google Sheets API" advanced service. No-op if service is not enabled.
 */
function copyInCellImageValues_(sourceRange, targetRange) {
  if (typeof Sheets === 'undefined') return false;
  try {
    var ssId = SpreadsheetApp.getActive().getId();
    var srcName = sourceRange.getSheet().getName();
    var a1 = "'" + srcName.replace(/'/g, "''") + "'!" + sourceRange.getA1Notation();
    var resp = Sheets.Spreadsheets.get(ssId, {
      ranges: [a1],
      fields: 'sheets(data(rowData(values(userEnteredValue))))',
    });
    var rows = ((resp.sheets || [])[0] || {}).data;
    rows = rows && rows[0] && rows[0].rowData ? rows[0].rowData : [];
    var requests = [];
    rows.forEach(function(row, ri) {
      (row.values || []).forEach(function(cell, ci) {
        var iv = cell.userEnteredValue && cell.userEnteredValue.imageValue;
        if (!iv) return;
        requests.push({
          updateCells: {
            rows: [{ values: [{ userEnteredValue: { imageValue: iv } }] }],
            fields: 'userEnteredValue.imageValue',
            start: {
              sheetId: targetRange.getSheet().getSheetId(),
              rowIndex: targetRange.getRow() - 1 + ri,
              columnIndex: targetRange.getColumn() - 1 + ci,
            },
          },
        });
      });
    });
    if (requests.length) {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
    }
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Copies a range to a target without adjusting formula references.
 * Prefers Sheets API v4 (preserves in-cell images); falls back to copyTo(PASTE_NORMAL),
 * then to manual value+format copy if both fail.
 * Relative formula references are overwritten with original strings after the paste.
 */
function copyRangePreservingFormulas_(sourceRange, targetRange) {
  const srcSheet = sourceRange.getSheet();
  const dstSheet = targetRange.getSheet();
  const hideSrc = srcSheet.isSheetHidden();
  const hideDst = dstSheet.isSheetHidden();
  if (hideSrc) srcSheet.showSheet();
  if (hideDst) dstSheet.showSheet();

  const ss = SpreadsheetApp.getActive();
  const priorActive = ss.getActiveSheet();

  try {
    targetRange.breakApart();

    let formulas;
    try {
      formulas = sourceRange.getFormulas();
    } catch (e) {
      formulas = null;
    }

    // Sheets API preserves in-cell images; fall back to copyTo if service not enabled.
    const copiedViaSheetsApi = tryCopyRangeViaSheetsApi_(sourceRange, targetRange);
    if (!copiedViaSheetsApi) {
      try {
        ss.setActiveSheet(srcSheet);
        SpreadsheetApp.flush();
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        SpreadsheetApp.flush();
      } catch (e) {
        try {
          const values = sourceRange.getValues();
          setRangeValuesChunked_(targetRange, values);
        } catch (ve) {}
        SpreadsheetApp.flush();
        copyRangeFormatPreservingMerges_(sourceRange, targetRange, ss, priorActive, srcSheet, dstSheet);
      }
    }

    if (formulas) {
      applySourceFormulasCellwise_(targetRange, formulas);
    }

    // Copy in-cell image values explicitly — PASTE_NORMAL misses them.
    copyInCellImageValues_(sourceRange, targetRange);
  } finally {
    if (hideSrc) srcSheet.hideSheet();
    if (hideDst) dstSheet.hideSheet();
    if (priorActive && !isSystemSheet_(priorActive.getName())) {
      try { ss.setActiveSheet(priorActive); } catch (e) {}
    }
  }
}

/**
 * PASTE_FORMAT после записи значений; copyTo между листами иногда нестабилен, если активен «чужой» лист.
 * При неудаче остаются значения и формулы (без объединений/рамок).
 */
function copyRangeFormatPreservingMerges_(sourceRange, targetRange, ss, priorActive, srcSheet, dstSheet) {
  const restorePrior = () => {
    if (
      priorActive &&
      !isSystemSheet_(priorActive.getName()) &&
      ss.getActiveSheet().getSheetId() !== priorActive.getSheetId()
    ) {
      try {
        ss.setActiveSheet(priorActive);
      } catch (e) {}
    }
  };

  const attempt = (sheetToActivate) => {
    if (sheetToActivate && ss.getActiveSheet().getSheetId() !== sheetToActivate.getSheetId()) {
      ss.setActiveSheet(sheetToActivate);
    }
    SpreadsheetApp.flush();
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  };

  try {
    attempt(srcSheet);
  } catch (e1) {
    try {
      attempt(dstSheet);
    } catch (e2) {
      try {
        Utilities.sleep(200);
        SpreadsheetApp.flush();
        attempt(dstSheet);
      } catch (e3) {}
    }
  } finally {
    restorePrior();
  }
}

function applyStoredDimensions_(targetSheet, targetRow, targetColumn, template) {
  // Group consecutive rows/columns with the same dimension into one setRowHeights/
  // setColumnWidths call (plural) instead of one call per row/column.
  const rowHeights = template.rowHeights || [];
  let i = 0;
  while (i < rowHeights.length) {
    const h = rowHeights[i];
    if (!h) { i++; continue; }
    let j = i + 1;
    while (j < rowHeights.length && rowHeights[j] === h) j++;
    targetSheet.setRowHeights(targetRow + i, j - i, h);
    i = j;
  }

  const columnWidths = template.columnWidths || [];
  i = 0;
  while (i < columnWidths.length) {
    const w = columnWidths[i];
    if (!w) { i++; continue; }
    let j = i + 1;
    while (j < columnWidths.length && columnWidths[j] === w) j++;
    targetSheet.setColumnWidths(targetColumn + i, j - i, w);
    i = j;
  }
}

function getRowHeights_(range) {
  const sheet = range.getSheet();
  const heights = [];
  for (let index = 0; index < range.getNumRows(); index += 1) {
    heights.push(sheet.getRowHeight(range.getRow() + index));
  }
  return heights;
}

function getColumnWidths_(range) {
  const sheet = range.getSheet();
  const widths = [];
  for (let index = 0; index < range.getNumColumns(); index += 1) {
    widths.push(sheet.getColumnWidth(range.getColumn() + index));
  }
  return widths;
}

function clearTemplateMarkerNote_(targetRange) {
  const note = targetRange.getCell(1, 1).getNote();
  if (note && note.indexOf('techmap-template') === 0) {
    targetRange.getCell(1, 1).clearNote();
  }
}

function hideLegacyTemplateSheets_(ss) {
  ss.getSheets().forEach((sheet) => {
    if (sheet.getName().indexOf(TECHMAP_APP.legacyTemplatePrefix) === 0) {
      sheet.hideSheet();
    }
  });
}

// Creates a new sheet with a unique name based on baseName.
// If a sheet named baseName already exists, appends -2, -3, ... until free.
// Google Sheets sheet names are limited to 100 characters.
function createUniqueSheet_(ss, baseName) {
  const MAX_LEN = 100;
  const base = String(baseName || 'Лист').substring(0, MAX_LEN);
  if (!ss.getSheetByName(base)) {
    return ss.insertSheet(base);
  }
  for (let i = 2; i <= 999; i++) {
    const suffix = '-' + i;
    const candidate = base.substring(0, MAX_LEN - suffix.length) + suffix;
    if (!ss.getSheetByName(candidate)) {
      return ss.insertSheet(candidate);
    }
  }
  return ss.insertSheet();
}

function ensureSheetCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function parseJsonArray_(value) {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function makeTemplateId_(title, catalog) {
  const base = slugify_(title) || 'template';
  let candidate = base;
  let index = 2;
  const existingIds = new Set((catalog || []).map((item) => item.id));
  while (existingIds.has(candidate)) {
    candidate = `${base}-${index}`;
    index += 1;
  }
  return candidate;
}

function slugify_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9а-яё]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-+/g, '-');
}

function normalizeString_(value) {
  return String(value || '').trim();
}

function toInt_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function isSystemSheet_(sheetName) {
  return (
    sheetName === TECHMAP_APP.librarySheetName ||
    sheetName === TECHMAP_APP.storeSheetName ||
    sheetName.indexOf(TECHMAP_APP.legacyTemplatePrefix) === 0
  );
}

function removeCanvasSheet_() {
  const ss = SpreadsheetApp.getActive();
  const canvas = ss.getSheetByName(TECHMAP_APP.canvasSheetName);
  if (!canvas) {
    return;
  }
  const nonSystem = ss.getSheets().filter((s) => !isSystemSheet_(s.getName()) && s.getName() !== TECHMAP_APP.canvasSheetName);
  if (nonSystem.length > 0) {
    ss.setActiveSheet(nonSystem[0]);
  }
  try {
    ss.deleteSheet(canvas);
  } catch (e) {
    canvas.hideSheet();
  }
}

// ===== Over-grid image handling =====
// Capture: sheet.getImages() at save time → base64 in imagesJson.
// Insert:  base64 → insertImage() if imagesJson present,
//          else   sheet.getImages() on _TC_STORE for manually-placed images.

function captureTemplateImages_(sourceSheet, range) {
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var endRow   = startRow + range.getNumRows() - 1;
  var endCol   = startCol + range.getNumColumns() - 1;
  var result = [];
  try {
    sourceSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow();
        var col = anchor.getColumn();
        if (row < startRow || row > endRow || col < startCol || col > endCol) return;
        var blob = img.getBlob();
        result.push({
          relRow:   row - startRow,
          relCol:   col - startCol,
          xOffset:  img.getAnchorCellXOffset(),
          yOffset:  img.getAnchorCellYOffset(),
          width:    img.getWidth(),
          height:   img.getHeight(),
          mimeType: blob.getContentType() || 'image/png',
          base64:   Utilities.base64Encode(blob.getBytes()),
        });
      } catch (e) {}
    });
  } catch (e) {}
  return result;
}

// Copies over-grid images from _TC_STORE to the target sheet at the given offset.
// Tagged images (tc: alt-text) use the extended margin zone matching copySourceImagesToStore_.
// Untagged images (manually placed in STORE) use strict slot bounds.
function insertOverGridImages_(sourceSheet, sourceRow, sourceCol, height, width, targetSheet, targetRow, targetCol) {
  var xlsxImagesInsert = null; // Lazy XLSX fallback for GAS 2024+ image type
  try {
    sourceSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow();
        var col = anchor.getColumn();
        var relRow = row - sourceRow;
        var relCol = col - sourceCol;
        var hasTcTag = false;
        try {
          var alt = img.getAltTextDescription();
          if (alt && alt.indexOf('tc:') === 0) {
            hasTcTag = true;
            var parts = alt.slice(3).split(',');
            if (parts.length === 2) {
              var pr = parseInt(parts[0], 10);
              var pc = parseInt(parts[1], 10);
              if (!isNaN(pr)) relRow = pr;
              if (!isNaN(pc)) relCol = pc;
            }
          }
        } catch (e) {}
        // Tagged: extend bounds to match the ±2 margin used in copySourceImagesToStore_.
        // Untagged (manual): strict slot bounds only.
        if (hasTcTag) {
          if (row < sourceRow || row > sourceRow + height + 1) return;
          if (col < sourceCol || col > sourceCol + width + 1) return;
        } else {
          if (row < sourceRow || row > sourceRow + height - 1) return;
          if (col < sourceCol || col > sourceCol + width - 1) return;
        }
        var destRow = Math.max(1, targetRow + relRow);
        var destCol = Math.max(1, targetCol + relCol);
        var blob = getOverGridImageBlob_(img, sourceSheet);
        if (!blob) {
          if (!xlsxImagesInsert) xlsxImagesInsert = buildXlsxImageMap_(sourceSheet);
          blob = (xlsxImagesInsert && xlsxImagesInsert[row + '_' + col]) || null;
        }
        if (!blob) return;
        var inserted = targetSheet.insertImage(blob, destCol, destRow,
          img.getAnchorCellXOffset(), img.getAnchorCellYOffset());
        if (inserted) inserted.setWidth(img.getWidth()).setHeight(img.getHeight());
      } catch (e) {}
    });
  } catch (e) {}
}

// Inserts over-grid images from imagesJson. Prefers driveFileId (fast Drive read);
// falls back to base64 for old templates saved before Drive caching was added.
function insertTemplateImages_(targetSheet, targetRow, targetCol, imagesJson) {
  var images = parseJsonArray_(imagesJson);
  if (!images.length) return;
  images.forEach(function(img) {
    var row = targetRow + (img.relRow || 0);
    var col = targetCol + (img.relCol || 0);
    try {
      var blob;
      if (img.driveFileId) {
        blob = DriveApp.getFileById(img.driveFileId).getBlob();
      } else if (img.base64) {
        blob = Utilities.newBlob(
          Utilities.base64Decode(img.base64),
          img.mimeType || 'image/png',
          'tc_image'
        );
      }
      if (!blob) return;
      var inserted = targetSheet.insertImage(blob, col, row, img.xOffset || 0, img.yOffset || 0);
      if (inserted && img.width && img.height) {
        inserted.setWidth(img.width).setHeight(img.height);
      }
    } catch (e) {}
  });
}

function deleteTemplateImages_(imagesJson) {
  var images = parseJsonArray_(imagesJson);
  images.forEach(function(img) {
    if (img.driveFileId) {
      try { DriveApp.getFileById(img.driveFileId).setTrashed(true); } catch (e) {}
    }
  });
}

function getOrCreateImageCacheFolder_() {
  var props = PropertiesService.getUserProperties();
  var folderId = props.getProperty('tc_image_cache_folder');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  var folder = DriveApp.createFolder('_TC_IMAGE_CACHE');
  props.setProperty('tc_image_cache_folder', folder.getId());
  return folder;
}

/**
 * Downloads the spreadsheet as XLSX (ZIP), unzips it, and returns a map of
 * { 'row_col': Blob } for every image found in the drawing XML of the given sheet.
 * Used as a last-resort fallback when getBlob() and getUrl() both fail (GAS 2024+ image type).
 */
function buildXlsxImageMap_(sheet) {
  try {
    var ss = SpreadsheetApp.getActive();
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return {};

    var files = {};
    Utilities.unzip(resp.getBlob().setContentType('application/zip')).forEach(function(f) { files[f.getName()] = f; });

    // Find 1-based sheet index
    var sheets = ss.getSheets();
    var sheetIdx = 0;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === sheet.getSheetId()) { sheetIdx = i + 1; break; }
    }
    if (!sheetIdx) return {};

    var sheetXml = files['xl/worksheets/sheet' + sheetIdx + '.xml'];
    if (!sheetXml) return {};
    var sheetStr = sheetXml.getDataAsString();

    var drawRIdM = sheetStr.match(/<drawing[^>]+r:id="([^"]+)"/);
    if (!drawRIdM) return {};

    var sheetRels = files['xl/worksheets/_rels/sheet' + sheetIdx + '.xml.rels'];
    if (!sheetRels) return {};
    var sheetRelsMap = parseXmlRels_(sheetRels.getDataAsString());
    var drawTarget = sheetRelsMap[drawRIdM[1]];
    if (!drawTarget) return {};
    var drawFile = normalizePath_('xl/worksheets/' + drawTarget);

    var drawXml = files[drawFile];
    if (!drawXml) return {};
    var drawStr = drawXml.getDataAsString();

    var drawRelsFile = normalizePath_(drawFile.replace('/drawings/', '/drawings/_rels/') + '.rels');
    var drawRels = files[drawRelsFile];
    var drawRelsMap = drawRels ? parseXmlRels_(drawRels.getDataAsString()) : {};

    var result = {};
    var anchorRe = /<xdr:(?:one|two)CellAnchor[^>]*>([\s\S]*?)<\/xdr:(?:one|two)CellAnchor>/g;
    var m;
    while ((m = anchorRe.exec(drawStr)) !== null) {
      var content = m[1];
      // In XLSX xdr:from: col is listed before row
      var posM = content.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
      if (!posM) continue;
      var col1 = parseInt(posM[1], 10) + 1; // to 1-based
      var row1 = parseInt(posM[2], 10) + 1;
      var embedM = content.match(/r:embed="([^"]+)"/);
      if (!embedM) continue;
      var imgTarget = drawRelsMap[embedM[1]];
      if (!imgTarget) continue;
      var imgFile = normalizePath_('xl/drawings/' + imgTarget);
      var imgBlob = files[imgFile];
      if (imgBlob) result[row1 + '_' + col1] = imgBlob;
    }
    return result;
  } catch (e) {
    return {};
  }
}

function parseXmlRels_(relsStr) {
  var map = {};
  var re = /<Relationship\b([^>]*)>/g;
  var m;
  while ((m = re.exec(relsStr)) !== null) {
    var attrs = m[1];
    var idM = attrs.match(/Id="([^"]+)"/);
    var tM = attrs.match(/Target="([^"]+)"/);
    if (idM && tM) map[idM[1]] = tM[1];
  }
  return map;
}

function normalizePath_(path) {
  var parts = path.split('/');
  var out = [];
  parts.forEach(function(p) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  });
  return out.join('/');
}

// Copies over-grid images from sourceRange to the template's slot in _TC_STORE.
// Captures images anchored up to 2 cells outside the range so logos in margin
// columns (e.g. anchor A1 when template starts at B1) are not missed.
// Stores the true relative offset in alt-text so insertOverGridImages_ can
// restore the correct position even when the anchor was clamped.
// Also saves each blob to Drive (_TC_IMAGE_CACHE folder) for fast retrieval.
// Returns an array of image metadata objects (driveFileId, relRow, relCol, …).
function copySourceImagesToStore_(sourceSheet, range, storeSheet, storeRow, storeCol) {
  var height = range.getNumRows();
  var width = range.getNumColumns();
  var startRow = range.getRow();
  var startCol = range.getColumn();

  clearStoreSlotImages_(storeSheet, storeRow, storeCol, height, width);

  var xlsxImages = null; // Lazy-loaded on first getBlob/getUrl failure
  var imagesData = [];
  var folder = null; // Lazy: only created if images are found

  try {
    sourceSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow();
        var col = anchor.getColumn();
        // ±2-cell margin catches images anchored just outside the selection
        if (row < startRow - 2 || row > startRow + height + 1) return;
        if (col < startCol - 2 || col > startCol + width + 1) return;
        var blob = getOverGridImageBlob_(img, sourceSheet);
        if (!blob) {
          if (!xlsxImages) xlsxImages = buildXlsxImageMap_(sourceSheet);
          blob = (xlsxImages && xlsxImages[row + '_' + col]) || null;
        }
        if (!blob) return;
        var relRow = row - startRow;
        var relCol = col - startCol;
        // Clamp anchor so the STORE column/row index stays ≥ 1
        var destRow = Math.max(storeRow, storeRow + relRow);
        var destCol = Math.max(storeCol, storeCol + relCol);
        var inserted = storeSheet.insertImage(
          blob, destCol, destRow,
          img.getAnchorCellXOffset(), img.getAnchorCellYOffset()
        );
        if (inserted) {
          inserted.setWidth(img.getWidth()).setHeight(img.getHeight());
          // Persist true relative offset for correct retrieval during insert
          try { inserted.setAltTextDescription('tc:' + relRow + ',' + relCol); } catch (e) {}
        }
        // Save to Drive for fast retrieval during insert (avoids XLSX re-download)
        try {
          if (!folder) folder = getOrCreateImageCacheFolder_();
          var driveFile = folder.createFile(blob.setName('tc_img_' + relRow + '_' + relCol));
          imagesData.push({
            driveFileId: driveFile.getId(),
            relRow: relRow,
            relCol: relCol,
            xOffset: img.getAnchorCellXOffset(),
            yOffset: img.getAnchorCellYOffset(),
            width: img.getWidth(),
            height: img.getHeight(),
          });
        } catch (e) {}
      } catch (e) {}
    });
  } catch (e) {}
  return imagesData;
}

// Removes all over-grid images belonging to the given slot of _TC_STORE.
// Within strict bounds: always removed. In the extended margin zone (±2 cols/rows
// matching copySourceImagesToStore_): only tc:-tagged images are removed to avoid
// disturbing images from adjacent template slots.
function clearStoreSlotImages_(storeSheet, storeRow, storeCol, height, width) {
  try {
    storeSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var r = anchor.getRow();
        var c = anchor.getColumn();
        // Strict slot bounds — always remove.
        if (r >= storeRow && r < storeRow + height &&
            c >= storeCol && c < storeCol + width) {
          img.remove();
          return;
        }
        // Extended margin zone — only remove if tagged by copySourceImagesToStore_.
        if (r >= storeRow && r < storeRow + height &&
            c >= storeCol && c <= storeCol + width + 1) {
          var tag = '';
          try { tag = img.getAltTextDescription() || ''; } catch (e) {}
          if (tag.indexOf('tc:') === 0) img.remove();
        }
      } catch (e) {}
    });
  } catch (e) {}
}

// Wipes the entire _TC_STORE sheet: all cells and all over-grid images.
// Called when the catalog becomes empty so no orphaned blocks remain.
function purgeEntireStore_(storeSheet) {
  const wasHidden = storeSheet.isSheetHidden();
  if (wasHidden) storeSheet.showSheet();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(storeSheet);
  SpreadsheetApp.flush();
  try {
    const lastRow = storeSheet.getLastRow();
    const lastCol = storeSheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      storeSheet.getRange(1, 1, lastRow, lastCol).clear();
    }
    storeSheet.getImages().forEach(function(img) { try { img.remove(); } catch (e) {} });
    SpreadsheetApp.flush();
  } catch (e) {}
  if (wasHidden) storeSheet.hideSheet();
}

function bumpCatalogVersion_() {
  try {
    PropertiesService.getUserProperties().setProperty('tc_catalog_version', Date.now().toString());
  } catch (e) {}
}

/**
 * Diagnostic helper — run manually from the Apps Script editor.
 * Scans the entire active sheet for in-cell images (userEnteredValue.imageValue)
 * and over-grid images (getImages()). Results appear in View > Logs.
 */
function debugInCellImages() {
  Logger.log('=== In-Cell Image Debug ===');
  if (typeof Sheets === 'undefined') {
    Logger.log('ERROR: Sheets API service NOT enabled.');
    Logger.log('Go to Extensions > Apps Script > Services (+) > Google Sheets API > Add');
    return 'Sheets service not enabled';
  }
  Logger.log('Sheets API service: OK');

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var dataRange = sheet.getDataRange();
  var a1 = "'" + sheetName.replace(/'/g, "''") + "'!" + dataRange.getA1Notation();
  Logger.log('Scanning sheet: ' + sheetName + '  range: ' + dataRange.getA1Notation());

  try {
    var resp = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: [a1],
      fields: 'sheets(data(rowData(values(userEnteredValue))))',
    });
    var rows = ((resp.sheets || [])[0] || {}).data;
    rows = rows && rows[0] && rows[0].rowData ? rows[0].rowData : [];
    var found = 0;
    rows.forEach(function(row, ri) {
      (row.values || []).forEach(function(cell, ci) {
        var iv = cell.userEnteredValue && cell.userEnteredValue.imageValue;
        if (iv) {
          found++;
          Logger.log('  in-cell image at row=' + (dataRange.getRow() + ri) +
            ' col=' + (dataRange.getColumn() + ci) +
            ' data=' + JSON.stringify(iv).substring(0, 300));
        }
      });
    });
    Logger.log('Total in-cell images (API): ' + found);
  } catch (e) {
    Logger.log('ERROR reading sheet data: ' + e.message);
  }

  var overGrid = sheet.getImages();
  Logger.log('Over-grid images (getImages()): ' + overGrid.length);
  overGrid.forEach(function(img, i) {
    var a = img.getAnchorCell();
    Logger.log('  [' + i + '] row=' + a.getRow() + ' col=' + a.getColumn() +
      ' ' + img.getWidth() + 'x' + img.getHeight());
    try {
      var blob = img.getBlob();
      Logger.log('    getBlob(): OK, ' + blob.getBytes().length + ' bytes, ' + blob.getContentType());
    } catch (e) {
      Logger.log('    getBlob(): FAILED - ' + e.message);
      var fallbackBlob = getOverGridImageBlob_(img, sheet);
      if (fallbackBlob) {
        Logger.log('    Sheets API fallback: OK, ' + fallbackBlob.getBytes().length + ' bytes');
      } else {
        Logger.log('    Sheets API fallback: FAILED');
      }
    }
    // Log available methods and test getUrl()
    var methods = [];
    ['getBlob','getAs','getUrl','getSourceUrl','getContentUrl','getImageUrl',
     'getAltTextTitle','getAltTextDescription','getCellImageMetadata'].forEach(function(m) {
      if (typeof img[m] === 'function') methods.push(m);
    });
    Logger.log('    available methods: ' + (methods.length ? methods.join(', ') : '(none of the above)'));
    if (typeof img.getUrl === 'function') {
      try {
        var url = img.getUrl();
        Logger.log('    getUrl(): ' + (url ? url.substring(0, 120) : 'null/empty'));
        if (url) {
          var fetchTest = UrlFetchApp.fetch(url, {
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true,
          });
          Logger.log('    fetch via getUrl(): HTTP ' + fetchTest.getResponseCode() +
            (fetchTest.getResponseCode() === 200 ? ' OK, ' + fetchTest.getContent().length + ' bytes' : ''));
        }
      } catch (e) {
        Logger.log('    getUrl() error: ' + e.message);
      }
    }
  });

  Logger.log('=== End Debug ===');
  return 'Done — check Logs';
}

/**
 * Diagnostic: tests XLSX image extraction on the active sheet.
 * Run from Apps Script editor while the template sheet is active.
 */
function debugXlsxImages() {
  Logger.log('=== XLSX Image Extraction Debug ===');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  Logger.log('Sheet: ' + sheet.getName() + ' (id=' + sheet.getSheetId() + ')');

  // Step 1: download XLSX
  var token = ScriptApp.getOAuthToken();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  Logger.log('Downloading XLSX...');
  var resp;
  try {
    resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  } catch (e) {
    Logger.log('FAILED to fetch: ' + e.message);
    return;
  }
  Logger.log('HTTP ' + resp.getResponseCode() + '  size=' + resp.getContent().length + ' bytes');
  if (resp.getResponseCode() !== 200) { Logger.log('Download failed'); return; }

  // Step 2: unzip
  var files = {};
  try {
    Utilities.unzip(resp.getBlob().setContentType('application/zip')).forEach(function(f) { files[f.getName()] = f; });
  } catch (e) {
    Logger.log('Unzip FAILED: ' + e.message);
    return;
  }
  var allNames = Object.keys(files).sort();
  Logger.log('Files in ZIP: ' + allNames.length);
  allNames.forEach(function(n) { Logger.log('  ' + n); });

  // Step 3: find sheet index
  var sheets = ss.getSheets();
  var sheetIdx = 0;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheet.getSheetId()) { sheetIdx = i + 1; break; }
  }
  Logger.log('Sheet index (1-based): ' + sheetIdx);

  // Step 4: call buildXlsxImageMap_
  var map = buildXlsxImageMap_(sheet);
  var keys = Object.keys(map);
  Logger.log('Images found by buildXlsxImageMap_: ' + keys.length);
  keys.forEach(function(k) {
    try { Logger.log('  key=' + k + '  bytes=' + map[k].getBytes().length + '  type=' + map[k].getContentType()); }
    catch(e) { Logger.log('  key=' + k + '  (getBytes failed: ' + e.message + ')'); }
  });

  Logger.log('=== End XLSX Debug ===');
  return 'Done — check Logs';
}

/**
 * Diagnostic helper — run manually from the Apps Script editor (Extensions > Apps Script).
 * Tests captureTemplateImages_ (getImages) on the active sheet's data range.
 * Results appear in View > Logs.
 */
function debugImageCapture() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();

  Logger.log('=== Image Debug ===');
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('getImages() count: ' + sheet.getImages().length);
  sheet.getImages().forEach(function(img, i) {
    var anchor = img.getAnchorCell();
    Logger.log('  [' + i + '] row=' + anchor.getRow() + ' col=' + anchor.getColumn() +
      ' ' + img.getWidth() + 'x' + img.getHeight() +
      ' xOff=' + img.getAnchorCellXOffset() + ' yOff=' + img.getAnchorCellYOffset());
  });

  Logger.log('captureTemplateImages_ on ' + range.getA1Notation() + '...');
  var captured = captureTemplateImages_(sheet, range);
  Logger.log('Captured count: ' + captured.length);
  captured.forEach(function(img, i) {
    Logger.log('  [' + i + '] relRow=' + img.relRow + ' relCol=' + img.relCol +
      ' ' + img.width + 'x' + img.height + ' ' + img.mimeType +
      ' base64len=' + (img.base64 ? img.base64.length : 0));
  });

  Logger.log('=== End Debug ===');
  return 'Done — check Logs';
}
