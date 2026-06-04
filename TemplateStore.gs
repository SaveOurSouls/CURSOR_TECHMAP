// ============================================================
//  TemplateStore.gs — каталог шаблонов и хранилище _TC_STORE
//  Зависимости: Config.gs, Utils.gs, RangeCopy.gs, ImageHandler.gs
// ============================================================

// ── Range resolution & validation ───────────────────────────

/**
 * Диапазон для сохранения: из данных диалога (на момент открытия), иначе текущее выделение.
 * Повторный getActiveWorkingRange_ при открытом модальном окне нестабилен.
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

// ── Infrastructure ───────────────────────────────────────────

function ensureInfrastructure_(ssArg) {
  const ss = ssArg || SpreadsheetApp.getActive();
  ensureCatalogSheet_(ss);
  ensureStoreSheet_(ss);
}

function ensureCatalogSheet_(ss) {
  let sheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHMAP_APP.librarySheetName);
    writeSheetHeader_(sheet, TECHMAP_APP.catalogHeaders, '#d9e2f3');
    sheet.hideSheet();
  }
  return sheet;
}

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

// ── Catalog CRUD ─────────────────────────────────────────────

// Кеш каталога на время одного выполнения скрипта (глобалы GAS сбрасываются
// между вызовами). Генерация зовёт insertTemplate→readCatalog_ на каждую
// операцию — без кеша это N полных чтений _TC_LIBRARY за один прогон.
// Инвалидируется при любой записи в каталог (invalidateCatalogCache_).
var _catalogCache_ = null;

function invalidateCatalogCache_() {
  _catalogCache_ = null;
}

function readCatalog_() {
  if (_catalogCache_) return _catalogCache_;
  const sheet = getCatalogSheetIfExists_(SpreadsheetApp.getActive());
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { _catalogCache_ = []; return _catalogCache_; }

  _catalogCache_ = sheet
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
  return _catalogCache_;
}

/**
 * Преобразует внутреннюю запись каталога в DTO для sidebar/диалога.
 * Единый контракт списка шаблонов для UI.
 */
function toCatalogListItem_(item) {
  return {
    id: item.id,
    title: item.title,
    category: item.category,
    description: item.description,
    sizeLabel: `${item.height} x ${item.width}`,
    updatedAt: item.updatedAt,
  };
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
  invalidateCatalogCache_();
}

function getTemplateById_(templateId) {
  const template = readCatalog_().find((item) => item.id === templateId);
  if (!template) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }
  return template;
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

// ── Store slot management ────────────────────────────────────

function allocateStoreLocation_(range, existingTemplate, catalog) {
  if (
    existingTemplate &&
    existingTemplate.height === range.getNumRows() &&
    existingTemplate.width === range.getNumColumns()
  ) {
    return { row: existingTemplate.storeRow, column: existingTemplate.storeColumn };
  }

  // Добавляем после последнего живого блока. Очищенные слоты не уплотняются —
  // store растёт медленно, но сохранение остаётся быстрым.
  const nextRow = (catalog || []).reduce((maxRow, item) => {
    if (existingTemplate && item.id === existingTemplate.id) return maxRow;
    return Math.max(maxRow, item.storeRow + item.height - 1);
  }, 0) + 1;

  return { row: nextRow, column: 1 };
}

function writeRangeToStore_(sourceRange, storeRow, storeColumn) {
  const storeSheet = ensureStoreSheet_(SpreadsheetApp.getActive());
  const height = sourceRange.getNumRows();
  const width = sourceRange.getNumColumns();
  ensureSheetCapacity_(storeSheet, storeRow + height - 1, storeColumn + width - 1);

  runWithSheetVisible_(storeSheet, () => {
    const targetRange = storeSheet.getRange(storeRow, storeColumn, height, width);
    clearStoreSlotForWrite_(targetRange); // снимет объединения и очистит слот
    SpreadsheetApp.flush();
    copyRangePreservingFormulas_(sourceRange, targetRange);
    targetRange.getCell(1, 1).setNote('techmap-template-store');
  });
}

/** Полностью очищает слот _TC_STORE (контент + формат + валидации + объединения). */
function clearStoreSlotForWrite_(targetRange) {
  // 1) Снять объединения — иначе clearContent на части объединённой ячейки падает.
  try { targetRange.breakApart(); } catch (e0) {}
  // 2) ВАЖНО: clear() БЕЗ аргументов чистит контент+формат+валидации. Прежний
  //    clear({contentsOnly:false}) был no-op (опции-флаги: ничего не true → ничего не
  //    чистится) — при записи маскировалось последующим copy, но при УДАЛЕНИИ слот
  //    оставался с данными (строка каталога ушла, данные в сторе остались).
  try {
    targetRange.clear();
  } catch (e) {
    // Фолбэк по категориям, если full clear падает на скрытом листе.
    try { targetRange.clearContent(); } catch (e2) {}
    try { targetRange.clearFormat(); } catch (e3) {}
    try { targetRange.clearDataValidations(); } catch (e4) {}
    try { targetRange.removeCheckboxes(); } catch (e5) {}
  }
}

// ── Dimension helpers ────────────────────────────────────────

function applyStoredDimensions_(targetSheet, targetRow, targetColumn, template) {
  applyRunLengthDimensions_(template.rowHeights || [], targetRow,
    (start, count, size) => targetSheet.setRowHeights(start, count, size));
  applyRunLengthDimensions_(template.columnWidths || [], targetColumn,
    (start, count, size) => targetSheet.setColumnWidths(start, count, size));
}

// Читает высоты строк И ширины колонок диапазона ОДНИМ вызовом Sheets API
// (rowMetadata/columnMetadata.pixelSize) вместо N поячеечных getRowHeight/getColumnWidth
// (на шаблоне 40×19 это ~59 round-trip → 1). Возвращает {rows, cols} или null при
// недоступности/несовпадении размеров → вызывающий падает на проверенный поэлементный цикл.
function readRangeDimensionsViaSheetsApi_(range) {
  if (typeof Sheets === 'undefined') return null;
  try {
    const ss    = SpreadsheetApp.getActive();
    const sheet = range.getSheet();
    const a1    = "'" + sheet.getName().replace(/'/g, "''") + "'!" + range.getA1Notation();
    const resp  = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: [a1],
      fields: 'sheets(data(rowMetadata(pixelSize),columnMetadata(pixelSize)))',
    });
    const data = (((resp.sheets || [])[0] || {}).data || [])[0] || {};
    const rows = (data.rowMetadata    || []).map((m) => (m && m.pixelSize) || 0);
    const cols = (data.columnMetadata || []).map((m) => (m && m.pixelSize) || 0);
    // Размеры не сошлись (API вернул не то, что ждём) — не доверяем, идём в фолбэк.
    if (rows.length !== range.getNumRows() || cols.length !== range.getNumColumns()) return null;
    return { rows, cols };
  } catch (e) {
    return null;
  }
}

function getRowHeights_(range) {
  const api = readRangeDimensionsViaSheetsApi_(range);
  if (api) return api.rows;
  const sheet = range.getSheet();
  const heights = [];
  for (let index = 0; index < range.getNumRows(); index += 1) {
    heights.push(sheet.getRowHeight(range.getRow() + index));
  }
  return heights;
}

function getColumnWidths_(range) {
  const api = readRangeDimensionsViaSheetsApi_(range);
  if (api) return api.cols;
  const sheet = range.getSheet();
  const widths = [];
  for (let index = 0; index < range.getNumColumns(); index += 1) {
    widths.push(sheet.getColumnWidth(range.getColumn() + index));
  }
  return widths;
}

function clearTemplateMarkerNote_(targetRange) {
  // Top-left always carries the store marker note (writeRangeToStore_ stamps it),
  // so clear unconditionally — avoids a per-insert getNote() round-trip.
  targetRange.getCell(1, 1).setNote('');
}

