const TECHMAP_APP = {
  menuTitle: 'Техкарты',
  librarySheetName: '_TC_LIBRARY',
  storeSheetName: '_TC_STORE',
  canvasSheetName: '_TC_CANVAS',
  legacyTemplatePrefix: '_TPL_',
  catalogHeaders: [
    'id',
    'title',
    'category',
    'description',
    'storeRow',
    'storeColumn',
    'height',
    'width',
    'sourceSheet',
    'sourceRange',
    'updatedAt',
    'rowHeightsJson',
    'columnWidthsJson',
  ],
};

if (
  typeof TECHMAP_DATA_MODEL !== 'undefined' &&
  TECHMAP_DATA_MODEL.catalogHeaders &&
  Array.isArray(TECHMAP_DATA_MODEL.catalogHeaders)
) {
  TECHMAP_APP.catalogHeaders = TECHMAP_DATA_MODEL.catalogHeaders.slice();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(TECHMAP_APP.menuTitle)
    .addItem('Открыть рабочую панель', 'showWorkspaceSidebar')
    .addItem('Сохранить выделение как шаблон', 'showSaveTemplateDialog')
    .addSeparator()
    .addItem('Проверить доступ к внешним БД', 'authorizeExternalDatabases')
    .addItem('Обновить все базы', 'refreshAllDatabases')
    .addItem('Обновить базу материалов', 'refreshMaterialsDatabase')
    .addItem('Обновить базу техопераций', 'refreshOperationDatabase')
    .addToUi();
}

function showWorkspaceSidebar() {
  const template = HtmlService.createTemplateFromFile('WorkspaceSidebar');
  template.initialTab = 'templates';
  const html = template
    .evaluate()
    .setTitle('Техкарты и материалы')
    .setWidth(420);

  SpreadsheetApp.getUi().showSidebar(html);
}

function refreshAllDatabases() {
  const ui = SpreadsheetApp.getUi();
  const matSummary = syncMaterialDatabase();
  const opsSummary = syncTechOperationsDatabase();
  ui.alert(
    'Все базы обновлены.',
    `Материалы: ${matSummary.recordCount} позиций\nТехоперации: ${opsSummary.recordCount} строк`,
    ui.ButtonSet.OK
  );
}

function authorizeExternalDatabases() {
  const ui = SpreadsheetApp.getUi();
  const materialName = SpreadsheetApp.openById(MATERIAL_DB_APP.sourceSpreadsheetId).getName();
  const techOpsName = SpreadsheetApp.openById(TECHOPS_DB_APP.sourceSpreadsheetId).getName();

  ui.alert(
    'Доступ к внешним БД проверен.',
    `Материалы: ${materialName}\nТехоперации: ${techOpsName}`,
    ui.ButtonSet.OK
  );
}

function showTemplateSidebar() {
  const ss = SpreadsheetApp.getActive();
  ensureInfrastructure_(ss);
  hideLegacyTemplateSheets_(ss);
  hideLibrarySheets();

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Библиотека техкарт')
    .setWidth(360);

  SpreadsheetApp.getUi().showSidebar(html);
}

function showSaveTemplateDialog() {
  ensureInfrastructure_();

  const html = HtmlService.createHtmlOutputFromFile('SaveTemplateDialog')
    .setWidth(460)
    .setHeight(430);

  SpreadsheetApp.getUi().showModalDialog(html, 'Сохранить шаблон');
}

function initializeDemoLibrary() {
  SpreadsheetApp.getUi().alert(
    'Демо-шаблоны из кода не используются.',
    'Нарисуйте блок на листе и сохраните его через «Техкарты → Сохранить выделение как шаблон», либо откройте «Открыть рабочую панель».',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getTemplateCatalog() {
  ensureInfrastructure_();
  return readCatalog_().map((item) => ({
    id: item.id,
    title: item.title,
    category: item.category,
    description: item.description,
    sizeLabel: `${item.height} x ${item.width}`,
    updatedAt: item.updatedAt,
  }));
}

function getSaveTemplateDialogState() {
  ensureInfrastructure_();

  const selection = getActiveWorkingRange_();
  return {
    selection: {
      sheetName: selection.getSheet().getName(),
      rangeA1: selection.getA1Notation(),
      height: selection.getNumRows(),
      width: selection.getNumColumns(),
    },
    templates: readCatalog_().map((item) => ({
      id: item.id,
      title: item.title,
      category: item.category,
      description: item.description,
      sizeLabel: `${item.height} x ${item.width}`,
    })),
  };
}

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

function saveSelectedRangeAsTemplate(formData) {
  ensureInfrastructure_();

  const range = resolveTemplateSourceRange_(formData);
  const title = normalizeString_(formData && formData.title);
  if (!title) {
    throw new Error('Укажите название шаблона.');
  }

  const category = normalizeString_(formData && formData.category);
  const description = normalizeString_(formData && formData.description);
  const existingTemplateId = normalizeString_(formData && formData.existingTemplateId);
  const catalog = readCatalog_();
  const existingTemplate = existingTemplateId
    ? catalog.find((item) => item.id === existingTemplateId)
    : null;

  const recordId = existingTemplate ? existingTemplate.id : makeTemplateId_(title, catalog);
  const storeLocation = allocateStoreLocation_(range, existingTemplate, catalog);
  writeRangeToStore_(range, storeLocation.row, storeLocation.column);

  const catalogSheet = ensureCatalogSheet_(SpreadsheetApp.getActive());
  runWithSheetVisible_(catalogSheet, () => {
    upsertCatalogRecord_(catalogSheet, {
      id: recordId,
      title,
      category,
      description,
      storeRow: storeLocation.row,
      storeColumn: storeLocation.column,
      height: range.getNumRows(),
      width: range.getNumColumns(),
      sourceSheet: range.getSheet().getName(),
      sourceRange: range.getA1Notation(),
      updatedAt: new Date().toISOString(),
      rowHeightsJson: JSON.stringify(getRowHeights_(range)),
      columnWidthsJson: JSON.stringify(getColumnWidths_(range)),
    });
  });

  hideLibrarySheets();

  return {
    action: existingTemplate ? 'updated' : 'created',
    id: recordId,
    title,
    sizeLabel: `${range.getNumRows()} x ${range.getNumColumns()}`,
  };
}

function insertTemplate(templateId) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }

  const ss = SpreadsheetApp.getActive();

  // Capture target sheet and cell BEFORE any infrastructure calls that might
  // change the active sheet (canvas creation, compaction, hide/show operations).
  const targetSheet = ss.getActiveSheet();
  if (isSystemSheet_(targetSheet.getName())) {
    throw new Error('Вставка шаблонов на служебные листы запрещена. Перейдите на рабочий лист.');
  }
  const activeRange = targetSheet.getActiveRange();
  if (!activeRange) {
    throw new Error('Не выбрана ячейка для вставки.');
  }
  const targetRow = activeRange.getRow();
  const targetColumn = activeRange.getColumn();

  // Only ensure infrastructure (catalog + store sheets exist). Do NOT call
  // Не вызывать сюда автосоздание демо — ломает активный лист и склад.
  ensureInfrastructure_();

  const template = getTemplateById_(templateId);
  const sourceSheet = ensureStoreSheet_(ss);
  const sourceRange = sourceSheet.getRange(
    template.storeRow,
    template.storeColumn,
    template.height,
    template.width
  );

  ensureSheetCapacity_(
    targetSheet,
    targetRow + template.height - 1,
    targetColumn + template.width - 1
  );
  applyStoredDimensions_(targetSheet, targetRow, targetColumn, template);

  const targetRange = targetSheet.getRange(
    targetRow,
    targetColumn,
    template.height,
    template.width
  );
  targetRange.breakApart();
  copyRangePreservingFormulas_(sourceRange, targetRange);
  clearTemplateMarkerNote_(targetRange);

  SpreadsheetApp.flush();

  return {
    title: template.title,
    sheetName: targetSheet.getName(),
    insertedRange: targetRange.getA1Notation(),
  };
}

function showLibrarySheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach((sheet) => {
    if (isSystemSheet_(sheet.getName())) {
      sheet.showSheet();
    }
  });
}

function hideLibrarySheets() {
  const ss = SpreadsheetApp.getActive();
  const activeSheet = ss.getActiveSheet();
  let fallbackSheet = null;

  ss.getSheets().forEach((sheet) => {
    if (!isSystemSheet_(sheet.getName()) && !fallbackSheet) {
      fallbackSheet = sheet;
    }
  });

  if (activeSheet && isSystemSheet_(activeSheet.getName()) && fallbackSheet) {
    ss.setActiveSheet(fallbackSheet);
  }

  ss.getSheets().forEach((sheet) => {
    if (isSystemSheet_(sheet.getName())) {
      sheet.hideSheet();
    }
  });
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

function ensureStoreSheet_(ss) {
  let sheet = ss.getSheetByName(TECHMAP_APP.storeSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHMAP_APP.storeSheetName);
    sheet.hideSheet();
  }
  return sheet;
}

function readCatalog_() {
  const sheet = ensureCatalogSheet_(SpreadsheetApp.getActive());
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, TECHMAP_APP.catalogHeaders.length)
    .getValues()
    .filter((row) => row[0])
    .map((row) => ({
      id: row[0],
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
  ]];

  const targetRow = existingIndex >= 0 ? existingIndex + 2 : lastRow + 1;
  catalogSheet.getRange(targetRow, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues(rowValues);
}

function deleteTemplate(templateId) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }

  const ss = SpreadsheetApp.getActive();

  const catalogSheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!catalogSheet) {
    throw new Error('Каталог шаблонов не найден.');
  }

  const lastRow = catalogSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }

  // Читаем весь каталог одним вызовом, чтобы исключить гонку состояний
  const allRows = catalogSheet.getRange(2, 1, lastRow - 1, TECHMAP_APP.catalogHeaders.length).getValues();
  const rowIndex = allRows.findIndex((r) => String(r[0]).trim() === templateId);
  if (rowIndex < 0) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }

  const rawRow = allRows[rowIndex];
  const title = String(rawRow[1] || templateId);
  const storeRow = Number(rawRow[4]) || 0;
  const height = Number(rawRow[6]) || 0;

  // Удаляем строки из _TC_STORE полностью (а не только контент)
  if (storeRow > 0 && height > 0) {
    const storeSheet = ss.getSheetByName(TECHMAP_APP.storeSheetName);
    if (storeSheet) {
      const storeMaxRow = storeSheet.getLastRow();
      const deleteFrom = storeRow;
      const deleteCount = Math.min(height, storeMaxRow - deleteFrom + 1);
      if (deleteCount > 0) {
        storeSheet.deleteRows(deleteFrom, deleteCount);
        // Сдвигаем storeRow в каталоге для всех шаблонов, чьи строки ниже удалённых
        for (let i = 0; i < allRows.length; i += 1) {
          if (i === rowIndex) {
            continue;
          }
          const otherStoreRow = Number(allRows[i][4]) || 0;
          if (otherStoreRow > storeRow) {
            const catalogRow = i + 2;
            catalogSheet.getRange(catalogRow, 5).setValue(otherStoreRow - deleteCount);
          }
        }
      }
    }
  }

  // Удаляем строку из каталога
  catalogSheet.deleteRow(rowIndex + 2);

  return { deleted: true, id: templateId, title };
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
  // Reuse exact slot when size matches
  if (
    existingTemplate &&
    existingTemplate.height === range.getNumRows() &&
    existingTemplate.width === range.getNumColumns()
  ) {
    return { row: existingTemplate.storeRow, column: existingTemplate.storeColumn };
  }

  // Compact the store first so there are no phantom empty rows from old deletions
  compactifyStore_(catalog);

  // After compaction catalog storeRow values are now consecutive;
  // place new template right after the last live block.
  const freshCatalog = readCatalog_();
  const nextRow = freshCatalog.reduce((maxRow, item) => {
    if (existingTemplate && item.id === existingTemplate.id) {
      return maxRow; // skip the template being replaced
    }
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
 * Should be called before inserting a new template block.
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
  // Sort live templates by their current storeRow so we rewrite top-to-bottom
  const live = (catalog || [])
    .filter((item) => item.storeRow > 0 && item.height > 0 && item.width > 0)
    .sort((a, b) => a.storeRow - b.storeRow);

  if (!live.length) {
    // Nothing to keep — clear everything
    const maxRow = storeSheet.getLastRow();
    if (maxRow > 0) {
      storeSheet.deleteRows(1, maxRow);
    }
    return;
  }

  // Check whether store is already compact (no gaps, rows start at 1)
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

  // Build compacted content by copying each live block in order
  // We do this on a temp hidden sheet to avoid overwriting source while reading
  const tempName = '_TC_COMPACT_TMP';
  let tempSheet = ss.getSheetByName(tempName);
  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }
  tempSheet = ss.insertSheet(tempName);
  tempSheet.hideSheet();

  let writeRow = 1;
  const newStoreRows = {}; // id -> new storeRow

  live.forEach((item) => {
    const srcRange = storeSheet.getRange(item.storeRow, item.storeColumn, item.height, item.width);
    ensureSheetCapacity_(tempSheet, writeRow + item.height - 1, item.width);
    const destRange = tempSheet.getRange(writeRow, 1, item.height, item.width);
    destRange.breakApart();
    copyRangePreservingFormulas_(srcRange, destRange);
    newStoreRows[item.id] = writeRow;
    writeRow += item.height;
  });

  // Clear old store and copy compacted content back
  const storeLastRow = storeSheet.getLastRow();
  if (storeLastRow > 0) {
    storeSheet.deleteRows(1, storeLastRow);
  }
  if (writeRow > 1) {
    const cols = live[0].width || 20;
    ensureSheetCapacity_(storeSheet, writeRow - 1, cols);
    const compactedRange = tempSheet.getRange(1, 1, writeRow - 1, cols);
    copyRangePreservingFormulas_(compactedRange, storeSheet.getRange(1, 1, writeRow - 1, cols));
  }

  ss.deleteSheet(tempSheet);

  // Update storeRow in catalog for all items that moved
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

  // Restore the active sheet that was active before compaction
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
  }
}

/**
 * Copies a range to a target range without adjusting formula references.
 * Writes values and formulas first on a broken-apart grid, then pastes format
 * (including merges) last. PASTE_FORMAT before setValues recreates merges and
 * makes setValues hit merged cells — that triggers "Ошибка службы: Таблицы" and
 * can leave blocks without visible text.
 */
function copyRangePreservingFormulas_(sourceRange, targetRange) {
  const srcSheet = sourceRange.getSheet();
  const dstSheet = targetRange.getSheet();
  const hideSrc = srcSheet.isSheetHidden();
  const hideDst = dstSheet.isSheetHidden();
  if (hideSrc) {
    srcSheet.showSheet();
  }
  if (hideDst) {
    dstSheet.showSheet();
  }

  const ss = SpreadsheetApp.getActive();
  const priorActive = ss.getActiveSheet();

  try {
    targetRange.breakApart();

    const values   = sourceRange.getValues();
    const formulas = sourceRange.getFormulas();

    targetRange.setValues(values);

    formulas.forEach((row, r) => {
      row.forEach((formula, c) => {
        if (formula !== '') {
          targetRange.getCell(r + 1, c + 1).setFormula(formula);
        }
      });
    });

    copyRangeFormatPreservingMerges_(sourceRange, targetRange, ss, priorActive, srcSheet, dstSheet);
  } finally {
    if (hideSrc) {
      srcSheet.hideSheet();
    }
    if (hideDst) {
      dstSheet.hideSheet();
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
      } catch (e3) {
        // В целевом диапазоне уже есть значения и формулы.
      }
    }
  } finally {
    restorePrior();
  }
}

function applyStoredDimensions_(targetSheet, targetRow, targetColumn, template) {
  const rowHeights = template.rowHeights || [];
  rowHeights.forEach((height, index) => {
    if (height) {
      targetSheet.setRowHeight(targetRow + index, height);
    }
  });

  const columnWidths = template.columnWidths || [];
  columnWidths.forEach((width, index) => {
    if (width) {
      targetSheet.setColumnWidth(targetColumn + index, width);
    }
  });
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
  if (note && note.indexOf('techmap-template:') === 0) {
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
