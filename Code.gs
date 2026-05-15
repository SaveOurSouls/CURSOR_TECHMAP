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
    'imagesJson',
  ],
};

/**
 * Возвращает полное имя текущего листа.
 * Используется в шапке техкарты как =SHEETNAME() или =GET_NAME(...)
 * @customfunction
 */
function SHEETNAME() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Возвращает полное имя текущего листа.
 * Аргумент принимается для совместимости с формулой =GET_NAME(Таблица1[Индекс сборки]),
 * но не используется — имя берётся из названия листа.
 * @param {*} _index Любое значение (игнорируется).
 * @customfunction
 */
function GET_NAME(_index) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Возвращает часть имени листа до разделителя " | ".
 * Пример: "CUT_WIRE_auto | Резка" → "CUT_WIRE_auto"
 * @customfunction
 */
function SHEET_CODE() {
  const name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const idx = name.indexOf(' | ');
  return idx >= 0 ? name.substring(0, idx) : name;
}

/**
 * Возвращает часть имени листа после разделителя " | ".
 * Пример: "CUT_WIRE_auto | Резка" → "Резка"
 * @customfunction
 */
function SHEET_TYPE() {
  const name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const idx = name.indexOf(' | ');
  return idx >= 0 ? name.substring(idx + 3) : '';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(TECHMAP_APP.menuTitle)
    .addItem('Открыть рабочую панель', 'showWorkspaceSidebar')
    .addItem('Сохранить выделение как шаблон', 'showSaveTemplateDialog')
    .addToUi();
}

function showWorkspaceSidebar() {
  const template = HtmlService.createTemplateFromFile('WorkspaceSidebar');
  template.initialTab = 'templates';
  const html = template
    .evaluate()
    .setTitle('Техкарты и материалы')
    .setWidth(700);

  SpreadsheetApp.getUi().showSidebar(html);
}

function showSaveTemplateDialog() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const range = sheet ? sheet.getActiveRange() : null;
  const sheetName = sheet ? sheet.getName() : '';

  const state = {
    selection: range && !isSystemSheet_(sheetName) ? {
      sheetName,
      rangeA1: range.getA1Notation(),
      height: range.getNumRows(),
      width: range.getNumColumns(),
    } : null,
    templates: readCatalog_().map((item) => ({
      id: item.id,
      title: item.title,
      category: item.category,
      description: item.description,
      sizeLabel: `${item.height} x ${item.width}`,
    })),
  };

  const tmpl = HtmlService.createTemplateFromFile('SaveTemplateDialog');
  tmpl.initialState = JSON.stringify(state);

  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(460).setHeight(430),
    'Сохранить шаблон'
  );
}

function getCatalogVersion() {
  return PropertiesService.getUserProperties().getProperty('tc_catalog_version') || '0';
}

function getTemplateCatalog() {
  return readCatalog_().map((item) => ({
    id: item.id,
    title: item.title,
    category: item.category,
    description: item.description,
    sizeLabel: `${item.height} x ${item.width}`,
    updatedAt: item.updatedAt,
  }));
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
  const storeSheet = ensureStoreSheet_(SpreadsheetApp.getActive());
  if (existingTemplate) {
    try { deleteTemplateImages_(existingTemplate.imagesJson || '[]'); } catch (e) {}
  }
  const imagesData = runWithSheetVisible_(storeSheet, () => {
    return copySourceImagesToStore_(range.getSheet(), range, storeSheet, storeLocation.row, storeLocation.column);
  }) || [];

  const catalogSheet = ensureCatalogSheet_(SpreadsheetApp.getActive());
  runWithSheetVisible_(catalogSheet, () => {
    let rowHeightsJson = '[]';
    let columnWidthsJson = '[]';
    try {
      rowHeightsJson = JSON.stringify(getRowHeights_(range));
    } catch (e) {}
    try {
      columnWidthsJson = JSON.stringify(getColumnWidths_(range));
    } catch (e) {}
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
      rowHeightsJson,
      columnWidthsJson,
      imagesJson: imagesData.length ? JSON.stringify(imagesData) : '[]',
    });
  });

  bumpCatalogVersion_();

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
  const template = getTemplateById_(templateId);

  // Create a new sheet named after the template; insert content starting at B2.
  const targetSheet = createUniqueSheet_(ss, template.title);
  const targetRow = 2;
  const targetColumn = 2;

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

  // insertTemplateImages_ reads blobs from Drive (fast). If Drive-cached images exist,
  // skip the slow insertOverGridImages_ STORE scan (which downloads XLSX as fallback).
  // For old templates without driveFileId, fall back to STORE scan for backward compat.
  const parsedImages = parseJsonArray_(template.imagesJson);
  const hasDriveImages = parsedImages.some((img) => img.driveFileId);
  if (parsedImages.length) {
    insertTemplateImages_(targetSheet, targetRow, targetColumn, template.imagesJson);
  }
  if (!hasDriveImages) {
    runWithSheetVisible_(sourceSheet, () => {
      insertOverGridImages_(
        sourceSheet, template.storeRow, template.storeColumn,
        template.height, template.width,
        targetSheet, targetRow, targetColumn
      );
    });
  }

  ss.setActiveSheet(targetSheet);

  return {
    title: template.title,
    sheetName: targetSheet.getName(),
    insertedRange: targetRange.getA1Notation(),
  };
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

  const allRows = catalogSheet.getRange(2, 1, lastRow - 1, TECHMAP_APP.catalogHeaders.length).getValues();
  const rowIndex = allRows.findIndex((r) => String(r[0]).trim() === templateId);
  if (rowIndex < 0) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }

  const rawRow = allRows[rowIndex];
  const title = String(rawRow[1] || templateId);
  const storeRow = Number(rawRow[4]) || 0;
  const height = Number(rawRow[6]) || 0;

  try { deleteTemplateImages_(rawRow[13] || '[]'); } catch (e) {}

  // Erase STORE slot — show sheet, set active, flush BEFORE clear so the server
  // sees a visible sheet; flush AFTER clear to commit before hiding again.
  const storeSheet = ss.getSheetByName(TECHMAP_APP.storeSheetName);
  if (storeSheet && storeRow > 0 && height > 0) {
    const cols = Math.max(Number(rawRow[7]) || 1, 1);
    const priorActive = ss.getActiveSheet();
    const wasHidden = storeSheet.isSheetHidden();
    if (wasHidden) storeSheet.showSheet();
    ss.setActiveSheet(storeSheet);
    SpreadsheetApp.flush();
    clearStoreSlotImages_(storeSheet, storeRow, 1, height, cols);
    try {
      storeSheet.getRange(storeRow, 1, height, cols).clear();
    } catch (e) {
      try { storeSheet.getRange(storeRow, 1, height, cols).clearContent(); } catch (e2) {}
      try { storeSheet.getRange(storeRow, 1, height, cols).clearFormat(); } catch (e3) {}
    }
    SpreadsheetApp.flush();
    if (wasHidden) storeSheet.hideSheet();
    try {
      if (priorActive && !isSystemSheet_(priorActive.getName())) {
        ss.setActiveSheet(priorActive);
      }
    } catch (e) {}
  }

  catalogSheet.deleteRow(rowIndex + 2);
  bumpCatalogVersion_();

  // If catalog is now empty, purge any orphaned content still in STORE.
  if (storeSheet && readCatalog_().length === 0) {
    purgeEntireStore_(storeSheet);
  }

  return { deleted: true, id: templateId, title };
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
