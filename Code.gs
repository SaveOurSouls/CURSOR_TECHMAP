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
    .addSeparator()
    .addItem('Генератор техкарт сборки', 'showAssemblyGeneratorDialog')
    .addSeparator()
    .addItem('Диагностика: Sheets API', 'diagnoseSheetsApi')
    .addItem('Диагностика: вставка строк', 'testInsertApi')
    .addItem('Диагностика: последняя вставка', 'showLastInsertOutcome')
    .addToUi();
}

/** Показывает исход последней вставки строк при генерации (api-ok / FALLBACK: ...). */
function showLastInsertOutcome() {
  const v = PropertiesService.getDocumentProperties().getProperty('tc_last_insert');
  SpreadsheetApp.getUi().alert('Последняя вставка строк',
    v || 'Ещё не было вставок (сгенерируй многопроводную сборку и проверь снова).',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Проверяет, работает ли быстрая вставка строк через Sheets API на листе с
 * объединениями (горизонтальным и вертикальным). Показывает реальную ошибку,
 * если падает — чтобы понять, почему генерация откатывается на медленный путь.
 */
function testInsertApi() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.insertSheet('_TEST_INSERT_' + Date.now());
  let msg;
  try {
    sh.getRange(1, 1, 7, 4).setValues([
      ['hdr','hdr','hdr','hdr'],
      ['a','b','c','d'],
      ['LBL','n1','x','q'],   // строка 3 — верх вертикальной метки + шаблон строки
      ['','n2','y','q'],      // строка 4
      ['','n3','z','q'],      // строка 5
      ['','n4','w','q'],      // строка 6 — низ вертикальной метки
      ['z','z','z','z'],
    ]);
    sh.getRange(3, 1, 4, 1).merge();   // ВЕРТИКАЛЬНАЯ метка A3:A6 — пересекает точку вставки
    sh.getRange(4, 3, 1, 2).merge();   // горизонтальное объединение в строке-шаблоне (C4:D4)
    SpreadsheetApp.flush();

    insertRowsViaSheetsApi_(sh, 4, 2, 4); // вставить 2 строки после 4 (в середину метки), шаблон = строка 4
    msg = 'УСПЕХ ✓ — быстрая вставка работает. Причина мигания в другом.';
  } catch (e) {
    msg = 'ОШИБКА быстрой вставки:\n' + (e && e.message);
  } finally {
    try { ss.deleteSheet(sh); } catch (_) {}
  }
  ui.alert('Тест вставки строк', msg, ui.ButtonSet.OK);
}

/**
 * Диагностика доступности Google Sheets API (advanced service).
 * От него зависят быстрые пути: копирование шаблона, вставка строк (insertDimension),
 * скрытое копирование из _TC_STORE. Если API недоступен — всё откатывается на
 * медленные мигающие пути (разрыв/сборка объединений, copyTo с показом листа).
 * Запускать из меню «Техкарты → Диагностика», результат — во всплывающем окне.
 */
function diagnoseSheetsApi() {
  const ui = SpreadsheetApp.getUi();
  if (typeof Sheets === 'undefined') {
    ui.alert('Sheets API: НЕ ОБЪЯВЛЕН',
      'Advanced-сервис Google Sheets API не подключён.\n\n' +
      'Открой редактор скриптов → Службы (+) → добавь «Google Sheets API» → Сохрани.\n' +
      'Это включит быстрые пути (нет мигания, быстрее в разы).',
      ui.ButtonSet.OK);
    return 'undefined';
  }
  try {
    Sheets.Spreadsheets.get(SpreadsheetApp.getActive().getId(), { fields: 'spreadsheetId' });
    ui.alert('Sheets API: РАБОТАЕТ ✓',
      'Быстрые пути активны. Если мигание всё ещё есть — причина в другом, сообщи.',
      ui.ButtonSet.OK);
    return 'ok';
  } catch (e) {
    ui.alert('Sheets API: ОШИБКА',
      'Сервис объявлен, но вызов падает:\n' + (e && e.message) + '\n\n' +
      'Скорее всего Google Sheets API не включён в Cloud-проекте скрипта. ' +
      'Открой редактор → Службы → переподключи «Google Sheets API», или включи API в ' +
      'связанном проекте Google Cloud.',
      ui.ButtonSet.OK);
    return 'error';
  }
}


function showWorkspaceSidebar() {
  const template = HtmlService.createTemplateFromFile('WorkspaceSidebar');
  template.initialTab = 'templates';
  const html = template
    .evaluate()
    .setWidth(820)
    .setHeight(680);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Техкарты');
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
    templates: readCatalog_().map(toCatalogListItem_),
  };

  const tmpl = HtmlService.createTemplateFromFile('SaveTemplateDialog');
  tmpl.initialState = embedJsonForHtml_(state);

  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(460).setHeight(430),
    'Сохранить шаблон'
  );
}

/** @returns {string} Метка версии каталога для инвалидации кеша на стороне sidebar. */
function getCatalogVersion() {
  return PropertiesService.getUserProperties().getProperty('tc_catalog_version') || '0';
}

/** @returns {Object[]} Список шаблонов для отображения в sidebar. */
function getTemplateCatalog() {
  return readCatalog_().map(toCatalogListItem_);
}


/**
 * Сохраняет выделенный диапазон как шаблон.
 * @param {Object} formData — данные из диалога SaveTemplateDialog.
 * @returns {Object} { action, id, title, sizeLabel }
 */
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

/**
 * Вставляет шаблон на новый лист.
 * Вызывается как из sidebar, так и из generateAssemblyTechCards.
 * @param {string} templateId
 * @returns {Object} { title, sheetName, insertedRange }
 */
function insertTemplate(templateId) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }

  const ss = SpreadsheetApp.getActive();
  const template = getTemplateById_(templateId);

  // Remove existing sheet with the same name so re-generation doesn't produce "-2" suffix.
  const existingSheet = ss.getSheetByName(template.title);
  if (existingSheet) ss.deleteSheet(existingSheet);

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

/**
 * Удаляет шаблон из каталога и очищает его слот в _TC_STORE.
 * @param {string} templateId
 * @returns {Object} { deleted, id, title }
 */
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
  invalidateCatalogCache_();
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
