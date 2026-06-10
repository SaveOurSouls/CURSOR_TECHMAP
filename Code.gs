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
    .addToUi();
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
  return withDocumentLock_(function() { return saveSelectedRangeAsTemplateImpl_(formData); });
}

function saveSelectedRangeAsTemplateImpl_(formData) {
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
    // Слот переехал (изменились размеры → новый storeRow) — очистить СТАРЫЙ, иначе
    // его содержимое и плавающие картинки остаются орфаном в _TC_STORE. Новый слот
    // не пересекается со старым (allocateStoreLocation_ кладёт после всех чужих блоков).
    const moved = existingTemplate.storeRow !== storeLocation.row
      || existingTemplate.storeColumn !== storeLocation.column;
    if (moved) {
      try {
        purgeStoreSlot_(storeSheet, existingTemplate.storeRow,
          existingTemplate.storeColumn || 1, existingTemplate.height, existingTemplate.width);
      } catch (e) {}
    }
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
    // Сколько картинок не удалось перенести — диалог покажет предупреждение
    // (раньше потеря была тихой). 0/undefined = всё перенесено.
    skippedImages: typeof _lastImageCaptureSkipped === 'number' ? _lastImageCaptureSkipped : 0,
  };
}

/**
 * Вставляет шаблон на новый лист.
 * Вызывается как из sidebar, так и из generateAssemblyTechCards.
 * @param {string} templateId
 * @returns {Object} { title, sheetName, insertedRange }
 */
function insertTemplate(templateId, opts) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }
  opts = opts || {};

  const ss = SpreadsheetApp.getActive();
  const template = getTemplateById_(templateId);

  // Заменяем одноимённый лист, чтобы при ре-генерации не плодить суффиксы «-2».
  // Но сносим ТОЛЬКО лист-техкарту (имя в формате "CODE | Тип", содержит " | "),
  // чтобы случайно не удалить важный пользовательский лист с совпавшим именем.
  // opts.noReplace — генератор вставляет НЕСКОЛЬКО листов с одним шаблоном (N разъёмов):
  // замена по заголовку снесла бы предыдущий разъём; чистку ре-гена делает purgeGeneratedSheets_.
  if (!opts.noReplace) {
    const existingSheet = ss.getSheetByName(template.title);
    if (existingSheet && template.title.indexOf(' | ') >= 0) ss.deleteSheet(existingSheet);
  }

  // Create a new sheet named after the template (+ opts.nameSuffix для уникальности на операцию).
  const targetSheet = createUniqueSheet_(ss, template.title + (opts.nameSuffix || ''));
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

  // Картинки вставляем ТОЛЬКО если они есть в метаданных шаблона. Пустой imagesJson
  // (после миграции в Drive) = картинок нет → не делаем ни flush, ни скан стора
  // (раньше store-scan гонялся впустую ~2с на каждый шаблон без картинок).
  const parsedImages = parseJsonArray_(template.imagesJson);
  const hasDriveImages = parsedImages.some((img) => img.driveFileId);
  if (parsedImages.length) {
    SpreadsheetApp.flush(); // зафиксировать копию шаблона до вставки картинок
    insertTemplateImages_(targetSheet, targetRow, targetColumn, template.imagesJson);
    if (!hasDriveImages) {
      // Старый шаблон без Drive-кеша — fallback на скан стора.
      runWithSheetVisible_(sourceSheet, () => {
        insertOverGridImages_(
          sourceSheet, template.storeRow, template.storeColumn,
          template.height, template.width,
          targetSheet, targetRow, targetColumn
        );
      });
    }
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
  return withDocumentLock_(function() { return deleteTemplateImpl_(templateId); });
}

function deleteTemplateImpl_(templateId) {
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

  // Очистка слота в _TC_STORE (картинки + содержимое + объединения) — общий путь
  // purgeStoreSlot_ (тот же, что при переносе слота на пересохранении).
  const storeSheet = ss.getSheetByName(TECHMAP_APP.storeSheetName);
  const storeColumn = Math.max(Number(rawRow[5]) || 1, 1);
  const cols = Math.max(Number(rawRow[7]) || 1, 1);
  purgeStoreSlot_(storeSheet, storeRow, storeColumn, height, cols);

  catalogSheet.deleteRow(rowIndex + 2);
  invalidateCatalogCache_();
  bumpCatalogVersion_();

  // If catalog is now empty, purge any orphaned content still in STORE.
  if (storeSheet && readCatalog_().length === 0) {
    purgeEntireStore_(storeSheet);
  }

  return { deleted: true, id: templateId, title };
}
