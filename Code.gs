const TECHMAP_APP = {
  menuTitle: 'Техкарты',
  librarySheetName: '_TC_LIBRARY',
  templatePrefix: '_TPL_',
  templateRangeA1: 'A1:L32',
  catalogHeaders: [
    'id',
    'title',
    'category',
    'description',
    'sheetName',
    'rangeA1',
    'imageConfigJson',
  ],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(TECHMAP_APP.menuTitle)
    .addItem('Открыть библиотеку шаблонов', 'showTemplateSidebar')
    .addSeparator()
    .addItem('Установить демо-библиотеку', 'initializeDemoLibrary')
    .addItem('Показать служебные листы', 'showLibrarySheets')
    .addItem('Скрыть служебные листы', 'hideLibrarySheets')
    .addToUi();
}

function showTemplateSidebar() {
  ensureDemoLibraryInstalled_();

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Библиотека техкарт')
    .setWidth(360);

  SpreadsheetApp.getUi().showSidebar(html);
}

function initializeDemoLibrary() {
  ensureDemoLibraryInstalled_(true);
  SpreadsheetApp.getUi().alert(
    'Демо-библиотека установлена.',
    'Созданы 5 шаблонов операций. Откройте "Техкарты -> Открыть библиотеку шаблонов", выберите шаблон и вставьте его в активную ячейку.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getTemplateCatalog() {
  ensureDemoLibraryInstalled_();
  return readCatalog_();
}

function insertTemplate(templateId) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  if (isSystemSheet_(sheet.getName())) {
    throw new Error('Вставка шаблонов на служебные листы запрещена. Перейдите на рабочий лист.');
  }

  const activeRange = sheet.getActiveRange();
  if (!activeRange) {
    throw new Error('Не выбрана ячейка для вставки.');
  }

  const template = readCatalog_().find((item) => item.id === templateId);
  if (!template) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }

  const sourceSheet = ss.getSheetByName(template.sheetName);
  if (!sourceSheet) {
    throw new Error(`Лист шаблона "${template.sheetName}" не найден.`);
  }

  const sourceRange = sourceSheet.getRange(template.rangeA1 || TECHMAP_APP.templateRangeA1);
  const targetRow = activeRange.getRow();
  const targetColumn = activeRange.getColumn();
  const targetRows = sourceRange.getNumRows();
  const targetColumns = sourceRange.getNumColumns();

  ensureSheetCapacity_(sheet, targetRow + targetRows - 1, targetColumn + targetColumns - 1);
  copyTemplateDimensions_(sourceSheet, sheet, sourceRange, targetRow, targetColumn);

  const targetRange = sheet.getRange(targetRow, targetColumn, targetRows, targetColumns);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  clearNotesInTargetHeader_(targetRange);

  SpreadsheetApp.flush();

  return {
    title: template.title,
    sheetName: sheet.getName(),
    insertedRange: targetRange.getA1Notation(),
  };
}

function openTemplateSheet(templateId) {
  if (!templateId) {
    throw new Error('Не передан идентификатор шаблона.');
  }

  ensureDemoLibraryInstalled_();
  const ss = SpreadsheetApp.getActive();
  const template = readCatalog_().find((item) => item.id === templateId);
  if (!template) {
    throw new Error(`Шаблон "${templateId}" не найден.`);
  }

  const sheet = ss.getSheetByName(template.sheetName);
  if (!sheet) {
    throw new Error(`Лист шаблона "${template.sheetName}" не найден.`);
  }

  sheet.showSheet();
  ss.setActiveSheet(sheet);
  ss.setActiveRange(sheet.getRange('A1'));
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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureDemoLibraryInstalled_(forceRebuild) {
  const ss = SpreadsheetApp.getActive();
  const catalogSheet = ensureCatalogSheet_(ss);
  const catalog = readCatalog_();
  const existingIds = new Set(catalog.map((item) => item.id));

  TECHMAP_DEMO_TEMPLATES.forEach((templateSpec) => {
    if (forceRebuild || !existingIds.has(templateSpec.id)) {
      createOrUpdateTemplate_(ss, catalogSheet, templateSpec);
    }
  });

  if (forceRebuild) {
    hideLibrarySheets();
  }
}

function ensureCatalogSheet_(ss) {
  let sheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TECHMAP_APP.librarySheetName);
    sheet.getRange(1, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues([TECHMAP_APP.catalogHeaders]);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#d9e2f3');
    sheet.hideSheet();
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, TECHMAP_APP.catalogHeaders.length)
    .getValues()[0];
  if (currentHeaders.join('|') !== TECHMAP_APP.catalogHeaders.join('|')) {
    sheet.getRange(1, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues([TECHMAP_APP.catalogHeaders]);
  }

  return sheet;
}

function readCatalog_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureCatalogSheet_(ss);
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
      sheetName: row[4],
      rangeA1: row[5],
      images: parseImageConfig_(row[6]),
    }));
}

function createOrUpdateTemplate_(ss, catalogSheet, templateSpec) {
  const sheetName = `${TECHMAP_APP.templatePrefix}${templateSpec.id}`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  renderTemplateSheet_(sheet, templateSpec);
  sheet.hideSheet();
  upsertCatalogRecord_(catalogSheet, {
    id: templateSpec.id,
    title: templateSpec.title,
    category: templateSpec.category,
    description: templateSpec.description,
    sheetName,
    rangeA1: TECHMAP_APP.templateRangeA1,
    imageConfigJson: JSON.stringify(templateSpec.images || []),
  });
}

function renderTemplateSheet_(sheet, templateSpec) {
  resetSheet_(sheet, 32, 12);
  applyBaseGrid_(sheet, templateSpec);
  drawHeader_(sheet, templateSpec);
  drawWarningBlock_(sheet);
  drawLeftSections_(sheet, templateSpec);
  drawImageZone_(sheet, templateSpec);
  drawFooterTable_(sheet, templateSpec);
  sheet.getRange('A1').setNote(`techmap-template:${templateSpec.id}`);
}

function resetSheet_(sheet, targetRows, targetColumns) {
  sheet.clear();
  try {
    sheet.getImages().forEach((image) => image.remove());
  } catch (error) {
    // getImages/remove may be unavailable in some editors; the template still works without cleanup.
  }

  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  if (maxRows < targetRows) {
    sheet.insertRowsAfter(maxRows, targetRows - maxRows);
  }
  if (maxColumns < targetColumns) {
    sheet.insertColumnsAfter(maxColumns, targetColumns - maxColumns);
  }

  if (maxRows > targetRows) {
    sheet.deleteRows(targetRows + 1, maxRows - targetRows);
  }
  if (maxColumns > targetColumns) {
    sheet.deleteColumns(targetColumns + 1, maxColumns - targetColumns);
  }

  sheet.getDataRange().breakApart();
  sheet.setHiddenGridlines(false);
}

function applyBaseGrid_(sheet, templateSpec) {
  const columnWidths = [155, 95, 95, 95, 110, 110, 110, 110, 110, 110, 110, 110];
  columnWidths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  for (let row = 1; row <= 32; row += 1) {
    const height = row <= 2 ? 28 : row >= 26 ? 26 : 24;
    sheet.setRowHeight(row, height);
  }

  const fullRange = sheet.getRange('A1:L32');
  fullRange
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#ffffff');

  sheet.getRange('A1:L32').setBorder(true, true, true, true, true, true, '#444444', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('E3:L24').setBorder(true, true, true, true, true, true, '#666666', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A26:F32').setBorder(true, true, true, true, true, true, '#666666', SpreadsheetApp.BorderStyle.SOLID);

  if (templateSpec.rowHeights) {
    templateSpec.rowHeights.forEach((item) => {
      sheet.setRowHeight(item.row, item.height);
    });
  }
}

function drawHeader_(sheet, templateSpec) {
  const blue = '#cfe2f3';
  const darkGreen = '#183b2b';

  sheet.getRange('A1:B2').merge();
  sheet.getRange('A1:B2')
    .setValue('Fraxis')
    .setBackground(darkGreen)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center');

  sheet.getRange('C1:D1').merge().setValue('Название проекта').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E1:F1').merge().setValue(templateSpec.projectCode || '630K.1').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('G1:J1').merge().setValue(templateSpec.operationHeader).setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('K1:L1').merge().setValue('Лист 1/1').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('C2:F2').merge().setValue('Наименование сборки по чертежу').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('G2:J2').merge().setValue(templateSpec.assemblyName).setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('K2:L2').merge().setValue('Рабочее место').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
}

function drawWarningBlock_(sheet) {
  sheet.getRange('E3:L4').merge();
  sheet
    .getRange('E3:L4')
    .setValue(
      'Важно! Механические манипуляции по настройке оборудования проводить в обесточенном состоянии. В иных случаях - исключительно ручным инструментом.'
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function drawLeftSections_(sheet, templateSpec) {
  let row = 3;
  row = drawLabeledListSection_(sheet, row, 'Инструмент', templateSpec.tools);
  row = drawLabeledListSection_(sheet, row, 'Оборудование', templateSpec.equipment);
  row = drawLabeledListSection_(sheet, row, 'Материалы', templateSpec.materials);
  row = drawInstructionSection_(sheet, row, templateSpec.steps);
  row = drawResultSection_(sheet, row, templateSpec.results, templateSpec.timings);

  if (row < 25) {
    sheet.getRange(row, 1, 25 - row, 4).setBorder(true, true, true, true, true, true, '#666666', SpreadsheetApp.BorderStyle.DOTTED);
  }
}

function drawLabeledListSection_(sheet, startRow, title, items) {
  const blue = '#d9e2f3';
  sheet.getRange(startRow, 1, 1, 3).merge().setValue(title).setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(startRow, 4).setValue('Кол-во').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');

  let row = startRow + 1;
  (items || []).forEach((item) => {
    sheet.getRange(row, 1, 1, 3).merge().setValue(item.name);
    sheet.getRange(row, 4).setValue(item.qty || '');
    row += 1;
  });

  return row;
}

function drawInstructionSection_(sheet, startRow, steps) {
  const blue = '#d9e2f3';
  sheet.getRange(startRow, 1, 1, 4).merge().setValue('Инструкция к выполнению').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  let row = startRow + 1;

  (steps || []).forEach((step, index) => {
    sheet.getRange(row, 1).setValue(index + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 3).merge().setValue(step);
    row += 1;
  });

  return row;
}

function drawResultSection_(sheet, startRow, results, timings) {
  const blue = '#d9e2f3';
  sheet.getRange(startRow, 1, 1, 3).merge().setValue('Результат').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(startRow, 4).setValue('Кол-во').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');

  let row = startRow + 1;
  (results || []).forEach((item, index) => {
    sheet.getRange(row, 1).setValue(index + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 2).merge().setValue(item.name);
    sheet.getRange(row, 4).setValue(item.qty || '');
    row += 1;
  });

  sheet.getRange(row, 1, 1, 3).merge().setValue('Расчетное время').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(row, 4).setValue('Мин').setBackground(blue).setFontWeight('bold').setHorizontalAlignment('center');
  row += 1;

  (timings || []).forEach((item, index) => {
    sheet.getRange(row, 1).setValue(index + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 2).merge().setValue(item.name);
    sheet.getRange(row, 4).setValue(item.minutes || '');
    row += 1;
  });

  return row;
}

function drawImageZone_(sheet, templateSpec) {
  const placeholders = templateSpec.imagePlaceholders || [];
  const imagesByAnchor = {};
  (templateSpec.images || []).forEach((image) => {
    imagesByAnchor[`${image.row}:${image.column}`] = image;
  });

  placeholders.forEach((placeholder) => {
    const range = sheet.getRange(placeholder.range);
    range.merge();
    range
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontColor('#666666')
      .setBorder(true, true, true, true, true, true, '#9aa0a6', SpreadsheetApp.BorderStyle.DASHED)
      .setBackground('#f7f7f7');

    const anchor = sheet.getRange(range.getRow(), range.getColumn());
    const image = imagesByAnchor[`${range.getRow()}:${range.getColumn()}`];
    if (image && image.url) {
      anchor.setFormula(buildImageFormula_(image.url, image.width || 220, image.height || 120));
      if (placeholder.caption) {
        anchor.setNote(placeholder.caption);
      }
    } else {
      anchor.setValue(placeholder.caption || 'Изображение');
    }
  });

  (templateSpec.sideLabels || []).forEach((label) => {
    sheet
      .getRange(label.range)
      .merge()
      .setValue(label.text)
      .setFontWeight('bold')
      .setFontSize(18)
      .setHorizontalAlignment('center')
      .setFontColor(label.color);
  });
}

function drawFooterTable_(sheet, templateSpec) {
  const header = templateSpec.footer.header || [];
  const rows = templateSpec.footer.rows || [];
  const headerRange = sheet.getRange(27, 1, 1, header.length);
  headerRange.setValues([header]).setBackground('#d9d9d9').setFontWeight('bold').setHorizontalAlignment('center');

  if (rows.length) {
    sheet.getRange(28, 1, rows.length, header.length).setValues(rows);
  }

  if (header.length >= 6 && rows.length) {
    sheet.getRange(28, header.length - 1, rows.length, 1).setBackground('#d9ead3');
    sheet.getRange(28, header.length, rows.length, 1).setBackground('#e2f0d9');
  }
}

function upsertCatalogRecord_(catalogSheet, record) {
  const lastRow = catalogSheet.getLastRow();
  const ids = lastRow > 1 ? catalogSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const existingIndex = ids.indexOf(record.id);
  const rowValues = [[
    record.id,
    record.title,
    record.category,
    record.description,
    record.sheetName,
    record.rangeA1,
    record.imageConfigJson,
  ]];

  if (existingIndex >= 0) {
    catalogSheet.getRange(existingIndex + 2, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues(rowValues);
  } else {
    catalogSheet.getRange(lastRow + 1, 1, 1, TECHMAP_APP.catalogHeaders.length).setValues(rowValues);
  }
}

function ensureSheetCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function copyTemplateDimensions_(sourceSheet, targetSheet, sourceRange, targetRow, targetColumn) {
  for (let i = 0; i < sourceRange.getNumRows(); i += 1) {
    targetSheet.setRowHeight(targetRow + i, sourceSheet.getRowHeight(sourceRange.getRow() + i));
  }

  for (let i = 0; i < sourceRange.getNumColumns(); i += 1) {
    targetSheet.setColumnWidth(targetColumn + i, sourceSheet.getColumnWidth(sourceRange.getColumn() + i));
  }
}

function clearNotesInTargetHeader_(targetRange) {
  const note = targetRange.getCell(1, 1).getNote();
  if (note && note.indexOf('techmap-template:') === 0) {
    targetRange.getCell(1, 1).clearNote();
  }
}

function parseImageConfig_(value) {
  if (!value) {
    return [];
  }

  try {
    return JSON.parse(value);
  } catch (error) {
    return [];
  }
}

function buildImageFormula_(url, width, height) {
  const safeUrl = String(url || '').replace(/"/g, '""');
  return `=IMAGE("${safeUrl}",4,${height},${width})`;
}

function isSystemSheet_(sheetName) {
  return (
    sheetName === TECHMAP_APP.librarySheetName ||
    sheetName.indexOf(TECHMAP_APP.templatePrefix) === 0
  );
}
