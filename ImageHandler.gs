// ============================================================
//  ImageHandler.gs — работа с изображениями (over-grid и in-cell)
//  Зависимости: Utils.gs
// ============================================================

// ── Over-grid image blob ─────────────────────────────────────

/**
 * Возвращает blob over-grid-изображения.
 * Сначала пробует getBlob() (старый GAS API), затем getUrl() + UrlFetchApp.
 * Возвращает null при неудаче обоих путей.
 */
function getOverGridImageBlob_(img, sourceSheet) {
  try {
    if (typeof img.getBlob === 'function') {
      var blob = img.getBlob();
      if (blob) return blob;
    }
  } catch (e) {}

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

// ── In-cell images (Sheets API) ──────────────────────────────

/**
 * Явно копирует in-cell image values через Sheets API.
 * PASTE_NORMAL не переносит in-cell изображения — читаем userEnteredValue.imageValue
 * из каждой ячейки источника и записываем в destination через updateCells.
 * Требует подключённого advanced-сервиса Google Sheets API.
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

// ── Template image capture ───────────────────────────────────

/**
 * Собирает over-grid изображения из диапазона range на листе sourceSheet.
 * Возвращает массив объектов с base64-данными.
 */
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
        // Тот же fallback (getBlob → getUrl → XLSX), что и в прод-пути copySourceImagesToStore_:
        // прямой getBlob() падает на типах изображений GAS 2024+.
        var blob = getOverGridImageBlob_(img, sourceSheet);
        if (!blob) return;
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

// ── Миграция картинок шаблонов в Drive-кеш ───────────────────

/**
 * Однократная миграция: переносит over-grid картинки ВСЕХ шаблонов из стора в
 * Drive-кеш и прописывает driveFileId в каталог. После этого вставка шаблонов
 * идёт по быстрому пути insertTemplateImages_ (чтение из Drive), а не по медленному
 * store-scan (скан всех картинок стора + XLSX). Картинки в сторе НЕ трогаются —
 * только читаются; обратимо (вернуть imagesJson в '[]' → снова store-scan).
 * Идемпотентна: шаблоны, уже имеющие driveFileId, пропускает.
 */
function migrateTemplateImagesToDrive() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var storeSheet   = ss.getSheetByName(TECHMAP_APP.storeSheetName);
  var catalogSheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!storeSheet || !catalogSheet) { ui.alert('Стор или каталог не найден.'); return; }

  var catalog = readCatalog_();
  var xlsxMap = null; // ленивый XLSX-fallback (кешируется на прогон)
  var folder = null;
  var migrated = 0, skipped = 0, imgCount = 0;

  // getImages() на полностью скрытом листе возвращает пусто — временно показываем стор.
  runWithSheetVisible_(storeSheet, function() {
  SpreadsheetApp.flush(); // зафиксировать показ листа до getImages
  var allImages;
  try { allImages = storeSheet.getImages(); } catch (e) { allImages = []; }

  catalog.forEach(function(t) {
    var existing = parseJsonArray_(t.imagesJson);
    if (existing.some(function(im) { return im.driveFileId; })) { skipped++; return; }
    var sr = t.storeRow, sc = t.storeColumn, h = t.height, w = t.width;
    if (!(sr > 0 && h > 0 && w > 0)) { skipped++; return; }

    var imagesData = [];
    allImages.forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow(), col = anchor.getColumn();
        var relRow = row - sr, relCol = col - sc;
        var hasTc = false;
        try {
          var alt = img.getAltTextDescription();
          if (alt && alt.indexOf('tc:') === 0) {
            hasTc = true;
            var parts = alt.slice(3).split(',');
            if (parts.length === 2) {
              var pr = parseInt(parts[0], 10), pc = parseInt(parts[1], 10);
              if (!isNaN(pr)) relRow = pr;
              if (!isNaN(pc)) relCol = pc;
            }
          }
        } catch (e) {}
        if (hasTc) {
          if (row < sr || row > sr + h + 1 || col < sc || col > sc + w + 1) return;
        } else {
          if (row < sr || row > sr + h - 1 || col < sc || col > sc + w - 1) return;
        }
        var blob = getOverGridImageBlob_(img, storeSheet);
        if (!blob) {
          if (!xlsxMap) xlsxMap = buildXlsxImageMap_(storeSheet);
          blob = (xlsxMap && xlsxMap[row + '_' + col]) || null;
        }
        if (!blob) return;
        if (!folder) folder = getOrCreateImageCacheFolder_();
        var driveFile = folder.createFile(blob.setName('tc_img_' + relRow + '_' + relCol));
        shareDriveItemForViewers_(driveFile);
        imagesData.push({
          driveFileId: driveFile.getId(),
          relRow: relRow, relCol: relCol,
          xOffset: img.getAnchorCellXOffset(), yOffset: img.getAnchorCellYOffset(),
          width: img.getWidth(), height: img.getHeight(),
        });
        imgCount++;
      } catch (e) {}
    });

    if (imagesData.length) {
      updateCatalogImagesJson_(catalogSheet, t.id, JSON.stringify(imagesData));
      migrated++;
    }
  });
  }); // runWithSheetVisible_

  invalidateCatalogCache_();
  bumpCatalogVersion_();
  ui.alert('Миграция картинок в Drive',
    'Шаблонов обновлено: ' + migrated +
    '\nКартинок перенесено: ' + imgCount +
    '\n\nГенерация теперь быстрее. Запускать повторно не нужно — новые шаблоны ' +
    'кешируются в Drive автоматически при сохранении.',
    ui.ButtonSet.OK);
}

/** Записывает imagesJson для шаблона по id в каталог (колонка imagesJson = 14). */
function updateCatalogImagesJson_(catalogSheet, id, json) {
  var lastRow = catalogSheet.getLastRow();
  if (lastRow < 2) return;
  var ids = catalogSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      catalogSheet.getRange(i + 2, 14).setValue(json);
      return;
    }
  }
}

// ── Over-grid image insertion ────────────────────────────────

/**
 * Копирует over-grid изображения из слота _TC_STORE на целевой лист.
 * Помеченные tc: изображения используют расширенную зону (±2 строки/столбца).
 * Непомеченные — только строгие границы слота.
 */
function insertOverGridImages_(sourceSheet, sourceRow, sourceCol, height, width, targetSheet, targetRow, targetCol) {
  var xlsxImagesInsert = null;
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

/**
 * Вставляет over-grid изображения из imagesJson на целевой лист.
 * Предпочитает driveFileId (быстрое чтение из Drive); fallback на base64
 * для старых шаблонов, сохранённых до добавления кеша Drive.
 */
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

/** Удаляет файлы изображений из Drive (кеш _TC_IMAGE_CACHE). */
function deleteTemplateImages_(imagesJson) {
  var images = parseJsonArray_(imagesJson);
  images.forEach(function(img) {
    if (img.driveFileId) {
      try { DriveApp.getFileById(img.driveFileId).setTrashed(true); } catch (e) {}
    }
  });
}

// ── Drive image cache ────────────────────────────────────────

/** Возвращает папку кэша изображений из Drive, создаёт при необходимости. */
function getOrCreateImageCacheFolder_() {
  var props = PropertiesService.getUserProperties();
  var folderId = props.getProperty('tc_image_cache_folder');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  var folder = DriveApp.createFolder(TECHMAP_APP.imageCacheFolderName);
  shareDriveItemForViewers_(folder);
  props.setProperty('tc_image_cache_folder', folder.getId());
  return folder;
}

// Делает Drive-объект доступным любому по ссылке (просмотр). Без этого вставка
// картинок через DriveApp.getFileById() падает у ДРУГИХ пользователей таблицы —
// файлы кеша принадлежат создателю и приватны. Диаграммы обжима не секретны.
function shareDriveItemForViewers_(item) {
  try { item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
}

// ── XLSX image extraction (last-resort fallback) ─────────────

// Кеш XLSX-карт изображений на время одного выполнения скрипта (глобалы GAS
// сбрасываются между вызовами). buildXlsxImageMap_ скачивает и распаковывает
// всю таблицу — без кеша это повторялось бы на КАЖДОЙ операции генерации.
var _xlsxImageMapCache_ = {};

/**
 * Кеширующая обёртка над buildXlsxImageMapUncached_: одно скачивание таблицы
 * на лист за выполнение. Кешируются и пустые результаты — чтобы не повторять
 * дорогой неудачный download N раз за один прогон генерации.
 */
function buildXlsxImageMap_(sheet) {
  var sid = sheet.getSheetId();
  if (Object.prototype.hasOwnProperty.call(_xlsxImageMapCache_, sid)) {
    return _xlsxImageMapCache_[sid];
  }
  var map = buildXlsxImageMapUncached_(sheet);
  _xlsxImageMapCache_[sid] = map;
  return map;
}

/**
 * Скачивает таблицу как XLSX (ZIP), разбирает drawing XML и возвращает
 * map { 'row_col': Blob } для каждого изображения на указанном листе.
 * Используется когда getBlob() и getUrl() недоступны (GAS 2024+ тип изображения).
 */
function buildXlsxImageMapUncached_(sheet) {
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
      var posM = content.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
      if (!posM) continue;
      var col1 = parseInt(posM[1], 10) + 1;
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

// ── Store slot image management ──────────────────────────────

/**
 * Копирует over-grid изображения из sourceRange в слот _TC_STORE.
 * Захватывает изображения с отступом ±2 ячейки от границ (для логотипов в margin-колонках).
 * Сохраняет реальное относительное смещение в alt-text для корректного восстановления.
 * Кеширует каждый blob в Drive (_TC_IMAGE_CACHE).
 * Возвращает массив метаданных изображений (driveFileId, relRow, relCol, …).
 */
function copySourceImagesToStore_(sourceSheet, range, storeSheet, storeRow, storeCol) {
  var height = range.getNumRows();
  var width = range.getNumColumns();
  var startRow = range.getRow();
  var startCol = range.getColumn();

  clearStoreSlotImages_(storeSheet, storeRow, storeCol, height, width);

  var xlsxImages = null;
  var imagesData = [];
  var folder = null;

  try {
    sourceSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow();
        var col = anchor.getColumn();
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
        var destRow = Math.max(storeRow, storeRow + relRow);
        var destCol = Math.max(storeCol, storeCol + relCol);
        var inserted = storeSheet.insertImage(
          blob, destCol, destRow,
          img.getAnchorCellXOffset(), img.getAnchorCellYOffset()
        );
        if (inserted) {
          inserted.setWidth(img.getWidth()).setHeight(img.getHeight());
          try { inserted.setAltTextDescription('tc:' + relRow + ',' + relCol); } catch (e) {}
        }
        try {
          if (!folder) folder = getOrCreateImageCacheFolder_();
          var driveFile = folder.createFile(blob.setName('tc_img_' + relRow + '_' + relCol));
          shareDriveItemForViewers_(driveFile);
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

/**
 * Удаляет over-grid изображения слота _TC_STORE.
 * В строгих границах слота — удаляет всегда.
 * В расширенной зоне (±2 колонки) — только tc:-помеченные.
 */
function clearStoreSlotImages_(storeSheet, storeRow, storeCol, height, width) {
  try {
    storeSheet.getImages().forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var r = anchor.getRow();
        var c = anchor.getColumn();
        if (r >= storeRow && r < storeRow + height &&
            c >= storeCol && c < storeCol + width) {
          img.remove();
          return;
        }
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

/**
 * Полностью очищает _TC_STORE: все ячейки и over-grid изображения.
 * Вызывается когда каталог становится пустым.
 */
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

// ── Debug helpers (run from Apps Script editor only) ─────────

/** Диагностика in-cell и over-grid изображений на активном листе. */
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

/** Диагностика извлечения изображений из XLSX для активного листа. */
function debugXlsxImages() {
  Logger.log('=== XLSX Image Extraction Debug ===');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  Logger.log('Sheet: ' + sheet.getName() + ' (id=' + sheet.getSheetId() + ')');

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

  var sheets = ss.getSheets();
  var sheetIdx = 0;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheet.getSheetId()) { sheetIdx = i + 1; break; }
  }
  Logger.log('Sheet index (1-based): ' + sheetIdx);

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

/** Диагностика captureTemplateImages_ на активном листе. */
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
