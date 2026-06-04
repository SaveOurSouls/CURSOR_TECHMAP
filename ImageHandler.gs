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

// Уменьшенное превью Drive-файла (blob) или null. insertImage режет blob по лимиту
// ~1 млн символов (≈750КБ бинарника); w800-превью обычно 0.5–0.7МБ → влезает, тогда как
// ~1МБ оригинал — нет. w1000+ часто возвращает оригинал (если нативная ширина ≤ размера).
function thumbBlob_(driveFileId, sz) {
  if (driveFileId == null || driveFileId === '') return null;
  try {
    var r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + driveFileId + '&sz=' + (sz || 'w800'),
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    var b = r.getBlob();
    return (b && b.getBytes().length > 0 && (b.getContentType() || '').indexOf('image') === 0) ? b : null;
  } catch (e) { return null; }
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
      try {
        var inserted = targetSheet.insertImage(blob, col, row, img.xOffset || 0, img.yOffset || 0);
        if (inserted && img.width && img.height) inserted.setWidth(img.width).setHeight(img.height);
      } catch (insErr) {
        // Оригинал превысил лимит insertImage (~1 млн символов) → плавающее w800-превью Drive в ту же позицию.
        var tb = thumbBlob_(img.driveFileId, 'w800');
        if (tb) {
          var ins2 = targetSheet.insertImage(tb, col, row, img.xOffset || 0, img.yOffset || 0);
          if (ins2 && img.width && img.height) ins2.setWidth(img.width).setHeight(img.height);
        }
      }
    } catch (e) { /* битый blob/Drive — пропускаем картинку */ }
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

    var blob = resp.getBlob();
    // Гард: на тяжёлой таблице XLSX огромен → Utilities.unzip падает с "Exceeded
    // memory limit" и роняет весь скрипт. Это крайний фолбэк для картинок —
    // лучше тихо пропустить (вернуть {}), чем крэшнуть генерацию.
    if (blob.getBytes().length > 25 * 1024 * 1024) return {};

    var files = {};
    Utilities.unzip(blob.setContentType('application/zip')).forEach(function(f) { files[f.getName()] = f; });

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
// Сколько over-grid картинок НЕ перенеслось при последнем copySourceImagesToStore_
// (getBlob/getUrl/XLSX не дали blob). Execution-scoped (как _lastPostMerges) — сейвер
// читает после вызова и показывает видимое предупреждение вместо тихой потери.
var _lastImageCaptureSkipped = 0;

function copySourceImagesToStore_(sourceSheet, range, storeSheet, storeRow, storeCol) {
  var height = range.getNumRows();
  var width = range.getNumColumns();
  var startRow = range.getRow();
  var startCol = range.getColumn();
  _lastImageCaptureSkipped = 0;

  clearStoreSlotImages_(storeSheet, storeRow, storeCol, height, width);

  var xlsxImages = null;
  var imagesData = [];
  var folder = null;
  var dbgInRange = 0;

  try {
    var allImgs = sourceSheet.getImages();
    Logger.log('IMG capture START «' + sourceSheet.getName() + '» getImages=' + allImgs.length
      + ' диапазон r' + startRow + '..' + (startRow + height - 1) + ' c' + startCol + '..' + (startCol + width - 1));
    allImgs.forEach(function(img) {
      try {
        var anchor = img.getAnchorCell();
        var row = anchor.getRow();
        var col = anchor.getColumn();
        if (row < startRow - 2 || row > startRow + height + 1) return;
        if (col < startCol - 2 || col > startCol + width + 1) return;
        dbgInRange += 1;
        var blob = getOverGridImageBlob_(img, sourceSheet);
        if (!blob) {
          if (!xlsxImages) xlsxImages = buildXlsxImageMap_(sourceSheet);
          blob = (xlsxImages && xlsxImages[row + '_' + col]) || null;
        }
        // Картинка реально потеряна (getBlob/getUrl/XLSX не дали blob) — логируем
        // позицию, чтобы «пропали картинки при сохранении» был диагностируем по clasp logs.
        if (!blob) { _lastImageCaptureSkipped += 1; Logger.log('IMG capture SKIP @' + row + ',' + col + ' — getBlob/getUrl/XLSX не дали blob'); return; }
        var relRow = row - startRow;
        var relCol = col - startCol;
        // 1) Drive-кеш — хост ссылки, лимита размера нет (основной путь). Делаем ПЕРВЫМ,
        //    чтобы driveFileId был доступен для =IMAGE-фолбэка при показе в store.
        var driveFileId = '';
        try {
          if (!folder) folder = getOrCreateImageCacheFolder_();
          var driveFile = folder.createFile(blob.setName('tc_img_' + relRow + '_' + relCol));
          shareDriveItemForViewers_(driveFile);
          driveFileId = driveFile.getId();
          imagesData.push({
            driveFileId: driveFileId,
            relRow: relRow,
            relCol: relCol,
            xOffset: img.getAnchorCellXOffset(),
            yOffset: img.getAnchorCellYOffset(),
            width: img.getWidth(),
            height: img.getHeight(),
          });
        } catch (e) { Logger.log('IMG Drive-step ERR @' + row + ',' + col + ': ' + (e && e.message)); }
        // 2) Показ В _TC_STORE как плавающей картинки. >2МБ — Google не даёт (лимит insertImage),
        //    пропускаем: картинка уже в Drive-кеше (не теряется), но для ПОКАЗА исходник должен быть <2МБ.
        try {
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
        } catch (e) {
          // Оригинал превысил лимит insertImage → плавающее w800-превью Drive в store.
          var tb = thumbBlob_(driveFileId, 'w800');
          if (tb) {
            try {
              var ins2 = storeSheet.insertImage(tb, Math.max(storeCol, storeCol + relCol), Math.max(storeRow, storeRow + relRow),
                img.getAnchorCellXOffset(), img.getAnchorCellYOffset());
              if (ins2) {
                ins2.setWidth(img.getWidth()).setHeight(img.getHeight());
                try { ins2.setAltTextDescription('tc:' + relRow + ',' + relCol); } catch (e3) {}
              }
              Logger.log('IMG store-insert OK (w800) @' + row + ',' + col);
            } catch (e2) { Logger.log('IMG store-insert thumb ERR @' + row + ',' + col + ': ' + (e2 && e2.message)); }
          } else { Logger.log('IMG store-insert пропущен (thumb не вышел) @' + row + ',' + col + ': ' + (e && e.message)); }
        }
      } catch (e) { Logger.log('IMG capture per-image ERR @' + (typeof row !== 'undefined' ? row : '?') + ',' + (typeof col !== 'undefined' ? col : '?') + ': ' + (e && e.message)); }
    });
  } catch (e) { Logger.log('IMG capture ERR (getImages): ' + (e && e.message)); }
  Logger.log('IMG capture ИТОГ: getImages=' + (typeof allImgs !== 'undefined' ? allImgs.length : '?')
    + ' вДиапазоне=' + dbgInRange + ' захвачено=' + imagesData.length + ' пропущено=' + _lastImageCaptureSkipped);
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
