// ============================================================
//  RangeCopy.gs — механика копирования диапазонов
//  Зависимости: Utils.gs
// ============================================================

/**
 * Копирует диапазон без корректировки ссылок формул.
 * Приоритет: Sheets API v4 (сохраняет in-cell изображения) →
 *            copyTo(PASTE_NORMAL) → ручная копия значений+формата.
 * После любого пути формулы накладываются поячеечно чтобы предотвратить
 * смещение относительных ссылок.
 */
function copyRangePreservingFormulas_(sourceRange, targetRange) {
  const srcSheet   = sourceRange.getSheet();
  const dstSheet   = targetRange.getSheet();
  const hideSrc    = srcSheet.isSheetHidden();
  const hideDst    = dstSheet.isSheetHidden();
  if (hideSrc) srcSheet.showSheet();
  if (hideDst) dstSheet.showSheet();

  const ss          = SpreadsheetApp.getActive();
  const priorActive = ss.getActiveSheet();

  try {
    targetRange.breakApart();

    let formulas;
    try { formulas = sourceRange.getFormulas(); } catch (e) { formulas = null; }

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

    if (formulas) applySourceFormulasCellwise_(targetRange, formulas);

    // in-cell изображения PASTE_NORMAL не переносит — явно копируем через API
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
 * Копирует диапазон через Sheets API v4 batchUpdate/copyPaste.
 * Требует подключённого advanced-сервиса Google Sheets API.
 * Возвращает true при успехе, false если сервис недоступен или вызов упал.
 */
function tryCopyRangeViaSheetsApi_(sourceRange, targetRange) {
  try {
    if (typeof Sheets === 'undefined') return false;
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        copyPaste: {
          source: {
            sheetId:          sourceRange.getSheet().getSheetId(),
            startRowIndex:    sourceRange.getRow() - 1,
            endRowIndex:      sourceRange.getRow() + sourceRange.getNumRows() - 1,
            startColumnIndex: sourceRange.getColumn() - 1,
            endColumnIndex:   sourceRange.getColumn() + sourceRange.getNumColumns() - 1,
          },
          destination: {
            sheetId:          targetRange.getSheet().getSheetId(),
            startRowIndex:    targetRange.getRow() - 1,
            endRowIndex:      targetRange.getRow() + targetRange.getNumRows() - 1,
            startColumnIndex: targetRange.getColumn() - 1,
            endColumnIndex:   targetRange.getColumn() + targetRange.getNumColumns() - 1,
          },
          pasteType:        'PASTE_NORMAL',
          pasteOrientation: 'NORMAL',
        },
      }],
    }, SpreadsheetApp.getActive().getId());
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * PASTE_FORMAT после записи значений.
 * copyTo между листами нестабилен если активен «чужой» лист — делаем 3 попытки.
 */
function copyRangeFormatPreservingMerges_(sourceRange, targetRange, ss, priorActive, srcSheet, dstSheet) {
  const restorePrior = () => {
    if (
      priorActive &&
      !isSystemSheet_(priorActive.getName()) &&
      ss.getActiveSheet().getSheetId() !== priorActive.getSheetId()
    ) {
      try { ss.setActiveSheet(priorActive); } catch (e) {}
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

/**
 * Накладывает формулы из массива formulas на targetRange.
 * Группирует смежные формулы одной строки в один setFormulas-вызов;
 * пустые ячейки между группами не трогаются (нет риска затереть значения).
 */
function applySourceFormulasCellwise_(targetRange, formulas) {
  if (!formulas || !formulas.length) return;
  const sheet    = targetRange.getSheet();
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();
  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r];
    if (!row) continue;
    let c = 0;
    while (c < row.length) {
      if (!row[c]) { c++; continue; }
      let end = c + 1;
      while (end < row.length && row[end]) end++;
      try {
        sheet.getRange(startRow + r, startCol + c, 1, end - c).setFormulas([row.slice(c, end)]);
      } catch (e) {
        for (let i = c; i < end; i++) {
          try { sheet.getRange(startRow + r, startCol + i).setFormula(row[i]); } catch (e2) {}
        }
      }
      c = end;
    }
  }
}

/**
 * setValues на большой сетке иногда даёт «Ошибка службы: Таблицы» —
 * режем на полосы по 25 строк, затем построчно при повторной ошибке.
 */
function setRangeValuesChunked_(targetRange, values) {
  if (!values || !values.length) return;
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
  const sheet    = targetRange.getSheet();
  for (let r = 0; r < numRows; r += maxChunk) {
    const h     = Math.min(maxChunk, numRows - r);
    const slice = values.slice(r, r + h);
    try {
      sheet.getRange(startRow + r, startCol, h, numCols).setValues(slice);
      SpreadsheetApp.flush();
    } catch (e) {
      setRangeValuesRowByRow_(sheet.getRange(startRow + r, startCol, h, numCols), slice);
    }
  }
}

/** Резервная запись значений построчно/поячеечно при сбое setValues на блоке. */
function setRangeValuesRowByRow_(targetRange, values) {
  const sheet    = targetRange.getSheet();
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();
  const numCols  = values[0] ? values[0].length : 0;
  for (let r = 0; r < values.length; r++) {
    try {
      sheet.getRange(startRow + r, startCol, 1, numCols).setValues([values[r]]);
    } catch (e) {
      for (let c = 0; c < numCols; c++) {
        try { sheet.getRange(startRow + r, startCol + c).setValue(values[r][c]); } catch (e2) {}
      }
    }
    if (r % 20 === 0) {
      try { SpreadsheetApp.flush(); } catch (fe) {}
      Utilities.sleep(20);
    }
  }
}
