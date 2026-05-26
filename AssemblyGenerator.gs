// ============================================================
//  Assembly Generator — генератор техкарт межплатных сборок
// ============================================================

const ASSEMBLY_GEN = {
  placeholders: {
    index:        '{{INDEX}}',
    name:         '{{NAME}}',
    wireName:     '{{WIRE_NAME}}',
    wireArt:      '{{WIRE_ART}}',
    wireQty:      '{{WIRE_QTY}}',
    length:       '{{LENGTH}}',
    semifinished: '{{SEMIFINISHED}}',
    result:       '{{RESULT}}',
    termNameA:    '{{TERM_A_NAME}}',
    termArtA:     '{{TERM_A_ART}}',
    termQtyA:     '{{TERM_A_QTY}}',
    connNameA:    '{{CONN_A_NAME}}',
    connArtA:     '{{CONN_A_ART}}',
    connQtyA:     '{{CONN_A_QTY}}',
    termNameB:    '{{TERM_B_NAME}}',
    termArtB:     '{{TERM_B_ART}}',
    termQtyB:     '{{TERM_B_QTY}}',
    connNameB:    '{{CONN_B_NAME}}',
    connArtB:     '{{CONN_B_ART}}',
    connQtyB:     '{{CONN_B_QTY}}',
    opNum:        '{{OP_NUM}}',
    tPrep:        '{{T_PREP}}',
    tOp:          '{{T_OP}}',
    tMachine:     '{{T_MACHINE}}',
  },
};

// ── Entry point ──────────────────────────────────────────────

function showAssemblyGeneratorDialog() {
  const data = getAssemblyGeneratorData_();
  const tmpl = HtmlService.createTemplateFromFile('AssemblyGeneratorDialog');
  tmpl.initialData = JSON.stringify(data);
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(680).setHeight(840),
    'Генератор техкарт — Межплатная сборка'
  );
}

// ── Data loading ─────────────────────────────────────────────

function getAssemblyGeneratorData_() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  const sheetData = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues()
    : [];

  const assemblyInfo = scanForTable1_(sheetData);
  const components   = scanForSpyTable_(sheetData);
  const templates    = readCatalog_().map(t => ({ id: t.id, title: t.title, category: t.category || '' }));
  const ops          = readOpRecordsForGenerator_();

  return { assemblyInfo, components, templates, ops };
}

// Find Таблица1: header row with BOTH "индекс" AND "наименование"
function scanForTable1_(data) {
  for (let r = 0; r < data.length - 1; r++) {
    const row   = data[r];
    const lower = row.map(c => String(c || '').toLowerCase().trim());
    if (!lower.some(c => c.includes('индекс')) || !lower.some(c => c.includes('наименование'))) continue;

    const headers = row.map(c => String(c || '').trim());
    const dataRow = data[r + 1];
    const result  = {};
    headers.forEach((h, i) => {
      if (!h || /^\d+$/.test(h)) return;
      if (h.length > 2 && !h.includes(' ')) return;
      const v = dataRow[i];
      if (v !== '' && v != null) result[h] = v;
    });
    return result;
  }
  return {};
}

// Find СПЯ table: header row with "Тип" + name column (after first "#")
function scanForSpyTable_(data) {
  for (let r = 0; r < data.length; r++) {
    const row   = data[r];
    const lower = row.map(c => String(c || '').toLowerCase().trim());

    const typeIdx = lower.findIndex(c => c === 'тип' || c === 'type');

    // ГРН: search by keyword first (explicit header), then fall back to column after '#'
    let nameIdx = lower.findIndex(c =>
      c === 'грн' || c.includes('грн') || c === 'наименование' || c === 'название'
    );
    if (nameIdx < 0) {
      const firstHashIdx = lower.findIndex(c => c === '#');
      nameIdx = (firstHashIdx >= 0 && firstHashIdx + 1 < lower.length) ? firstHashIdx + 1 : -1;
    }

    if (typeIdx < 0 || nameIdx < 0) continue;

    // Артикул column — separate from ГРН
    const artIdx    = lower.findIndex(c => c.includes('артикул') || c === 'art' || c === 'article');
    const qtyIdx    = lower.findIndex(c => c.includes('кол-во') || c.includes('qty'));
    const sideIdx   = lower.findIndex(c => c === 'ст.' || c === 'ст' || c.includes('сторона') || c === 'side');
    const lengthIdx = lower.findIndex(c => c.includes('длина') || c === 'мм' || c === 'mm');

    const components = [];
    for (let dr = r + 1; dr < data.length; dr++) {
      const drow = data[dr];
      if (drow.every(c => c === '' || c == null)) break;
      const type = String(drow[typeIdx] || '').trim();
      const name = String(drow[nameIdx] || '').trim();
      if (!type && !name) continue;
      // art from Артикул column; fallback to name if column absent
      const art = artIdx >= 0 ? String(drow[artIdx] || '').trim() : name;
      components.push({
        id:     `spy-${dr}`,
        type,
        name,
        art,
        qty:    qtyIdx    >= 0 ? (Number(drow[qtyIdx])    || 1) : 1,
        side:   sideIdx   >= 0 ? String(drow[sideIdx]   || '').trim().toUpperCase() : '',
        length: lengthIdx >= 0 ? (Number(drow[lengthIdx]) || 0) : 0,
      });
    }
    return components;
  }
  return [];
}

function readOpRecordsForGenerator_() {
  try {
    const snapshot = getTechOperationsSnapshot_();
    return (snapshot.records || [])
      .filter(r => r.tabKey === 'op')
      .map(r => ({
        opNumber: r.opNumber || '',
        opName:   r.opName   || '',
        label:    r.displayText || '',
        tOp:      r.tOp      || '',
        tPrep:    r.tPrep    || '',
        tMachine: r.tMachine || '',
      }));
  } catch (e) { return []; }
}

// ── Generator ─────────────────────────────────────────────────

// config.wires  = [{name, art, qty, length}, ...]  — ordered array of active wire entries
// config.ops    = [{type, wireIdx, templateId, opNum, tPrep, tOp, tMachine}, ...]
//   wireIdx = -1 means "use all wires combined" (for the single CUT_WIRE op)
// config.sideA  = {termName, termArt, termQty, connName, connArt, connQty}
// config.sideB  = same shape or null
function generateAssemblyTechCards(config) {
  if (!config || !Array.isArray(config.ops) || !config.ops.length) {
    throw new Error('Нет операций для создания.');
  }

  const ss = SpreadsheetApp.getActive();
  const createdSheets = [];
  let prevResult = '';

  for (const op of config.ops) {
    if (!op.templateId) continue;

    // For cutWire: combine all wires into one placeholder set
    const wireData = (op.type === 'cutWire' && Array.isArray(config.wires))
      ? buildCombinedWireData_(config.wires)
      : null;

    const insertResult = insertTemplate(op.templateId);
    const sheet = ss.getSheetByName(insertResult.sheetName);
    if (!sheet) throw new Error(`Лист "${insertResult.sheetName}" не найден.`);

    const thisResult = computeOperationResult_(op.type, config, prevResult, wireData);

    // Token replacement (for {{INDEX}}, {{NAME}} etc. if present in template)
    const phMap = buildPlaceholderMap_(op, config, prevResult, thisResult, wireData);
    replacePlaceholders_(sheet, phMap);

    // Structural fill: finds Комплектующие / Полуфабрикат / time rows by label,
    // fills them without requiring placeholder tokens in the template.
    fillTechCardStructurally_(sheet, op, op.type, config, prevResult, thisResult, wireData);

    prevResult = thisResult;

    createdSheets.push(insertResult.sheetName);
  }

  return { ok: true, sheets: createdSheets };
}

// Combines multiple wire entries into a single data object for the CUT_WIRE tech card.
// Single wire: returns as-is. Multiple wires: joins fields with newlines.
function buildCombinedWireData_(wires) {
  if (!wires || !wires.length) return {};
  if (wires.length === 1) return wires[0];
  return {
    name:   wires.map(w => w.name   || '').join('\n'),
    art:    wires.map(w => w.art    || '').join('\n'),
    qty:    wires.reduce((s, w) => s + (Number(w.qty) || 1), 0),
    length: wires.map(w => String(w.length || '')).join('\n'),
  };
}

function computeOperationResult_(opType, config, prevResult, wireData) {
  const wd    = wireData || {};
  const sA    = config.sideA || {};
  const sB    = config.sideB || {};
  const wires = Array.isArray(config.wires) ? config.wires : [];

  switch (opType) {
    case 'cutWire': {
      // Format: "art lengthмм; art2 length2мм; ..."
      const src = wires.length > 0 ? wires : (function() {
        const arts    = String(wd.art || wd.name || '').split('\n').filter(Boolean);
        const lengths = String(wd.length || '').split('\n').filter(Boolean);
        return arts.map((a, i) => ({ art: a, length: lengths[i] || '' }));
      })();
      const parts = src.map(w => {
        const a = (w.art || w.name || '').trim();
        const l = w.length ? `${w.length}мм` : '';
        return [a, l].filter(Boolean).join(' ');
      }).filter(Boolean);
      return parts.join('; ') || (wd.art || wd.name || '');
    }
    case 'prsTermA': return [prevResult, sA.termName].filter(Boolean).join(' + ');
    case 'insTermA': return [prevResult, sA.connName].filter(Boolean).join(' → ') + ' ст.А';
    case 'prsTermB': return [prevResult, sB.termName].filter(Boolean).join(' + ');
    case 'insTermB': return [prevResult, sB.connName].filter(Boolean).join(' → ') + ' ст.В';
    default:         return prevResult;
  }
}

function buildPlaceholderMap_(op, config, prevResult, thisResult, wireData) {
  const p  = ASSEMBLY_GEN.placeholders;
  const sA = config.sideA || {};
  const sB = config.sideB || {};
  const wd = wireData || {};
  return {
    [p.index]:        config.assemblyIndex || '',
    [p.name]:         config.assemblyName  || '',
    [p.wireName]:     wd.name   || '',
    [p.wireArt]:      wd.art    || '',
    [p.wireQty]:      String(wd.qty    || ''),
    [p.length]:       String(wd.length || ''),
    [p.semifinished]: prevResult,
    [p.result]:       thisResult,
    [p.termNameA]:    sA.termName  || '',
    [p.termArtA]:     sA.termArt   || '',
    [p.termQtyA]:     String(sA.termQty || ''),
    [p.connNameA]:    sA.connName  || '',
    [p.connArtA]:     sA.connArt   || '',
    [p.connQtyA]:     String(sA.connQty || ''),
    [p.termNameB]:    sB.termName  || '',
    [p.termArtB]:     sB.termArt   || '',
    [p.termQtyB]:     String(sB.termQty || ''),
    [p.connNameB]:    sB.connName  || '',
    [p.connArtB]:     sB.connArt   || '',
    [p.connQtyB]:     String(sB.connQty || ''),
    [p.opNum]:        op.opNum     || '',
    [p.tPrep]:        op.tPrep     || '',
    [p.tOp]:          op.tOp       || '',
    [p.tMachine]:     op.tMachine  || '',
  };
}

function replacePlaceholders_(sheet, phMap) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const range    = sheet.getRange(1, 1, lastRow, lastCol);
  const values   = range.getValues();
  const formulas = range.getFormulas();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) continue;
      const val = values[r][c];
      if (typeof val !== 'string' || val === '') continue;

      let nv = val;
      for (const [token, repl] of Object.entries(phMap)) {
        if (nv.includes(token)) nv = nv.split(token).join(String(repl));
      }
      if (nv !== val) sheet.getRange(r + 1, c + 1).setValue(nv);
    }
  }
}

// ── Structural fill ───────────────────────────────────────────
// Fills Комплектующие / Полуфабрикат / time rows by detecting the card structure.
// Works without placeholder tokens so templates remain usable for manual work.
//
// CUT_WIRE Комплектующие: one row per wire (Артикул, ГРН, Норма=qty×length/1000 м)
// Результат / Расчетное время Наименование: "art lengthмм; ..."  Норма: time value
function fillTechCardStructurally_(sheet, op, opType, config, prevResult, thisResult, wireData) {
  let lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  let values   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();

  const sA = config.sideA || {};
  const sB = config.sideB || {};
  const wd = wireData || {};

  // Component data for non-CUT_WIRE ops (single row fill)
  const comp =
      opType === 'prsTermA' ? { art: sA.termArt || sA.termName || '', name: sA.termName || '', norm: String(sA.termQty || '') }
    : opType === 'insTermA' ? { art: sA.connArt || sA.connName || '', name: sA.connName || '', norm: String(sA.connQty || '') }
    : opType === 'prsTermB' ? { art: sB.termArt || sB.termName || '', name: sB.termName || '', norm: String(sB.termQty || '') }
    : opType === 'insTermB' ? { art: sB.connArt || sB.connName || '', name: sB.connName || '', norm: String(sB.connQty || '') }
    : null;

  // Individual wire list for CUT_WIRE (each wire → own row)
  const isCutWire = opType === 'cutWire';
  const wires = isCutWire && Array.isArray(config.wires) && config.wires.length > 0
      ? config.wires : null;

  // ── Debug: dump all non-empty rows ───────────────────────────
  Logger.log('=== fillStructure: sheet=%s opType=%s comp=%s wires=%s',
    sheet.getName(), opType, JSON.stringify(comp), JSON.stringify(wires));
  for (let r = 0; r < values.length; r++) {
    const cells = [];
    for (let c = 0; c < values[r].length; c++) {
      const v = String(values[r][c] || '').trim();
      if (v) cells.push('c' + (c + 1) + ':' + v.substring(0, 25));
    }
    if (cells.length) Logger.log('  r%s: %s', r + 1, cells.join(' | '));
  }

  // ── Global column indices (scan ALL cells) ────────────────────
  let artCol = -1, grnCol = -1, normCol = -1, seqCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (artCol  < 0 && (cell === 'артикул'  || cell === 'art' || cell === 'обозначение'))          artCol  = c;
      if (grnCol  < 0 && (cell === 'грн'       || cell === 'наименование' || cell === 'название'
                          || (cell.includes('грн') && cell.length < 6)))                              grnCol  = c;
      if (normCol < 0 && (cell === 'норма'     || cell === 'кол-во' || cell === 'qty'
                          || (cell.includes('норма') && cell.length < 10)))                           normCol = c;
      if (seqCol  < 0 && cell === '№')                                                               seqCol  = c;
    }
    if (artCol >= 0 && grnCol >= 0 && normCol >= 0) break;
  }
  Logger.log('  GlobalCols: art=%s grn=%s norm=%s seq=%s', artCol, grnCol, normCol, seqCol);

  // ── Section detection (scan ALL cells in each row) ────────────
  let kompRow = -1, sfInRow = -1, resultRow = -1, sfOutRow = -1, timeRow = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (!cell) continue;
      if (timeRow   < 0 && /расч[её]?тное\s*врем/i.test(cell))               { timeRow   = r; break; }
      if (resultRow < 0 && timeRow < 0 && /\bрезультат\b/.test(cell))          { resultRow = r; break; }
      if (kompRow   < 0 && cell.includes('комплектующ'))                        { kompRow   = r; break; }
      if (/полуфабрикат|^п\/ф/.test(cell)) {
        if (resultRow >= 0 || timeRow >= 0) { if (sfOutRow < 0) { sfOutRow = r; break; } }
        else                               { if (sfInRow  < 0) { sfInRow  = r; break; } }
      }
    }
  }
  Logger.log('  Sections: komp=%s sfIn=%s result=%s sfOut=%s time=%s',
    kompRow, sfInRow, resultRow, sfOutRow, timeRow);

  // ── Helpers ──────────────────────────────────────────────────
  function setCell(r, col, val) {
    if (col < 0 || col >= lastCol) return;
    const v = val == null ? '' : String(val);
    if (!v) return;
    const fml = formulas[r] && formulas[r][col];
    // Skip cells with real calculation formulas.
    // Allow overwriting `=""` / `=''` (blank-string formulas used for conditional formatting).
    if (fml && !/^=""$|^=''$/.test(fml.trim())) return;
    Logger.log('  -> setCell[r%s,c%s]=%s (fml=%s)', r + 1, col + 1, v, fml || '');
    sheet.getRange(r + 1, col + 1).setValue(v);
  }

  function resolveCols(dataRow) {
    const loc = findColHeadersAbove_(values, dataRow);
    return {
      art:  loc.art  >= 0 ? loc.art  : artCol,
      name: loc.name >= 0 ? loc.name : grnCol,
      norm: loc.norm >= 0 ? loc.norm : normCol,
    };
  }

  // ── Fill Комплектующие ────────────────────────────────────────
  if (kompRow >= 0) {
    const cols = resolveCols(kompRow);

    if (wires) {
      // CUT_WIRE: fill one row per wire
      // Find blank component slots between kompRow and the next section boundary
      const bound = Math.min(
        ...[sfInRow, resultRow, timeRow].filter(x => x > kompRow).concat([values.length])
      );
      const slots = [kompRow];
      for (let r = kompRow + 1; r < bound; r++) {
        const lbl = String(values[r][0] || '').trim();
        if (lbl && !/^\d+$/.test(lbl)) break; // non-numeric label = boundary row
        const artV  = cols.art  >= 0 ? String(values[r][cols.art]  || '').trim() : '';
        const grnV  = cols.name >= 0 ? String(values[r][cols.name] || '').trim() : '';
        if (!artV && !grnV) slots.push(r);
        else break;
      }
      Logger.log('  CutWire: %s wires / %s slots (before insert)', wires.length, slots.length);

      // Вставляем дополнительные строки если слотов меньше чем проводов.
      // insertRowsAfterSafe_ снимает объединения, вставляет, копирует, восстанавливает.
      if (wires.length > slots.length) {
        const insertCount = wires.length - slots.length;
        const lastSlot    = slots[slots.length - 1];
        const ok = insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, kompRow + 1);
        if (ok) {
          for (let i = 0; i < insertCount; i++) slots.push(lastSlot + 1 + i);
          const adj = r => (r >= 0 && r > lastSlot) ? r + insertCount : r;
          sfInRow   = adj(sfInRow);
          resultRow = adj(resultRow);
          sfOutRow  = adj(sfOutRow);
          timeRow   = adj(timeRow);
          lastRow   = sheet.getLastRow();
          values    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
          formulas  = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
          Logger.log('  Inserted %s rows → slots %s', insertCount, JSON.stringify(slots.map(r => r + 1)));
        } else {
          Logger.log('  insertRowsAfterSafe_ returned false — filling available slots only');
        }
      }

      for (let i = 0; i < Math.min(wires.length, slots.length); i++) {
        const w = wires[i];
        const r = slots[i];
        const normVal = (w.qty > 0 && w.length > 0) ? w.qty * w.length / 1000 : (w.length || '');
        if (seqCol >= 0) setCell(r, seqCol, i + 1);
        setCell(r, cols.art,  w.art  || w.name || '');
        setCell(r, cols.name, w.name || '');
        setCell(r, cols.norm, normVal !== '' ? normVal : '');
      }
    } else if (comp) {
      setCell(kompRow, cols.art,  comp.art);
      setCell(kompRow, cols.name, comp.name);
      setCell(kompRow, cols.norm, comp.norm);
    }
  }

  // ── Fill input Полуфабрикат ───────────────────────────────────
  if (sfInRow >= 0 && prevResult) {
    const cols = resolveCols(sfInRow);
    setCell(sfInRow, cols.name >= 0 ? cols.name : cols.art, prevResult);
  }

  // ── Fill Результат → Полуфабрикат (Наименование = thisResult) ─
  if (sfOutRow >= 0 && thisResult) {
    const cols  = resolveCols(sfOutRow);
    const nameC = cols.name >= 0 ? cols.name : cols.art;
    setCell(sfOutRow, nameC, thisResult);
  } else if (resultRow >= 0 && thisResult && sfOutRow < 0) {
    // No separate Полуфабрикат row found — fill resultRow itself or first non-header row below it
    const cols  = resolveCols(resultRow);
    const nameC = cols.name >= 0 ? cols.name : cols.art;
    const existing = nameC >= 0 ? String(values[resultRow][nameC] || '').trim() : '';
    if (!existing) {
      setCell(resultRow, nameC, thisResult);
    } else {
      let filled = false;
      for (let r = resultRow + 1; r < Math.min(values.length, resultRow + 6) && !filled; r++) {
        // Skip sub-header rows
        const rowCells = values[r].map(c => String(c || '').toLowerCase().trim());
        if (rowCells.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение')) continue;
        const v = nameC >= 0 ? String(values[r][nameC] || '').trim() : '';
        const fml = formulas[r] && formulas[r][nameC];
        if ((!v || fml === '=""' || fml === "=''") && !(fml && !/^=""$|^=''$/.test(fml.trim()))) {
          setCell(r, nameC, thisResult);
          filled = true;
        }
      }
      if (!filled) setCell(resultRow, nameC, thisResult);
    }
  }

  // ── Fill Расчетное время ─────────────────────────────────────
  // Time data rows after "Расчетное время" header:
  //   Наименование ← thisResult (wire description)
  //   Норма        ← T_PREP / T_OP / T_MACHINE in order
  if (timeRow >= 0) {
    const timeDataRows = [];
    for (let r = timeRow + 1; r < values.length; r++) {
      const rowCells = values[r].map(c => String(c || '').toLowerCase().trim());
      if (!rowCells.some(v => v)) { if (timeDataRows.length > 0) break; continue; }
      // Skip sub-header rows (contain column header keywords like "наименование", "норма")
      if (rowCells.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение' || v === 'факт')) continue;
      timeDataRows.push(r);
    }
    Logger.log('  timeDataRows=%s', JSON.stringify(timeDataRows.map(r => r + 1)));

    if (timeDataRows.length > 0) {
      const cols     = resolveCols(timeDataRows[0]);
      const tNormCol = cols.norm >= 0 ? cols.norm : normCol;
      const tNameCol = cols.name >= 0 ? cols.name : grnCol;

      const tVals = timeDataRows.length <= 1 ? [op.tOp   || '']
                  : timeDataRows.length === 2 ? [op.tPrep || '', op.tOp || '']
                  : [op.tPrep || '', op.tOp || '', op.tMachine || ''];

      for (let i = 0; i < timeDataRows.length && i < tVals.length; i++) {
        const r = timeDataRows[i];
        setCell(r, tNormCol, tVals[i]);
        // First time row gets the wire description in Наименование
        if (i === 0 && thisResult) setCell(r, tNameCol, thisResult);
      }
    }
  }
}

// Inserts `count` rows after `afterRow` (1-based), handling merged cells gracefully.
// Unmerges the whole sheet, inserts rows, copies srcRow template, then restores merges.
// Returns true on success, false if an unrecoverable error occurs.
function insertRowsAfterSafe_(sheet, afterRow, count, srcRow) {
  const lastCol = sheet.getLastColumn();

  // Snapshot all merged ranges before touching anything
  const fullRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  const merges = fullRange.getMergedRanges().map(r => ({
    r1: r.getRow(), r2: r.getLastRow(),
    c1: r.getColumn(), c2: r.getLastColumn()
  }));

  // Break all merges so insert/copy can proceed cleanly
  merges.forEach(m => {
    try {
      sheet.getRange(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1).breakApart();
    } catch (e) { /* already unmerged or out of range */ }
  });

  try {
    sheet.insertRowsAfter(afterRow, count);
    // srcRow is 1-based; after insertion it may have shifted
    const adjustedSrc = srcRow > afterRow ? srcRow + count : srcRow;
    const srcRange = sheet.getRange(adjustedSrc, 1, 1, lastCol);
    for (let i = 1; i <= count; i++) {
      srcRange.copyTo(sheet.getRange(afterRow + i, 1, 1, lastCol));
    }
  } catch (e) {
    Logger.log('  insertRowsAfterSafe_: insert/copy failed (%s)', e.message);
    // Restore merges before returning failure
    merges.forEach(m => {
      try { sheet.getRange(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1).merge(); } catch (e2) {}
    });
    return false;
  }

  // Restore merges with row positions adjusted for the insertion
  merges.forEach(m => {
    let r1 = m.r1, r2 = m.r2;
    if (r1 > afterRow)      { r1 += count; r2 += count; }  // entirely below → shift
    else if (r2 > afterRow) { r2 += count; }                // spans insertion → extend
    // entirely above afterRow → unchanged
    try { sheet.getRange(r1, m.c1, r2 - r1 + 1, m.c2 - m.c1 + 1).merge(); } catch (e) {}
  });

  return true;
}

// Looks backward from dataRow to find a header row with Артикул/ГРН/Норма columns.
function findColHeadersAbove_(values, dataRow) {
  for (let r = dataRow - 1; r >= Math.max(0, dataRow - 20); r--) {
    const lc = values[r].map(c => String(c || '').toLowerCase().trim());
    const artI  = lc.findIndex(c => c === 'артикул'  || c === 'art' || c === 'обозначение');
    const grnI  = lc.findIndex(c => c === 'грн'       || c === 'наименование'
                                     || (c.includes('грн') && c.length < 6));
    const normI = lc.findIndex(c => c === 'норма'     || c === 'кол-во' || c === 'qty');
    if (artI >= 0 && grnI >= 0) return { art: artI, name: grnI, norm: normI >= 0 ? normI : -1 };
    if (grnI >= 0 && normI >= 0) return { art: -1, name: grnI, norm: normI };
  }
  return { art: -1, name: -1, norm: -1 };
}
