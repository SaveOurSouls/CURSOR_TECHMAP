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

    // ГРН is always right after the first "#" auto-number column
    const firstHashIdx = lower.findIndex(c => c === '#');
    let nameIdx = (firstHashIdx >= 0 && firstHashIdx + 1 < lower.length) ? firstHashIdx + 1 : -1;
    if (nameIdx < 0) {
      nameIdx = lower.findIndex(c =>
        c.includes('грн') || c.includes('грм') || c.includes('наименование') || c.includes('название')
      );
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
  const wd = wireData || {};
  const sA = config.sideA || {};
  const sB = config.sideB || {};

  switch (opType) {
    case 'cutWire': {
      // Build a readable summary: "Name 150мм; Name2 200мм"
      const names   = String(wd.name   || '').split('\n').filter(Boolean);
      const lengths = String(wd.length || '').split('\n').filter(Boolean);
      const parts   = names.map((n, i) => {
        const l = lengths[i] ? `${lengths[i]}мм` : '';
        return [n, l].filter(Boolean).join(' ');
      });
      return parts.length ? parts.join('; ') : (wd.name || '');
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
function fillTechCardStructurally_(sheet, op, opType, config, prevResult, thisResult, wireData) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const values   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();

  const sA = config.sideA || {};
  const sB = config.sideB || {};
  const wd = wireData || {};

  const comp =
      opType === 'cutWire'  ? { art: wd.art  || wd.name  || '', name: wd.name  || wd.art  || '', norm: String(wd.length  || '') }
    : opType === 'prsTermA' ? { art: sA.termArt || sA.termName || '', name: sA.termName || '', norm: String(sA.termQty || '') }
    : opType === 'insTermA' ? { art: sA.connArt || sA.connName || '', name: sA.connName || '', norm: String(sA.connQty || '') }
    : opType === 'prsTermB' ? { art: sB.termArt || sB.termName || '', name: sB.termName || '', norm: String(sB.termQty || '') }
    : opType === 'insTermB' ? { art: sB.connArt || sB.connName || '', name: sB.connName || '', norm: String(sB.connQty || '') }
    : null;

  // ── Debug: dump all non-empty rows ───────────────────────────
  Logger.log('=== fillStructure: sheet=%s opType=%s comp=%s', sheet.getName(), opType, JSON.stringify(comp));
  for (let r = 0; r < values.length; r++) {
    const cells = [];
    for (let c = 0; c < values[r].length; c++) {
      const v = String(values[r][c] || '').trim();
      if (v) cells.push('c' + (c + 1) + ':' + v.substring(0, 25));
    }
    if (cells.length) Logger.log('  r%s: %s', r + 1, cells.join(' | '));
  }

  // ── Global column indices (scan ALL cells) ────────────────────
  let artCol = -1, grnCol = -1, normCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (artCol  < 0 && (cell === 'артикул'     || cell === 'art'       || cell === 'обозначение')) artCol  = c;
      if (grnCol  < 0 && (cell === 'грн'          || cell === 'наименование' || cell === 'название'
                          || (cell.includes('грн') && cell.length < 6)))                              grnCol  = c;
      if (normCol < 0 && (cell === 'норма'        || cell === 'кол-во'    || cell === 'qty'
                          || (cell.includes('норма') && cell.length < 10)))                           normCol = c;
    }
    if (artCol >= 0 && grnCol >= 0 && normCol >= 0) break;
  }
  Logger.log('  GlobalCols: art=%s grn=%s norm=%s', artCol, grnCol, normCol);

  // ── Section detection (scan ALL cells in each row) ────────────
  let kompRow = -1, sfInRow = -1, resultRow = -1, sfOutRow = -1, timeRow = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (!cell) continue;

      if (timeRow   < 0 && /расч[её]?тное\s*врем/i.test(cell))                { timeRow   = r; break; }
      if (resultRow < 0 && timeRow < 0 && /\bрезультат\b/.test(cell))          { resultRow = r; break; }
      if (kompRow   < 0 && cell.includes('комплектующ'))                        { kompRow   = r; break; }
      if (/полуфабрикат|^п\/ф/.test(cell)) {
        if (resultRow >= 0 || timeRow >= 0) { if (sfOutRow < 0) sfOutRow = r; }
        else                               { if (sfInRow  < 0) sfInRow  = r; }
        break;
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
    if (formulas[r] && formulas[r][col]) return;
    Logger.log('  -> setCell[r%s,c%s]=%s', r + 1, col + 1, v);
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
  if (kompRow >= 0 && comp) {
    const cols = resolveCols(kompRow);
    setCell(kompRow, cols.art,  comp.art);
    setCell(kompRow, cols.name, comp.name);
    setCell(kompRow, cols.norm, comp.norm);
  }

  // ── Fill input Полуфабрикат ───────────────────────────────────
  if (sfInRow >= 0 && prevResult) {
    const cols = resolveCols(sfInRow);
    setCell(sfInRow, cols.name >= 0 ? cols.name : cols.art, prevResult);
  }

  // ── Fill Результат section ────────────────────────────────────
  if (resultRow >= 0 && thisResult) {
    const cols  = resolveCols(resultRow);
    const nameC = cols.name >= 0 ? cols.name : cols.art;
    // Fill the result row itself; if its name-column already has non-formula text, try next row
    const existing = nameC >= 0 ? String(values[resultRow][nameC] || '').trim() : '';
    if (!existing) {
      setCell(resultRow, nameC, thisResult);
    } else {
      // Result header has text in name col — look for the first empty data row right below
      let filled = false;
      for (let r = resultRow + 1; r < Math.min(values.length, resultRow + 4) && !filled; r++) {
        const v = nameC >= 0 ? String(values[r][nameC] || '').trim() : '';
        if (!v && !(formulas[r] && formulas[r][nameC])) {
          setCell(r, nameC, thisResult);
          filled = true;
        }
      }
      if (!filled) setCell(resultRow, nameC, thisResult); // fallback: overwrite
    }
  }

  // ── Fill output Полуфабрикат (after Результат section) ───────
  if (sfOutRow >= 0 && thisResult) {
    const cols = resolveCols(sfOutRow);
    setCell(sfOutRow, cols.name >= 0 ? cols.name : cols.art, thisResult);
  }

  // ── Fill time section ────────────────────────────────────────
  // Collect ALL non-empty data rows after the time header, fill in order:
  // 1 row → tOp; 2 rows → tPrep,tOp; 3+ rows → tPrep,tOp,tMachine
  if (timeRow >= 0) {
    const timeDataRows = [];
    for (let r = timeRow + 1; r < values.length; r++) {
      if (values[r].some(c => String(c || '').trim())) timeDataRows.push(r);
      else if (timeDataRows.length > 0) break;
    }

    const tNormCol = normCol >= 0 ? normCol : (artCol >= 0 ? artCol + 1 : 1);
    const tVals    = timeDataRows.length <= 1 ? [op.tOp   || '']
                   : timeDataRows.length === 2 ? [op.tPrep || '', op.tOp || '']
                   : [op.tPrep || '', op.tOp || '', op.tMachine || ''];

    for (let i = 0; i < timeDataRows.length && i < tVals.length; i++) {
      setCell(timeDataRows[i], tNormCol, tVals[i]);
    }
  }
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
