// ============================================================
//  AssemblyGenerator.gs — генератор техкарт межплатных сборок
//  Зависимости: Config.gs, Utils.gs, TemplateStore.gs, OperationDatabase.gs
// ============================================================

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
  const terRecords   = readTerRecordsForGenerator_();

  return { assemblyInfo, components, templates, ops, terRecords };
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
      if (h.length > 2 && !/[\s\-]/.test(h)) return;
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

    let nameIdx = lower.findIndex(c =>
      c === 'грн' || c.includes('грн') || c === 'наименование' || c === 'название'
    );
    if (nameIdx < 0) {
      const firstHashIdx = lower.findIndex(c => c === '#');
      nameIdx = (firstHashIdx >= 0 && firstHashIdx + 1 < lower.length) ? firstHashIdx + 1 : -1;
    }

    if (typeIdx < 0 || nameIdx < 0) continue;

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

function readTerRecordsForGenerator_() {
  try {
    let snapshot = getTechOperationsSnapshot_();
    // Auto-resync if schema version changed (e.g. new L+/L- extraction was added)
    if (String(snapshot.meta && snapshot.meta.schemaVersion) !== String(TECHOPS_DB_APP.schemaVersion)) {
      syncTechOperationsDatabase();
      snapshot = getTechOperationsSnapshot_();
    }
    return (snapshot.records || [])
      .filter(r => r.tabKey === 'ter' && r.terArticle)
      .map(r => ({
        article: r.terArticle || '',
        lPlus:   r.terLPlus   || '',
        lMinus:  r.terLMinus  || '',
      }));
  } catch (e) { return []; }
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

// config.wires  = [{name, art, qty, length}, ...]
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

  try {
    for (const op of config.ops) {
      if (!op.templateId) continue;

      const wireData = (op.type === 'cutWire' && Array.isArray(config.wires))
        ? buildCombinedWireData_(config.wires)
        : null;

      const insertResult = insertTemplate(op.templateId);
      const sheet = ss.getSheetByName(insertResult.sheetName);
      if (!sheet) throw new Error(`Лист "${insertResult.sheetName}" не найден.`);

      const thisResult = computeOperationResult_(op.type, config, prevResult, wireData);

      const phMap = buildPlaceholderMap_(op, config, prevResult, thisResult, wireData);
      replacePlaceholders_(sheet, phMap);

      fillTechCardStructurally_(sheet, op, op.type, config, prevResult, thisResult, wireData);

      prevResult = thisResult;
      createdSheets.push(insertResult.sheetName);
    }

    return { ok: true, sheets: createdSheets };
  } catch (e) {
    // Откат: удаляем все созданные листы при ошибке
    createdSheets.forEach(name => {
      try {
        const s = ss.getSheetByName(name);
        if (s) ss.deleteSheet(s);
      } catch (_) {}
    });
    throw e;
  }
}

// Combines multiple wire entries into a single data object for the CUT_WIRE tech card.
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
    [p.tolerance]:    buildCutTolerance_(wireData),
    [p.lengthKd]:     buildCutLengthKd_(wireData),
  };
}

// Returns "(+/-)Xмм" — linear tolerance scaled from toleranceMmPerM at 1000mm.
// Rounds to nearest 0.5mm; minimum 0.5mm.
function buildCutTolerance_(wireData) {
  const wd  = wireData || {};
  const len = parseFloat(String(wd.length || '').split('\n')[0].replace(',', '.')) || 0;
  if (!len) return '';
  const raw = len / 1000 * (ASSEMBLY_GEN.toleranceMmPerM || 8);
  const tol = Math.max(1, Math.ceil(raw));
  const tolStr = String(tol);
  return `(+/-)${tolStr}мм`;
}

// Returns "Xмм" — the actual cut length for the [L КД] placeholder.
function buildCutLengthKd_(wireData) {
  const wd  = wireData || {};
  const len = parseFloat(String(wd.length || '').split('\n')[0].replace(',', '.')) || 0;
  return len ? String(len).replace('.', ',') + 'мм' : '';
}

function replacePlaceholders_(sheet, phMap) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const range    = sheet.getRange(1, 1, lastRow, lastCol);
  const values   = range.getValues();
  const formulas = range.getFormulas();

  const dirtyRows = {};
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) continue;
      const val = values[r][c];
      if (typeof val !== 'string' || val === '') continue;
      let nv = val;
      for (const [token, repl] of Object.entries(phMap)) {
        if (nv.includes(token)) nv = nv.split(token).join(String(repl));
      }
      if (nv !== val) {
        if (!dirtyRows[r]) dirtyRows[r] = [];
        dirtyRows[r].push({ c, val: nv });
      }
    }
  }

  for (const [rStr, changes] of Object.entries(dirtyRows)) {
    const r = Number(rStr);
    changes.sort((a, b) => a.c - b.c);
    let i = 0;
    while (i < changes.length) {
      const start = changes[i].c;
      const spanVals = [changes[i].val];
      let j = i + 1;
      while (j < changes.length && changes[j].c === changes[j - 1].c + 1) {
        spanVals.push(changes[j].val);
        j++;
      }
      try {
        sheet.getRange(r + 1, start + 1, 1, spanVals.length).setValues([spanVals]);
      } catch (e) {
        for (let k = 0; k < spanVals.length; k++) {
          try { sheet.getRange(r + 1, start + k + 1).setValue(spanVals[k]); } catch (e2) {}
        }
      }
      i = j;
    }
  }
}

// ── Structural fill ───────────────────────────────────────────

function fillTechCardStructurally_(sheet, op, opType, config, prevResult, thisResult, wireData) {
  let lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  let values   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  let mergeMap = buildMergeMap_(sheet);

  const sA = config.sideA || {};
  const sB = config.sideB || {};
  const wd = wireData || {};

  const comp =
      opType === 'prsTermA' ? { art: sA.termArt || sA.termName || '', name: sA.termName || '', norm: String(sA.termQty || '') }
    : opType === 'insTermA' ? { art: sA.connArt || sA.connName || '', name: sA.connName || '', norm: String(sA.connQty || '') }
    : opType === 'prsTermB' ? { art: sB.termArt || sB.termName || '', name: sB.termName || '', norm: String(sB.termQty || '') }
    : opType === 'insTermB' ? { art: sB.connArt || sB.connName || '', name: sB.connName || '', norm: String(sB.connQty || '') }
    : null;

  const isCutWire = opType === 'cutWire';
  const wires = isCutWire && Array.isArray(config.wires) && config.wires.length > 0
      ? config.wires : null;

  // ── Global column indices ─────────────────────────────────────
  let artCol = -1, grnCol = -1, normCol = -1, seqCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (artCol  < 0 && (cell === 'артикул'  || cell === 'art' || cell === 'обозначение'))         artCol  = c;
      if (grnCol  < 0 && (cell === 'грн'       || cell === 'наименование' || cell === 'название'
                          || (cell.includes('грн') && cell.length < 6)))                             grnCol  = c;
      if (normCol < 0 && (cell === 'норма'     || cell === 'кол-во' || cell === 'qty'
                          || (cell.includes('норма') && cell.length < 10)))                          normCol = c;
      if (seqCol  < 0 && cell === '№')                                                              seqCol  = c;
    }
    if (artCol >= 0 && grnCol >= 0 && normCol >= 0) break;
  }

  // ── Section detection ─────────────────────────────────────────
  function detectSections(v) {
    let kp = -1, sfI = -1, res = -1, sfO = -1, tm = -1;
    for (let r = 0; r < v.length; r++) {
      for (let c = 0; c < v[r].length; c++) {
        const cell = String(v[r][c] || '').toLowerCase().trim();
        if (!cell) continue;
        if (tm  < 0 && /расс?ч[её]?тное\s*врем/i.test(cell))       { tm  = r; break; }
        if (res < 0 && tm < 0 && /результат/.test(cell))            { res = r; break; }
        if (kp  < 0 && cell.includes('комплектующ'))                { kp  = r; break; }
        if (/полуфабрикат|^п\/ф/.test(cell)) {
          if (res >= 0 || tm >= 0) { if (sfO < 0) { sfO = r; break; } }
          else                     { if (sfI < 0) { sfI = r; break; } }
        }
      }
    }
    return { kompRow: kp, sfInRow: sfI, resultRow: res, sfOutRow: sfO, timeRow: tm };
  }

  let { kompRow, sfInRow, resultRow, sfOutRow, timeRow } = detectSections(values);

  // ── Helpers ──────────────────────────────────────────────────
  function setCell(r, col, val) {
    if (col < 0 || col >= lastCol) return;
    const v = val == null ? '' : String(val);
    if (!v) return;
    const fml = formulas[r] && formulas[r][col];
    if (fml && !/^=""$|^=''$/.test(fml.trim())) return;
    let wr = r + 1, wc = col + 1;
    if (mergeMap[wr] && mergeMap[wr][wc]) { const p = mergeMap[wr][wc]; wr = p.r; wc = p.c; }
    try { sheet.getRange(wr, wc).setValue(v); } catch (e) {}
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
      const bound = Math.min(
        ...[sfInRow, resultRow, timeRow].filter(x => x > kompRow).concat([values.length])
      );
      const slots = [kompRow];
      for (let r = kompRow + 1; r < bound; r++) {
        const lbl = String(values[r][0] || '').trim();
        if (lbl && !/^\d+$/.test(lbl)) break;
        const artV  = cols.art  >= 0 ? String(values[r][cols.art]  || '').trim() : '';
        const grnV  = cols.name >= 0 ? String(values[r][cols.name] || '').trim() : '';
        if (!artV && !grnV) slots.push(r);
        else break;
      }

      if (wires.length > slots.length) {
        const insertCount = wires.length - slots.length;
        const lastSlot    = slots[slots.length - 1];
        const ok = insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, kompRow + 1);
        if (ok) {
          for (let i = 0; i < insertCount; i++) slots.push(lastSlot + 1 + i);
          lastRow  = sheet.getLastRow();
          values   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
          formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
          mergeMap = buildMergeMap_(sheet);
          const sec = detectSections(values);
          sfInRow   = sec.sfInRow;
          resultRow = sec.resultRow;
          sfOutRow  = sec.sfOutRow;
          timeRow   = sec.timeRow;
        }
      }

      const partQty = (config.partQty > 0) ? config.partQty : 1;
      for (let i = 0; i < Math.min(wires.length, slots.length); i++) {
        const w = wires[i];
        const r = slots[i];
        let normVal = '';
        if (w.qty > 0 && w.length > 0) {
          const n = w.qty * w.length * partQty / 1000;
          if (isFinite(n)) normVal = String(n).replace('.', ',');
        } else if (w.length > 0) {
          const n = w.length * partQty / 1000;
          if (isFinite(n)) normVal = String(n).replace('.', ',');
        }
        if (seqCol >= 0) {
          let sr = r + 1, sc = seqCol + 1;
          if (mergeMap[sr] && mergeMap[sr][sc]) { const p = mergeMap[sr][sc]; sr = p.r; sc = p.c; }
          try { sheet.getRange(sr, sc).setValue(i + 1); } catch (e) {}
        }
        setCell(r, cols.art,  w.art  || w.name || '');
        setCell(r, cols.name, w.name || '');
        setCell(r, cols.norm, normVal);
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

  // ── Fill Результат → Полуфабрикат ────────────────────────────
  if (thisResult && (sfOutRow >= 0 || resultRow >= 0)) {
    let fRes = -1, fSfOut = -1;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cell = String(values[r][c] || '').toLowerCase().trim();
        if (!cell) continue;
        if (fRes < 0 && /результат/.test(cell))                         { fRes   = r; break; }
        if (fRes >= 0 && fSfOut < 0 && /полуфабрикат|^п\/ф/.test(cell)) { fSfOut = r; break; }
      }
    }

    if (fSfOut >= 0) {
      const hdr   = findColHeadersAbove_(values, fSfOut);
      const nameC = hdr.name >= 0 ? hdr.name : (hdr.art >= 0 ? hdr.art : grnCol);
      const normC = hdr.norm >= 0 ? hdr.norm : normCol;

      if (isCutWire && wires && wires.length > 0 && nameC >= 0) {
        const partQty = (config.partQty > 0) ? config.partQty : 1;

        let resTimeBound = values.length;
        for (let r = fSfOut + 1; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            if (/расс?ч[её]?тное\s*врем/i.test(String(values[r][c] || ''))) {
              resTimeBound = r; break;
            }
          }
          if (resTimeBound < values.length) break;
        }

        const resSlots = [fSfOut];
        for (let r = fSfOut + 1; r < resTimeBound; r++) {
          const firstCell = String(values[r][0] || '').toLowerCase().trim();
          if (firstCell && !/полуфабрикат|^п\/ф/.test(firstCell) && !/^\d+$/.test(firstCell)) break;
          const nameV = String(values[r][nameC] || '').trim();
          if (!nameV) resSlots.push(r);
          else break;
        }

        if (wires.length > resSlots.length) {
          const insertCount = wires.length - resSlots.length;
          const lastSlot    = resSlots[resSlots.length - 1];
          const ok = insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, fSfOut + 1);
          if (ok) {
            for (let i = 0; i < insertCount; i++) resSlots.push(lastSlot + 1 + i);
            values   = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
            formulas = sheet.getRange(1, 1, values.length, lastCol).getFormulas();
            mergeMap = buildMergeMap_(sheet);
            const sec2 = detectSections(values);
            timeRow    = sec2.timeRow;
          }
        }

        for (let i = 0; i < Math.min(wires.length, resSlots.length); i++) {
          const w      = wires[i];
          const rowNum = resSlots[i] + 1;
          const wName  = [w.art || w.name, w.length ? w.length + 'мм' : ''].filter(Boolean).join(' ');
          const wNorm  = String(w.qty * partQty).replace('.', ',');
          if (seqCol >= 0) {
            let sr = rowNum, sc = seqCol + 1;
            if (mergeMap[sr] && mergeMap[sr][sc]) { const p = mergeMap[sr][sc]; sr = p.r; sc = p.c; }
            try { sheet.getRange(sr, sc).setValue(i + 1); } catch (e) {}
          }
          fillMergedCell_(sheet, rowNum, nameC + 1, wName, mergeMap);
          if (normC >= 0) fillMergedCell_(sheet, rowNum, normC + 1, wNorm, mergeMap);
        }
      } else if (nameC >= 0) {
        fillMergedCell_(sheet, fSfOut + 1, nameC + 1, thisResult, mergeMap);
      }
    }
  }

  // ── Fill Расчетное время ──────────────────────────────────────
  if (timeRow >= 0) {
    const timeDataRows = [];
    for (let r = timeRow + 1; r < values.length; r++) {
      const rowCells = values[r].map(c => String(c || '').toLowerCase().trim());
      if (!rowCells.some(v => v)) { if (timeDataRows.length > 0) break; continue; }
      if (rowCells.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение' || v === 'факт')) continue;
      timeDataRows.push(r);
    }

    if (timeDataRows.length > 0) {
      const cols     = resolveCols(timeDataRows[0]);
      const tNormCol = cols.norm >= 0 ? cols.norm : normCol;
      const tNameCol = cols.name >= 0 ? cols.name : grnCol;

      if (isCutWire && wires && wires.length > 0 && tNameCol >= 0) {
        const partQty  = (config.partQty > 0) ? config.partQty : 1;
        const tOpSec   = parseFloat(String(op.tOp   || '').replace(',', '.')) || 0;
        const tPrepSec = parseFloat(String(op.tPrep || '').replace(',', '.')) || 0;
        const tOpMin   = tOpSec   / 60;
        const tPrepMin = tPrepSec / 60;

        const timeSlots = [...timeDataRows];
        for (let r = timeDataRows[timeDataRows.length - 1] + 1; r < values.length; r++) {
          const rowCells = values[r].map(c => String(c || '').toLowerCase().trim());
          if (!rowCells.some(v => v)) break;
          if (rowCells.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение' || v === 'факт')) continue;
          const nameV = tNameCol >= 0 ? String(values[r][tNameCol] || '').trim() : '';
          if (!nameV) timeSlots.push(r);
          else break;
        }

        if (wires.length > timeSlots.length) {
          const insertCount = wires.length - timeSlots.length;
          const lastSlot    = timeSlots[timeSlots.length - 1];
          const ok = insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, timeDataRows[0] + 1);
          if (ok) {
            for (let i = 0; i < insertCount; i++) timeSlots.push(lastSlot + 1 + i);
            values   = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
            formulas = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getFormulas();
            mergeMap = buildMergeMap_(sheet);
          }
        }

        for (let i = 0; i < Math.min(wires.length, timeSlots.length); i++) {
          const w      = wires[i];
          const rowNum = timeSlots[i] + 1;
          const wName  = [w.art || w.name, w.length ? w.length + 'мм' : ''].filter(Boolean).join(' ');
          const rawNorm = (tOpMin > 0 ? tOpMin * w.qty * partQty : w.qty * partQty) + tPrepMin;
          const wNorm = isFinite(rawNorm)
            ? String(rawNorm % 1 === 0 ? rawNorm : rawNorm.toFixed(2)).replace('.', ',')
            : '';
          if (seqCol >= 0) {
            let sr = rowNum, sc = seqCol + 1;
            if (mergeMap[sr] && mergeMap[sr][sc]) { const p = mergeMap[sr][sc]; sr = p.r; sc = p.c; }
            try { sheet.getRange(sr, sc).setValue(i + 1); } catch (e) {}
          }
          fillMergedCell_(sheet, rowNum, tNameCol + 1, wName, mergeMap);
          if (tNormCol >= 0) fillMergedCell_(sheet, rowNum, tNormCol + 1, wNorm, mergeMap);
        }
      } else {
        const secToMin = v => {
          const s = parseFloat(String(v || '').replace(',', '.'));
          if (!s) return '';
          const m = s / 60;
          return String(m % 1 === 0 ? m : m.toFixed(2)).replace('.', ',');
        };
        const tVals = timeDataRows.length <= 1 ? [secToMin(op.tOp)]
                    : timeDataRows.length === 2 ? [secToMin(op.tPrep), secToMin(op.tOp)]
                    : [secToMin(op.tPrep), secToMin(op.tOp), secToMin(op.tMachine)];

        for (let i = 0; i < timeDataRows.length && i < tVals.length; i++) {
          const r = timeDataRows[i];
          setCell(r, tNormCol, tVals[i]);
          if (i === 0 && thisResult) fillMergedCell_(sheet, r + 1, tNameCol + 1, thisResult, mergeMap);
        }
      }
    }
  }

  // ── Fill Допуск column ────────────────────────────────────────
  if (isCutWire && wires && wires.length > 0) {
    let dopuskCol = -1;
    outer: for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (/^допуск$/i.test(String(values[r][c] || '').trim())) { dopuskCol = c; break outer; }
      }
    }
    if (dopuskCol >= 0) {
      const tolStr = buildCutTolerance_(wd);
      const lenStr = buildCutLengthKd_(wd);
      for (let r = 0; r < values.length; r++) {
        const rowText = values[r].map(c => String(c || '').toLowerCase()).join(' ');
        const cur     = String(values[r][dopuskCol] || '').trim().toLowerCase();
        if (/тестовый\s*рез/.test(rowText) && (!cur || cur === '-' || cur === '—')) {
          fillMergedCell_(sheet, r + 1, dopuskCol + 1, tolStr, mergeMap);
        }
        if (lenStr && (cur === '[l кд]' || /измерить\s*длину|длина\s*должна/.test(rowText) && (!cur || cur === '[l кд]'))) {
          fillMergedCell_(sheet, r + 1, dopuskCol + 1, lenStr, mergeMap);
        }
      }
    }
  }
}

// ── Row insertion ─────────────────────────────────────────────

// Inserts `count` rows after `afterRow` (1-based), handling merged cells gracefully.
// Unmerges everything, inserts, copies srcRow template, then restores merges.
function insertRowsAfterSafe_(sheet, afterRow, count, srcRow) {
  const lastCol = sheet.getLastColumn();

  const fullRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  const merges = fullRange.getMergedRanges().map(r => ({
    r1: r.getRow(), r2: r.getLastRow(),
    c1: r.getColumn(), c2: r.getLastColumn()
  }));

  merges.forEach(m => {
    try {
      sheet.getRange(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1).breakApart();
    } catch (e) {}
  });

  try {
    sheet.insertRowsAfter(afterRow, count);
    const adjustedSrc = srcRow > afterRow ? srcRow + count : srcRow;
    const srcRange = sheet.getRange(adjustedSrc, 1, 1, lastCol);
    for (let i = 1; i <= count; i++) {
      srcRange.copyTo(sheet.getRange(afterRow + i, 1, 1, lastCol));
    }
  } catch (e) {
    merges.forEach(m => {
      try { sheet.getRange(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1).merge(); } catch (e2) {}
    });
    return false;
  }

  merges.forEach(m => {
    let r1 = m.r1, r2 = m.r2;
    if (r1 > afterRow)      { r1 += count; r2 += count; }
    else if (r2 > afterRow) { r2 += count; }
    try { sheet.getRange(r1, m.c1, r2 - r1 + 1, m.c2 - m.c1 + 1).merge(); } catch (e) {}
  });

  // Применяем горизонтальные объединения исходной строки к каждой новой
  const srcRowMerges = merges.filter(m => m.r1 === srcRow && m.r2 === srcRow && m.c2 > m.c1);
  for (let i = 1; i <= count; i++) {
    srcRowMerges.forEach(m => {
      try { sheet.getRange(afterRow + i, m.c1, 1, m.c2 - m.c1 + 1).merge(); } catch (e) {}
    });
  }

  return true;
}

// ── Merge utilities ───────────────────────────────────────────

// Returns mergeMap[row1based][col1based] = {r, c} pointing to top-left of merged range.
function buildMergeMap_(sheet) {
  const map = {};
  try {
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr < 1 || lc < 1) return map;
    const merges = sheet.getRange(1, 1, lr, lc).getMergedRanges();
    for (const m of merges) {
      const r1 = m.getRow(), c1 = m.getColumn();
      const r2 = m.getLastRow(), c2 = m.getLastColumn();
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (r !== r1 || c !== c1) {
            if (!map[r]) map[r] = {};
            map[r][c] = { r: r1, c: c1 };
          }
        }
      }
    }
  } catch (e) {}
  return map;
}

// Writes value to a merged cell — resolves to the top-left cell using pre-computed mergeMap.
function fillMergedCell_(sheet, row1, col1, value, mergeMap) {
  const v = value == null ? '' : String(value);
  if (!v) return;
  try {
    let wr = row1, wc = col1;
    if (mergeMap && mergeMap[wr] && mergeMap[wr][wc]) {
      const p = mergeMap[wr][wc];
      wr = p.r;
      wc = p.c;
    }
    sheet.getRange(wr, wc).setValue(v);
  } catch (e) {}
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
