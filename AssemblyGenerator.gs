// ============================================================
//  Assembly Generator — генератор техкарт межплатных сборок
//  Читает Таблица1 + СПЯ из активного листа, показывает
//  конфигуратор, создаёт 1–5 листов операций из шаблонов _TC_LIBRARY.
// ============================================================

const ASSEMBLY_GEN = {
  placeholders: {
    index:       '{{INDEX}}',
    name:        '{{NAME}}',
    wireName:    '{{WIRE_NAME}}',
    wireArt:     '{{WIRE_ART}}',
    wireQty:     '{{WIRE_QTY}}',
    length:      '{{LENGTH}}',
    semifinished:'{{SEMIFINISHED}}',
    result:      '{{RESULT}}',
    termNameA:   '{{TERM_A_NAME}}',
    termArtA:    '{{TERM_A_ART}}',
    termQtyA:    '{{TERM_A_QTY}}',
    connNameA:   '{{CONN_A_NAME}}',
    connArtA:    '{{CONN_A_ART}}',
    connQtyA:    '{{CONN_A_QTY}}',
    termNameB:   '{{TERM_B_NAME}}',
    termArtB:    '{{TERM_B_ART}}',
    termQtyB:    '{{TERM_B_QTY}}',
    connNameB:   '{{CONN_B_NAME}}',
    connArtB:    '{{CONN_B_ART}}',
    connQtyB:    '{{CONN_B_QTY}}',
    opNum:       '{{OP_NUM}}',
    tPrep:       '{{T_PREP}}',
    tOp:         '{{T_OP}}',
    tMachine:    '{{T_MACHINE}}',
  },

  // Operation sequence identifiers
  opTypes: ['cutWire', 'prsTermA', 'insTermA', 'prsTermB', 'insTermB'],

  opLabels: {
    cutWire:  'Резка',
    prsTermA: 'Опрессовка терминалов (А)',
    insTermA: 'Монтаж терминалов (А)',
    prsTermB: 'Опрессовка терминалов (В)',
    insTermB: 'Монтаж терминалов (В)',
  },
};

// ── Entry point ──────────────────────────────────────────────

function showAssemblyGeneratorDialog() {
  const data = getAssemblyGeneratorData_();
  const tmpl = HtmlService.createTemplateFromFile('AssemblyGeneratorDialog');
  tmpl.initialData = JSON.stringify(data);
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(660).setHeight(800),
    'Генератор техкарт — Межплатная сборка'
  );
}

// ── Data loading ─────────────────────────────────────────────

function getAssemblyGeneratorData_() {
  const ss = SpreadsheetApp.getActive();
  const assemblyInfo = readNamedRangeAsMap_(ss, 'Таблица1');
  const components   = readSpyTable_(ss, 'СПЯ');
  const templates    = readCatalog_().map(t => ({
    id:       t.id,
    title:    t.title,
    category: t.category || '',
  }));
  const ops = readOpRecordsForGenerator_();
  return { assemblyInfo, components, templates, ops, opTypes: ASSEMBLY_GEN.opTypes, opLabels: ASSEMBLY_GEN.opLabels };
}

function readNamedRangeAsMap_(ss, rangeName) {
  try {
    const range = ss.getRangeByName(rangeName);
    if (!range) return {};
    const values = range.getValues();
    const result = {};
    values.forEach(row => {
      const k = String(row[0] || '').trim();
      if (k) result[k] = row.length > 1 ? row[1] : '';
    });
    return result;
  } catch (e) { return {}; }
}

function readSpyTable_(ss, rangeName) {
  try {
    const range = ss.getRangeByName(rangeName);
    if (!range) return [];
    const vals = range.getValues();
    if (vals.length < 2) return [];

    const hdrs = vals[0].map(h => String(h || '').trim().toLowerCase());
    const idxOf = (aliases) => {
      for (const a of aliases) {
        const i = hdrs.findIndex(h => h.includes(a));
        if (i >= 0) return i;
      }
      return -1;
    };

    const ti = idxOf(['тип', 'type', 'вид']);
    const ni = idxOf(['наименование', 'название', 'name']);
    const ai = idxOf(['артикул', 'обозначение']);
    const qi = idxOf(['кол-во', 'количество', 'qty']);
    const si = idxOf(['сторона', 'side', 'конец']);
    const li = idxOf(['длина', 'length']);

    return vals.slice(1)
      .filter(r => r.some(c => c !== '' && c != null))
      .map((r, idx) => ({
        id:     `spy-${idx}`,
        type:   ti >= 0 ? String(r[ti]   || '').trim() : '',
        name:   ni >= 0 ? String(r[ni]   || '').trim() : '',
        art:    ai >= 0 ? String(r[ai]   || '').trim() : '',
        qty:    qi >= 0 ? (Number(r[qi]) || 1)         : 1,
        side:   si >= 0 ? String(r[si]   || '').trim().toUpperCase() : '',
        length: li >= 0 ? (Number(r[li]) || 0)         : 0,
      }));
  } catch (e) { return []; }
}

function readOpRecordsForGenerator_() {
  try {
    const snapshot = getTechOperationsSnapshot_();
    return (snapshot.records || [])
      .filter(r => r.tabKey === 'op')
      .map(r => ({
        opNumber: r.opNumber  || '',
        opName:   r.opName    || '',
        label:    r.displayText || '',
        tOp:      r.tOp      || '',
        tPrep:    r.tPrep    || '',
        tMachine: r.tMachine || '',
      }));
  } catch (e) { return []; }
}

// ── Generator ─────────────────────────────────────────────────

function generateAssemblyTechCards(config) {
  if (!config || !Array.isArray(config.ops) || !config.ops.length) {
    throw new Error('Нет операций для создания.');
  }

  const ss = SpreadsheetApp.getActive();
  const createdSheets = [];
  let prevResult = '';

  for (const op of config.ops) {
    if (!op.templateId) continue;

    const insertResult = insertTemplate(op.templateId);
    const sheet = ss.getSheetByName(insertResult.sheetName);
    if (!sheet) throw new Error(`Лист "${insertResult.sheetName}" не найден после вставки шаблона.`);

    const thisResult  = computeOperationResult_(op.type, config, prevResult);
    const phMap       = buildPlaceholderMap_(op, config, prevResult, thisResult);
    replacePlaceholders_(sheet, phMap);

    prevResult = thisResult;
    createdSheets.push(insertResult.sheetName);
  }

  return { ok: true, sheets: createdSheets };
}

function computeOperationResult_(opType, config, prevResult) {
  const wire    = config.wireName  || '';
  const lenStr  = config.length ? `${config.length}мм` : '';
  const wireStr = [wire, lenStr].filter(Boolean).join(' ');
  const sA = config.sideA || {};
  const sB = config.sideB || {};

  switch (opType) {
    case 'cutWire':  return wireStr;
    case 'prsTermA': return [wireStr, sA.termName].filter(Boolean).join(' + ');
    case 'insTermA': return [prevResult, sA.connName].filter(Boolean).join(' → ') + ' ст.А';
    case 'prsTermB': return [wireStr, sB.termName].filter(Boolean).join(' + ');
    case 'insTermB': return [prevResult, sB.connName].filter(Boolean).join(' → ') + ' ст.В';
    default:         return prevResult;
  }
}

function buildPlaceholderMap_(op, config, prevResult, thisResult) {
  const p  = ASSEMBLY_GEN.placeholders;
  const sA = config.sideA || {};
  const sB = config.sideB || {};
  return {
    [p.index]:       config.assemblyIndex || '',
    [p.name]:        config.assemblyName  || '',
    [p.wireName]:    config.wireName      || '',
    [p.wireArt]:     config.wireArt       || '',
    [p.wireQty]:     String(config.wireQty || ''),
    [p.length]:      String(config.length  || ''),
    [p.semifinished]:prevResult,
    [p.result]:      thisResult,
    [p.termNameA]:   sA.termName  || '',
    [p.termArtA]:    sA.termArt   || '',
    [p.termQtyA]:    String(sA.termQty || ''),
    [p.connNameA]:   sA.connName  || '',
    [p.connArtA]:    sA.connArt   || '',
    [p.connQtyA]:    String(sA.connQty || ''),
    [p.termNameB]:   sB.termName  || '',
    [p.termArtB]:    sB.termArt   || '',
    [p.termQtyB]:    String(sB.termQty || ''),
    [p.connNameB]:   sB.connName  || '',
    [p.connArtB]:    sB.connArt   || '',
    [p.connQtyB]:    String(sB.connQty || ''),
    [p.opNum]:       op.opNum     || '',
    [p.tPrep]:       op.tPrep     || '',
    [p.tOp]:         op.tOp       || '',
    [p.tMachine]:    op.tMachine  || '',
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
