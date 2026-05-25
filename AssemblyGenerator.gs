// ============================================================
//  Assembly Generator — генератор техкарт межплатных сборок
//  Сканирует активный лист (Google Sheets Tables),
//  создаёт 1–5 листов операций из шаблонов _TC_LIBRARY.
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
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  const sheetData = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues()
    : [];

  const assemblyInfo = scanForTable1_(sheetData);
  const components   = scanForSpyTable_(sheetData);
  const templates    = readCatalog_().map(t => ({ id: t.id, title: t.title, category: t.category || '' }));
  const ops          = readOpRecordsForGenerator_();

  return { assemblyInfo, components, templates, ops, opLabels: ASSEMBLY_GEN.opLabels };
}

// Scan sheet rows for Таблица1: find a header row containing "Индекс" or "Наименование сборки"
function scanForTable1_(data) {
  for (let r = 0; r < data.length - 1; r++) {
    const row = data[r];
    const lower = row.map(c => String(c || '').toLowerCase().trim());
    const hasIndex = lower.some(c => c.includes('индекс'));
    const hasName  = lower.some(c => c.includes('наименование'));
    if (!hasIndex && !hasName) continue;

    const headers  = row.map(c => String(c || '').trim());
    const dataRow  = data[r + 1];
    const result   = {};
    headers.forEach((h, i) => {
      if (h && dataRow[i] !== '' && dataRow[i] != null) result[h] = dataRow[i];
    });
    return result;
  }
  return {};
}

// Scan sheet rows for СПЯ table: find header row with "Тип" + ("ГРМ" or "Артикул" or "Наименование")
function scanForSpyTable_(data) {
  for (let r = 0; r < data.length; r++) {
    const row   = data[r];
    const lower = row.map(c => String(c || '').toLowerCase().trim());

    const typeIdx = lower.findIndex(c => c === 'тип' || c === 'type');
    const nameIdx = lower.findIndex(c => c === 'грм' || c.includes('артикул') || c.includes('наименование') || c.includes('название'));

    if (typeIdx < 0 || nameIdx < 0) continue;

    // Found header row — read until empty
    const qtyIdx    = lower.findIndex(c => c.includes('кол-во') || c.includes('qty'));
    const sideIdx   = lower.findIndex(c => c === 'ст.' || c === 'ст' || c.includes('сторона') || c === 'side');
    const lengthIdx = lower.findIndex(c => c.includes('длина') || c === 'мм' || c === 'mm');

    const components = [];
    for (let dr = r + 1; dr < data.length; dr++) {
      const drow = data[dr];
      if (drow.every(c => c === '' || c == null)) break;
      const type = typeIdx >= 0 ? String(drow[typeIdx] || '').trim() : '';
      const name = nameIdx >= 0 ? String(drow[nameIdx] || '').trim() : '';
      if (!type && !name) continue;
      components.push({
        id:     `spy-${dr}`,
        type,
        name,
        art:    name,
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

    const thisResult = computeOperationResult_(op.type, config, prevResult);
    const phMap      = buildPlaceholderMap_(op, config, prevResult, thisResult);
    replacePlaceholders_(sheet, phMap);

    prevResult = thisResult;
    createdSheets.push(insertResult.sheetName);
  }

  return { ok: true, sheets: createdSheets };
}

function computeOperationResult_(opType, config, prevResult) {
  const wire   = config.wireName || '';
  const lenStr = config.length ? `${config.length}мм` : '';
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
    [p.index]:        config.assemblyIndex || '',
    [p.name]:         config.assemblyName  || '',
    [p.wireName]:     config.wireName      || '',
    [p.wireArt]:      config.wireArt       || '',
    [p.wireQty]:      String(config.wireQty || ''),
    [p.length]:       String(config.length  || ''),
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
