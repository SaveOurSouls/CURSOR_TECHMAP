// ============================================================
//  AssemblyGenerator.gs — генератор техкарт межплатных сборок
//  Зависимости: Config.gs, Utils.gs, TemplateStore.gs, OperationDatabase.gs
// ============================================================

// ── Entry point ──────────────────────────────────────────────

function showAssemblyGeneratorDialog() {
  const data = getAssemblyGeneratorData_();
  const tmpl = HtmlService.createTemplateFromFile('AssemblyGeneratorDialog');
  tmpl.initialData = embedJsonForHtml_(data);
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(680).setHeight(840),
    'Генератор техкарт — Межплатная сборка'
  );
}

// ── Data loading ─────────────────────────────────────────────

function getAssemblyGeneratorData_() {
  ensureTechOperationsSnapshotReady_(); // быстро: синк только если кеш пуст/схема сменилась
  // ВАЖНО: при открытии генератора НИКАКИХ записей в таблицу — любая мутация запускает
  // пересчёт всех =GET_NAME на существующих листах → таймаут сайдбара. Открытие = только чтение.
  const ss  = SpreadsheetApp.getActive();
  const src = findAssemblySourceData_(ss); // ищем таблицу на любом листе, не только активном

  const templates    = readCatalog_().map(t => ({ id: t.id, title: t.title, category: t.category || '' }));
  // Снапшот техопераций читаем ОДИН раз и передаём обоим читателям — иначе каждый
  // дёргает getTechOperationsSnapshot_() сам и заново десериализует чанк-кеш (×3 на открытие).
  const snapshot     = getTechOperationsSnapshot_();
  const ops          = readOpRecordsForGenerator_(snapshot);
  const terRecords   = readTerRecordsForGenerator_(snapshot);
  const coaxRecords  = readCoaxRecordsForGenerator_(snapshot);
  const wireDia      = readWireDiaTable_();
  const coaxCableDia = readCoaxCableDiaTable_();

  return { assemblyInfo: src.assemblyInfo, components: src.components,
           sourceSheet: src.sourceSheet, templates, ops, terRecords, coaxRecords, wireDia, coaxCableDia };
}

// Ищет лист с исходной таблицей сборки (СПЯ + изделие). Активный лист в приоритете,
// затем остальные пользовательские (системные пропускаем). Предпочитаем лист, где есть
// И СПЯ, И таблица изделия; иначе — первый с СПЯ. Чинит «данные не подтягиваются»,
// когда активным оказался не тот лист (напр. после удаления созданных листов).
function findAssemblySourceData_(ss) {
  const active = ss.getActiveSheet();
  const ordered = [active].concat(ss.getSheets().filter(s => s.getSheetId() !== active.getSheetId()));
  let fallback = null;
  for (let i = 0; i < ordered.length; i++) {
    const sheet = ordered[i];
    if (isSystemSheet_(sheet.getName())) continue;
    // Сгенерированные листы-техкарты («CODE | Тип») исходную таблицу СПЯ не содержат —
    // пропускаем их в скане (кроме активного, i===0). На документе с десятками созданных
    // карт это убирает десятки полных getValues при открытии генератора (ключевая причина тормозов).
    if (i > 0 && sheet.getName().indexOf(' | ') >= 0) continue;
    if (sheet.getLastRow() < 1) continue;
    const data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues();
    const components = scanForSpyTable_(data);
    if (!components.length) continue;
    const assemblyInfo = scanForTable1_(data);
    if (Object.keys(assemblyInfo).length) {
      return { assemblyInfo, components, sourceSheet: sheet.getName() }; // полное совпадение
    }
    if (!fallback) fallback = { assemblyInfo, components, sourceSheet: sheet.getName() };
  }
  return fallback || { assemblyInfo: {}, components: [], sourceSheet: '' };
}

// Читает справочник Ø изоляции проводов из листа справочника кабелей источника.
// Лист: «СПР.КАБ» (ранее «СПР.КОАКС») — пробуем оба имени.
// Формат: строка заголовков с «AWG» | «ГОСТ» | «Ø жилы» | <марки…>; ниже — данные.
// Возвращает { marks:[{name,norm}], byAwg:{ '30': {'силикон':1.2, …}, … } }.
// Используется генератором для авто-подстановки Ø в формулу запаса на свивку.
function readWireDiaTable_() {
  try {
    const ss = SpreadsheetApp.openById(getSourceSpreadsheetId_());
    const sh = ss.getSheetByName('СПР.КАБ') || ss.getSheetByName('СПР.КОАКС');
    if (!sh || sh.getLastRow() < 2) return null;
    const vals = sh.getDataRange().getValues();
    let hr = -1;
    for (let r = 0; r < vals.length; r++) {
      if (vals[r].some(c => String(c || '').trim().toLowerCase() === 'awg')) { hr = r; break; }
    }
    if (hr < 0) return null;
    const headers = vals[hr].map(c => String(c || '').trim());
    const awgCol  = headers.findIndex(h => h.toLowerCase() === 'awg');
    const gostCol = headers.findIndex(h => h.toLowerCase() === 'гост');
    // Колонки марок — все непустые заголовки правее «Ø жилы» (первые 3 — AWG/ГОСТ/Ø жилы).
    const marks = [];
    headers.forEach((h, c) => {
      if (c > awgCol + 2 && h) marks.push({ name: h, norm: h.toLowerCase().replace(/[\s-]/g, ''), col: c });
    });
    const coreCol = awgCol + 2; // 3-я колонка после AWG — «Ø жилы» (медь), нужен для Z₀
    const byAwg = {};
    const byGost = {};
    for (let r = hr + 1; r < vals.length; r++) {
      const awg = String(vals[r][awgCol] || '').trim();
      if (!awg) continue;
      const row = {};
      const core = parseFloat(String(vals[r][coreCol] || '').replace(',', '.'));
      if (isFinite(core) && core > 0) row.__core = core;
      marks.forEach(m => {
        const v = parseFloat(String(vals[r][m.col] || '').replace(',', '.'));
        if (isFinite(v) && v > 0) row[m.norm] = v;
      });
      byAwg[awg] = row;
      // Индекс по ГОСТ-сечению (нормализуем число: «0,50»→«0.5»).
      const g = gostCol >= 0 ? parseFloat(String(vals[r][gostCol] || '').replace(',', '.')) : NaN;
      if (isFinite(g)) byGost[String(g)] = row;
    }
    return { marks: marks.map(m => ({ name: m.name, norm: m.norm })), byAwg, byGost };
  } catch (e) { return null; }
}

// Слоевые диаметры коакс-кабеля из СПР.КАБ (верхняя таблица: «Кабель | D1 | D2 | D3»).
// Возвращает { byCable: { '<норм.марка>': {d1,d2,d3} } } — для авто-подстановки D1-D3 по проводу.
function readCoaxCableDiaTable_() {
  try {
    const ss = SpreadsheetApp.openById(getSourceSpreadsheetId_());
    const sh = ss.getSheetByName('СПР.КАБ') || ss.getSheetByName('СПР.КОАКС');
    if (!sh || sh.getLastRow() < 2) return null;
    const vals = sh.getDataRange().getValues();
    let hr = -1;
    for (let r = 0; r < vals.length; r++) {
      const low = vals[r].map(c => String(c || '').trim().toLowerCase());
      if (low[0] === 'кабель' && low.indexOf('d1') >= 0) { hr = r; break; }
    }
    if (hr < 0) return null;
    const headers = vals[hr].map(c => String(c || '').trim().toLowerCase());
    const d1 = headers.indexOf('d1'), d2 = headers.indexOf('d2'), d3 = headers.indexOf('d3');
    const norm = s => String(s || '').toLowerCase().replace(/[\s-]/g, '');
    const get  = (row, i) => (i >= 0 ? String(row[i] || '').trim() : '');
    const byCable = {};
    for (let r = hr + 1; r < vals.length; r++) {
      const mark = String(vals[r][0] || '').trim();
      if (!mark) break;                                  // пустая строка = конец коакс-таблицы
      byCable[norm(mark)] = { d1: get(vals[r], d1), d2: get(vals[r], d2), d3: get(vals[r], d3) };
    }
    return { byCable };
  } catch (e) { return null; }
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
      if (h.length > 2 && !/[\s-]/.test(h)) return;
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
      // Кол-во: пустая ячейка → 1 (по умолчанию); явный 0 → позиция НЕ используется (пропуск,
      // «0 не выводить вообще»); число (в т.ч. дробная норма провода) — как есть.
      const rawQty = qtyIdx >= 0 ? drow[qtyIdx] : '';
      const qty = (rawQty === '' || rawQty == null) ? 1 : (Number(String(rawQty).replace(',', '.')) || 0);
      if (qty === 0) continue;
      components.push({
        id:     `spy-${dr}`,
        type,
        name,
        art,
        qty,
        side:   sideIdx   >= 0 ? String(drow[sideIdx]   || '').trim().toUpperCase() : '',
        length: lengthIdx >= 0 ? (Number(drow[lengthIdx]) || 0) : 0,
      });
    }
    return components;
  }
  return [];
}

function readTerRecordsForGenerator_(preloaded) {
  try {
    let snapshot = preloaded || getTechOperationsSnapshot_();
    // Auto-resync if schema version changed (e.g. new L+/L- extraction was added).
    // Передан готовый снапшот (открытие генератора) → схему уже выверил
    // ensureTechOperationsSnapshotReady_, повторный ресинк не нужен.
    if (!preloaded && String(snapshot.meta && snapshot.meta.schemaVersion) !== String(TECHOPS_DB_APP.schemaVersion)) {
      syncTechOperationsDatabase();
      snapshot = getTechOperationsSnapshot_();
    }
    // Диапазон совместимого провода (От/До AWG, От/До мм²) лежит в exportValues —
    // ищем по именам колонок (meta.columnHeadersByTab.ter), без отдельных extra-полей.
    const terHeaders = ((snapshot.meta && snapshot.meta.columnHeadersByTab) || {}).ter || [];
    const norm_ = h => String(h || '').toLowerCase().replace(/[\s.]/g, '');
    const colIdx_ = pred => terHeaders.findIndex(h => pred(norm_(h)));
    const awgFromC = colIdx_(h => h.indexOf('от') === 0 && h.indexOf('awg') >= 0);
    const awgToC   = colIdx_(h => h.indexOf('до') === 0 && h.indexOf('awg') >= 0);
    const mm2FromC = colIdx_(h => h.indexOf('от') === 0 && (h.indexOf('мм2') >= 0 || h.indexOf('мм²') >= 0));
    const mm2ToC   = colIdx_(h => h.indexOf('до') === 0 && (h.indexOf('мм2') >= 0 || h.indexOf('мм²') >= 0));
    const cell_ = (exp, i) => (i >= 0 ? String(exp[i] || '') : '');
    const recs = (snapshot.records || [])
      .filter(r => r.tabKey === 'ter' && r.terArticle)
      .map(r => {
        const exp = r.exportValues || [];
        return {
          article:      r.terArticle    || '',
          step:         r.terStep       || '',
          strip:        r.terStrip      || '',
          lPlus:        r.terLPlus      || '',
          lMinus:       r.terLMinus     || '',
          applicator:   r.terApplicator || '',
          crimpHeight:  r.terCrimpHeight  || '',
          pullForceMin: r.terPullForceMin || '',
          pullForceMax: r.terPullForceMax || '',
          awgFrom:      cell_(exp, awgFromC),
          awgTo:        cell_(exp, awgToC),
          mm2From:      cell_(exp, mm2FromC),
          mm2To:        cell_(exp, mm2ToC),
        };
      });
    return recs;
  } catch (e) { return []; }
}

function readOpRecordsForGenerator_(preloaded) {
  try {
    const snapshot = preloaded || getTechOperationsSnapshot_();
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

// Коакс-изделия из снапшота (вкладка coax). Доп-колонки (D1..D3, L1..L3, L+/L−,
// «Тип пина», «Тип экрана») лежат в exportValues — читаем по именам заголовков
// (meta.columnHeadersByTab.coax), без отдельных extra-полей и без бампа схемы (как ter-ридер).
function readCoaxRecordsForGenerator_(preloaded) {
  try {
    const snapshot = preloaded || getTechOperationsSnapshot_();
    const headers  = ((snapshot.meta && snapshot.meta.columnHeadersByTab) || {}).coax || [];
    const norm_    = h => String(h || '').trim().toLowerCase();
    const find_    = pred => headers.findIndex(h => pred(norm_(h)));
    const col = {
      article:    find_(h => h === 'артикул'),
      type:       find_(h => h === 'тип/серия' || h === 'тип / серия' || h === 'тип серия'),
      wire:       find_(h => h === 'провод'),
      mfr:        find_(h => h === 'производитель' || h === 'бренд'),
      program:    find_(h => h === 'программа'),
      d1: find_(h => h === 'd1'), d2: find_(h => h === 'd2'), d3: find_(h => h === 'd3'),
      l1: find_(h => h === 'l1'), l2: find_(h => h === 'l2'), l3: find_(h => h === 'l3'),
      lPlus:      find_(h => h === 'l+'),
      lMinus:     find_(h => h === 'l-' || h === 'l−'),
      pinType:    find_(h => h === 'тип пина'),
      shieldType: find_(h => h === 'тип экрана'),
      connArticle: find_(h => h === 'артикул разъёма' || h === 'артикул разьёма' || h === 'артикул разъема'),
    };
    const cell = (exp, i) => (i >= 0 ? String(exp[i] || '').trim() : '');
    return (snapshot.records || [])
      .filter(r => r.tabKey === 'coax')
      .map(r => {
        const e = r.exportValues || [];
        return {
          article:    r.coaxArticle || cell(e, col.article),
          type:       r.coaxType    || cell(e, col.type),
          wire:       r.coaxWire    || cell(e, col.wire),
          mfr:        r.coaxMfr     || cell(e, col.mfr),
          program:    cell(e, col.program),
          d1: cell(e, col.d1), d2: cell(e, col.d2), d3: cell(e, col.d3),
          l1: cell(e, col.l1), l2: cell(e, col.l2), l3: cell(e, col.l3),
          lPlus:      cell(e, col.lPlus),
          lMinus:     cell(e, col.lMinus),
          pinType:    cell(e, col.pinType),
          shieldType: cell(e, col.shieldType),
          connArticle: cell(e, col.connArticle),
        };
      });
  } catch (e) { return []; }
}

// Добавляет строку в БД.КОАКС из диалога генератора (когда BOM-разъёма нет в базе).
// Пишет по ИМЕНАМ колонок в источник; инвалидирует кеш снапшота (подхватится при след.
// открытии). Активную таблицу НЕ трогает (источник — отдельный документ). Возвращает
// запись в формате readCoaxRecordsForGenerator_ для немедленного матча в диалоге.
function addCoaxRecordToDb(payload) {
  payload = payload || {};
  const connArticle = String(payload.connArticle || '').trim();
  if (!connArticle) throw new Error('Не задан артикул разъёма.');
  const src   = SpreadsheetApp.openById(getSourceSpreadsheetId_());
  const sheet = src.getSheetByName(TECHOPS_DB_APP.tabs.coax.sourceSheetName);
  if (!sheet) throw new Error('Нет листа БД.КОАКС.');
  const hr  = TECHOPS_DB_APP.tabs.coax.headerRowNumber || 2;
  const lc  = sheet.getLastColumn();
  const hdr = sheet.getRange(hr, 1, 1, lc).getDisplayValues()[0]
    .map(h => String(h).replace(/\s+/g, ' ').trim().toLowerCase());
  const row = new Array(lc).fill('');
  const set = (name, val) => {
    const i = hdr.indexOf(String(name).toLowerCase());
    if (i >= 0 && val != null && String(val) !== '') row[i] = val;
  };
  const descr = [payload.type, payload.wire].filter(Boolean).join(' ');
  // Колонка «Артикул» (col A) = part-номер разъёма (ключ матчинга); если его нет — описание.
  // Колонка «Артикул разъёма» (col P) — рудимент: пишем для совместимости, если она ещё есть
  // (set() безопасно no-op, когда колонку удалили).
  const article = connArticle || descr;
  set('артикул', article);
  set('тип/серия', payload.type);
  set('провод', payload.wire);
  set('производитель', payload.mfr);
  set('программа', payload.program);
  set('d1', payload.d1); set('d2', payload.d2); set('d3', payload.d3);
  set('l1', payload.l1); set('l2', payload.l2); set('l3', payload.l3);
  set('l+', payload.lPlus); set('l-', payload.lMinus);
  set('тип пина', payload.pinType);
  set('тип экрана', payload.shieldType);
  set('артикул разъёма', connArticle);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, lc).setValues([row]);
  try { getTechOpsCache_().clear(); } catch (e) {}   // следующее открытие пере-синкнет снапшот
  return {
    article: article, type: payload.type || '', wire: payload.wire || '',
    mfr: payload.mfr || '', program: payload.program || '',
    d1: payload.d1 || '', d2: payload.d2 || '', d3: payload.d3 || '',
    l1: payload.l1 || '', l2: payload.l2 || '', l3: payload.l3 || '',
    lPlus: payload.lPlus || '', lMinus: payload.lMinus || '',
    pinType: payload.pinType || '', shieldType: payload.shieldType || '',
    connArticle: connArticle,
  };
}

// ── Generator ─────────────────────────────────────────────────

/**
 * @typedef {Object} AssemblyWire   Один провод изделия.
 * @property {string} name    Наименование провода.
 * @property {string} art     Артикул.
 * @property {string|number} qty     Количество.
 * @property {string|number} length  Длина, мм.
 */

/**
 * @typedef {Object} AssemblySide   Сторона A или B (терминал + разъём).
 * @property {string} termName  Наименование терминала.
 * @property {string} termArt   Артикул терминала.
 * @property {string|number} termQty
 * @property {string} connName  Наименование разъёма.
 * @property {string} connArt   Артикул разъёма.
 * @property {string|number} connQty
 */

/**
 * @typedef {Object} AssemblyOp   Одна операция в цепочке генерации.
 * @property {'cutWire'|'prsTerm'|'insTerm'|'solderConn'|'twist'|'tin'} type  Тип операции.
 * @property {'A'|'B'} [side]     Сторона эндпоинта (для term/solder-операций).
 * @property {Object} [ep]        Снимок эндпоинта: {connArt, connName, termArt, termName, sdrTmpl, double}.
 * @property {Array}  [wires]     Маршрутизированные провода операции (для term/solder).
 * @property {string} templateId  ID шаблона техкарты для вставки.
 * @property {string} [opNum]     Номер операции (CODE из БД.ОП) для поиска времени.
 * @property {string|number} [tPrep]     Подготовительное время.
 * @property {string|number} [tOp]       Оперативное время.
 * @property {string|number} [tMachine]  Машинное время.
 */

/**
 * @typedef {Object} AssemblyConfig   Полный конфиг генерации, собранный диалогом.
 * @property {string|number} assemblyIndex  Индекс/обозначение изделия.
 * @property {string} assemblyName          Наименование изделия (идёт в результат последнего листа).
 * @property {string|number} partQty        Количество изделий в партии.
 * @property {AssemblyWire[]} wires         Провода изделия.
 * @property {AssemblyOp[]}   ops           Цепочка операций (только с templateId генерируются).
 * @property {AssemblySide|null} sideA      Сторона A или null.
 * @property {AssemblySide|null} sideB      Сторона B или null.
 */

/**
 * Генерирует техкарты сборки: по листу на каждую операцию с templateId.
 * Результат каждого листа (норма + наименование) переносится в «вход» следующего;
 * на последнем листе полуфабрикат переименовывается в изделие (config.assemblyName).
 * @param {AssemblyConfig} config
 * @returns {Object} Сводка результата генерации для клиента.
 */
function generateAssemblyTechCards(config) {
  if (!config || !Array.isArray(config.ops) || !config.ops.length) {
    throw new Error('Нет операций для создания.');
  }

  const ss = SpreadsheetApp.getActive();
  const isCoax = config.mode === 'coax';
  const createdSheets = [];
  let prevResult = '';
  let prevResultNorm = ''; // норма выхода предыдущей операции → норма входа текущей
  // Коакс: ПОСТОРОННЯЯ цепочка — pin A продолжает strip A, pin B продолжает strip B (а не
  // последнюю по списку). Иначе вход «Пайка A» = «Разделка B» (две стороны режутся подряд).
  const coaxBySide = {};      // 'A'|'B' → последний результат стороны
  const coaxNormBySide = {};  // 'A'|'B' → норма выхода стороны
  let tutResult = '', tutNorm = '';   // параллельная заготовка ТУТ (резка ТУТ) → 2-й вход «Монтаж ТУТ»
  const terRecords = isCoax ? [] : readTerRecordsForGenerator_();

  // Только операции с выбранным шаблоном; последняя из них — финальный лист (изделие).
  const ops = config.ops.filter(o => o.templateId);

  // Регенерация = замена: сносим листы прошлого прогона (только свои, по списку в
  // свойствах документа) ДО вставки — иначе createUniqueSheet_ плодит дубли «-2/-3»,
  // а «Удалить созданные» видит лишь последнюю пачку. Ручные листы не трогаем.
  const docProps = PropertiesService.getDocumentProperties();
  purgeGeneratedSheets_(safeJsonParse_(docProps.getProperty('TECHMAP_LAST_GENERATED'), []));

  try {
    for (let i = 0; i < ops.length; i++) {
      const op = ops[i];
      const isLast = i === ops.length - 1;

      const wireData = (!isCoax && isCutOp_(op.type))
        ? buildCombinedWireData_(cutWiresOf_(op, config))
        : null;

      // noReplace: при N разъёмах несколько операций используют один шаблон — замена по
      // заголовку снесла бы лист предыдущего разъёма. Суффикс делает имя листа читаемым/уникальным.
      const insertResult = insertTemplate(op.templateId, { noReplace: true, nameSuffix: sheetSuffix_(op) });
      const sheet = ss.getSheetByName(insertResult.sheetName);
      if (!sheet) throw new Error(`Лист "${insertResult.sheetName}" не найден.`);

      // Вход карты: для коакс-операции СО стороной — последний результат той же стороны
      // (pin A ← strip A), иначе общий prevResult.
      const sideKey  = isCoax && (op.side === 'A' || op.side === 'B') ? op.side : null;
      const inResult = (sideKey && coaxBySide[sideKey] != null) ? coaxBySide[sideKey] : prevResult;
      const inNorm   = (sideKey && coaxNormBySide[sideKey] != null) ? coaxNormBySide[sideKey] : prevResultNorm;

      const thisResult = computeOperationResult_(op, config, inResult, wireData, op.side);

      // Резка ТУТ — параллельная заготовка: запоминаем её результат и норму (отрезки),
      // чтобы «Монтаж ТУТ» показал её во ВТОРОМ ряду входа «Полуфабрикат».
      if (op.type === 'cutTut')          { tutResult = thisResult; tutNorm = String((op.tutCount || 0) * partQty_(config)); }
      else if (op.type === 'coaxCutTut') { tutResult = thisResult; tutNorm = String(Math.max(1, Number(op.units) || 1) * partQty_(config)); }
      if (op.type === 'insTut' || op.type === 'coaxInsTut') op._auxInput = { name: tutResult, norm: tutNorm };

      const terData = isCoax ? null : buildTerData_(op, config, terRecords);
      const phMap = buildPlaceholderMap_(op, config, inResult, thisResult, wireData, terData);
      const sheetState = replacePlaceholders_(sheet, phMap);

      fillTechCardStructurally_(sheet, op, op.type, config, inResult, thisResult, wireData, terData, sheetState, inNorm, isLast);

      // CUT_TUT — параллельная заготовка (режет ТУТ), главный полуфабрикат-кабель не двигает:
      // его результат НЕ перетекает в следующую операцию (STRIP берёт кабель от CUT_WIRE).
      // Резка/надевание ТУТ — параллельная подготовка (термоусадка), главную цепочку не двигают.
      if ((!isCoax || coaxAdvancesMain_(op.type)) && op.type !== 'cutTut' && op.type !== 'insTut') {
        prevResult = thisResult;
        prevResultNorm = computeOutputNorm_(op, config);
        if (sideKey) { coaxBySide[sideKey] = thisResult; coaxNormBySide[sideKey] = prevResultNorm; }
      }
      createdSheets.push(insertResult.sheetName);
    }

    // Запоминаем созданные листы — кнопка «Удалить созданные» в диалоге их снесёт
    // (а следующая генерация авто-зачистит по этому же списку).
    docProps.setProperty('TECHMAP_LAST_GENERATED', JSON.stringify(createdSheets));
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

// Удаляет листы последней генерации (имена сохранены в свойствах документа).
// Используется кнопкой «Удалить созданные» в диалоге для быстрых итераций тестов.
function deleteLastGeneratedSheets() {
  const props = PropertiesService.getDocumentProperties();
  const names = safeJsonParse_(props.getProperty('TECHMAP_LAST_GENERATED'), []);
  const deleted = purgeGeneratedSheets_(names || []);
  props.deleteProperty('TECHMAP_LAST_GENERATED');
  return { deleted: deleted, names: names || [] };
}

// Сносит перечисленные листы (только существующие; служебные/ручные не трогает),
// затем возвращает фокус на рабочий лист и прячет служебные обратно. Возвращает
// число удалённых. Общий путь для кнопки «Удалить созданные» и авто-зачистки перед
// регенерацией.
function purgeGeneratedSheets_(names) {
  if (!Array.isArray(names) || !names.length) return 0;
  const ss = SpreadsheetApp.getActive();
  let deleted = 0;
  names.forEach((name) => {
    const s = ss.getSheetByName(name);
    if (s && !isSystemSheet_(name) && ss.getSheets().length > 1) { ss.deleteSheet(s); deleted++; }
  });

  // После удаления активного листа Sheets перескакивает на соседний и мог
  // открыть/показать служебный (_TC_TECHOPS_DB и т.п.) — возвращаем фокус на
  // рабочий лист и прячем служебные обратно.
  if (deleted > 0) {
    const active = ss.getActiveSheet();
    if (active && isSystemSheet_(active.getName())) {
      const fallback = ss.getSheets().find((sh) => !isSystemSheet_(sh.getName()));
      if (fallback) ss.setActiveSheet(fallback);
    }
    ss.getSheets().forEach((sh) => {
      if (isSystemSheet_(sh.getName()) && !sh.isSheetHidden()) sh.hideSheet();
    });
  }
  return deleted;
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

// Подпись проводов свивки. multiline=true → пары по строкам с нумерацией
// (читаемо для ячейки карты); иначе компактно в одну строку (для результата).
function twistWiresLabel_(config, multiline) {
  const t      = config.twist || {};
  const wires  = Array.isArray(config.wires) ? config.wires : [];
  const nameOf = i => { const w = wires[i]; return w ? (w.art || w.name || '').trim() : ''; };
  if (t.mode === 'pairs' && Array.isArray(t.pairs) && t.pairs.length) {
    const list = t.pairs.map(p => (p || []).map(nameOf).filter(Boolean).join(' + ')).filter(Boolean);
    return multiline ? list.map((s, i) => `${i + 1}) ${s}`).join('\n') : list.join('; ');
  }
  const idx = Array.isArray(t.wireIndices) ? t.wireIndices : [];
  return idx.map(nameOf).filter(Boolean).join(multiline ? '\n' : ', ');
}

// Число свивок: количество пар (режим пар) либо 1 (режим «все вместе»).
function twistCount_(config) {
  const t = config.twist || {};
  if (t.mode === 'pairs' && Array.isArray(t.pairs)) return t.pairs.length || 0;
  return 1;
}

// ── Коакс-сборка ──────────────────────────────────────────────
// Все коакс-опера­ции имеют тип с префиксом 'coax'. Перечень опкеев и их шаблонов —
// в клиенте (getActiveOpsCoax_). Сторона: A = кириллица «А», B = латиница «B» (как в каталоге).
function isCoaxOp_(opType)  { return typeof opType === 'string' && opType.indexOf('coax') === 0; }
function coaxSideLabel_(side) { return side === 'B' ? 'B' : 'А'; }

// Двигает ли операция «главный» полуфабрикат-кабель (результат перетекает дальше).
// CUT_TUT — единственная параллельная заготовка (режет ТУТ), главный поток не трогает.
function coaxAdvancesMain_(opType) { return isCoaxOp_(opType) && opType !== 'coaxCutTut'; }

// Наименование результата коакс-операции (накопительное, перетекает в вход следующей).
// Топология жгута: «Разъём ‹артА› — кабель ‹арт› — разъём ‹артВ›» (В — если двусторонний).
// Единая база наименований всех СБОРОЧНЫХ коакс-операций (после пайки) — чтобы везде однотипно.
// Кабель с длиной: «LMR-100 600мм» (длина — если задана).
function coaxCableRef_(cx) {
  const cable = cx.cableArt || cx.cableName || (cx.sideA && cx.sideA.wire) || 'кабель';
  return cx.cableLength ? `${cable} ${cx.cableLength}мм` : cable;
}
function coaxHarness_(cx) {
  cx = cx || {};
  const cA = (cx.sideA && (cx.sideA.article || cx.sideA.connName)) || 'разъём';
  let s = `Разъём ${cA} — кабель ${coaxCableRef_(cx)}`;
  const cB = cx.sideB && (cx.sideB.article || cx.sideB.connName);
  if (cB) s += ` — разъём ${cB}`;
  return s;
}

// Наименование коакс-полуфабриката. Заготовки (резка/разделка) — описание этапа; сборочные
// операции (пайка→…→усадка) — единая база «топология жгута» + что сделали (однотипно).
function computeCoaxResult_(opType, config, prevResult, side) {
  const cx    = config.coax || {};
  const sd    = side === 'B' ? (cx.sideB || {}) : (cx.sideA || {});
  const S     = coaxSideLabel_(side);                      // 'А'/'B' (коакс латиница B, как каталог)
  // Артикул кабеля приоритетнее ГРН (BOM-имя бывает мусорным, напр. «2»).
  const cable = cx.cableArt || cx.cableName || (cx.sideA && cx.sideA.wire) || 'кабель';
  const conn  = sd.article || 'разъём';
  switch (opType) {
    case 'coaxCut':                                                         // заготовка: «артикул Lмм»
      return [cable, cx.cableLength ? `${cx.cableLength}мм` : ''].filter(Boolean).join(' ');
    case 'coaxCutTut': {                                                    // параллельная заготовка
      const tlen = cx.tutLength ? `${cx.tutLength}мм` : '';
      const mat  = cx.tutArt || cx.tutName || 'ТУТ';
      return [mat, tlen].filter(Boolean).join(', отрезок ');
    }
    // Привязка к кабелю/разъёму; без «П/ф:» (по требованию — убрать приписку во всём генераторе).
    case 'coaxStrip':        return `Разделка ${S} под ${conn} на кабеле: ${cable}`;
    case 'coaxPin': {
      // Топология жгута появляется здесь. Сторона A — частичная (разъём А + кабель);
      // сторона B — полная (разъём А — кабель — разъём В).
      if (side === 'B') return coaxHarness_(cx);
      const cA = (cx.sideA && (cx.sideA.article || cx.sideA.connName)) || 'разъём';
      return `Разъём ${cA} — кабель ${coaxCableRef_(cx)}`;
    }
    case 'coaxInsTut':       return `Кабель ${coaxCableRef_(cx)} с надетыми ТУТ`;
    case 'coaxInsSleeve':    return `Кабель ${coaxCableRef_(cx)} с надетыми гильзами`;
    // После топологии — однотипно: «топология (что сделали)».
    case 'coaxHousing':      return `${coaxHarness_(cx)} (корпус смонтирован, ст. ${S})`;
    case 'coaxShield':       return `${coaxHarness_(cx)} (экран опрессован)`;
    case 'coaxSolderShield': return `${coaxHarness_(cx)} (экран припаян, ст. ${S})`;
    case 'coaxCapCrimp':     return `${coaxHarness_(cx)} (крышка обжата)`;
    case 'coaxCapScrew':     return `${coaxHarness_(cx)} (крышка закреплена винтами)`;
    case 'coaxHeat':         return `${coaxHarness_(cx)} (ТУТ усажен)`;
    default:                 return prevResult || '';
  }
}

// ── Endpoint-aware op helpers (N разъёмов на сторону + роутинг) ──
// Новый контракт: term/solder-операции несут op.ep (снимок эндпоинта:
// {connArt, connName, termArt, termName, sdrTmpl, double}) и op.wires (маршрутизированные
// провода). Коакс (opType 'coax*') и его config.coax.sideA/B — отдельная ветка, не трогаем.
function termKind_(op) {
  var t = op && op.type;
  if (t === 'prsTerm') return 'prs';
  if (t === 'insTerm') return 'ins';
  if (t === 'solderConn') return 'sdr';
  return null;
}
function opEndpoint_(op) { return (op && op.ep) || {}; }
function opSideLabel_(op) { return (op && op.side === 'B') ? 'В' : 'А'; }
// Локация в наименовании п/ф: для сложных сборок (>2 разъёмов) — наименование разъёма по
// чертежу (ep.label), иначе «ст. А/В». Сторона неоднозначна, когда на ней несколько разъёмов.
function opLoc_(op, config) {
  var ep = opEndpoint_(op);
  var label = ep.label || ep.connName || '';
  if (config && config.bigAssembly && label) return label;
  return 'ст. ' + opSideLabel_(op);
}
function opRoutedWires_(op, config) {
  if (op && Array.isArray(op.wires)) return op.wires;
  return Array.isArray(config.wires) ? config.wires : [];
}
// Число терминалов операции: контакты на маршрут-провода; двойной обжим = пары (2 провода
// в 1 терминал).
function opTermCount_(op, config) {
  var w = opRoutedWires_(op, config);
  var n = w.reduce(function (s, x) { return s + (Number(x.qty) || 1); }, 0);
  return (op.ep && op.ep.double) ? Math.ceil(n / 2) : n;
}
// Сторона A строится из НАРЕЗАННЫХ проводов (по строке: опрессовка/монтаж каждого).
// Сторона B продолжается из готовой заготовки стороны A (вход/выход одной строкой).
function opIsBuildSide_(op) { return !op || op.side !== 'B'; }
// Резка провода: обычная (cutWire) и с разделкой (cutWireStrip — пред-зачистка концов под
// пайку/двойной обжим). cutTut (резка ТУТ) — отдельный тип, не провод.
function isCutOp_(t) { return t === 'cutWire' || t === 'cutWireStrip'; }
function cutWiresOf_(op, config) {
  return (op && Array.isArray(op.wires)) ? op.wires : (Array.isArray(config.wires) ? config.wires : []);
}
// Суффикс имени листа: для term/solder — разъём; для резки под двойной обжим — пометка
// (чтобы два листа резки с одним шаблоном не схлопнулись).
function sheetSuffix_(op) {
  if (op && op.type === 'cutWireStrip') return ' — разделка';
  if (op && op.type === 'cutTut') return ' — ТУТ';
  if (!termKind_(op)) return '';
  var ep = opEndpoint_(op);
  var conn = ep.connName || ep.connArt || '';
  return conn ? (' — ' + conn) : '';
}

function computeOperationResult_(opOrType, config, prevResult, wireData, opSide) {
  var op = (opOrType && typeof opOrType === 'object') ? opOrType : { type: opOrType, side: opSide };
  var opType = op.type;
  if (isCoaxOp_(opType)) return computeCoaxResult_(opType, config, prevResult, op.side);
  const wd    = wireData || {};
  const kind  = termKind_(op);

  // Наименование полуфабриката — КРАТКОЕ описание этапа (без перечисления всех проводов и
  // без накопления всей истории — иначе ячейка превращается в нечитаемую простыню).
  if (kind === 'prs') {
    const ep  = opEndpoint_(op);
    const t   = ep.termArt || ep.termName || '';
    const n   = opRoutedWires_(op, config).length;
    const dbl = ep.double ? ', двойной обжим' : '';
    return `Заготовка: ${n} пров. с обжатым терм. ${t}${dbl} (${opLoc_(op, config)})`;
  }
  if (kind === 'ins') {
    const ep   = opEndpoint_(op);
    const n    = opRoutedWires_(op, config).length;
    const conn = ep.connArt || ep.connName || '';
    return `Разъём ${conn} смонтирован (${opLoc_(op, config)}, ${n} пров.)`;
  }
  if (kind === 'sdr') {
    const ep   = opEndpoint_(op);
    const conn = ep.connArt || ep.connName || '';
    return `Разъём ${conn} припаян (${opLoc_(op, config)})`;
  }

  if (opType === 'cutTut') {
    const mat  = op.tutArt || op.tutName || 'ТУТ';
    const segs = Array.isArray(op.tutSegs) ? op.tutSegs : [];
    if (segs.some(s => s.len)) {
      const parts = segs.map(s => s.len ? `${s.count}×${s.len}мм` : `${s.count} шт`);
      return `${mat}: ${parts.join(' + ')}`;
    }
    const c = op.tutCount || 0;
    return `${mat}${c ? ', ' + c + ' отрезков' : ''}`.trim();
  }
  if (opType === 'insTut') {
    const c = op.tutCount || 0;
    return `ТУТ надет на провода${c ? ' (' + c + ' конц.)' : ''}`;
  }

  if (isCutOp_(opType)) {
    const cw  = cutWiresOf_(op, config);
    const src = cw.length > 0 ? cw : (function() {
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

  switch (opType) {
    case 'twist': {
      // Свивка не меняет наименование полуфабриката — просто помечаем «(со свивкой)».
      // Детали (пары проводов + шаг) выводятся в маркер-ячейки шаблона, не в результат.
      return prevResult ? `${prevResult} (со свивкой)` : 'Свивка';
    }
    case 'tin': {
      // Лужение не меняет наименование полуфабриката — помечаем «(с лужением)».
      return prevResult ? `${prevResult} (с лужением)` : 'Лужение';
    }
    case 'heatTut': {
      // Усадка ТУТ — финишная операция, помечаем «(ТУТ усажен)».
      return prevResult ? `${prevResult} (ТУТ усажен)` : 'Усадка ТУТ';
    }
    default:         return prevResult;
  }
}

// Resolves ter record for terminal operations and builds terData object.
function buildTerData_(op, config, terRecords) {
  const kind = termKind_(op);
  if (kind !== 'prs' && kind !== 'ins') return null;
  const side = opEndpoint_(op);
  if (!side.termArt) return null;
  const termArt   = side.termArt || '';
  const termExact = termArt.toLowerCase().trim();
  const normStr   = s => s.toLowerCase().replace(/[\s./()_-]/g, '');
  const termNorm  = normStr(termArt);
  const normR     = r => normStr(r.article || '');
  let rec = (terRecords || []).find(r => (r.article || '').toLowerCase().trim() === termExact)
         || (termNorm.length >= 3 ? (terRecords || []).find(r => normR(r) === termNorm) : null);
  // Pass 3: strip leading manufacturer word ("MOLEX 5034290000" → "5034290000")
  if (!rec) {
    const words = termArt.trim().split(/\s+/);
    if (words.length > 1) {
      const normNoMfr = normStr(words.slice(1).join(' '));
      if (normNoMfr.length >= 3) rec = (terRecords || []).find(r => normR(r) === normNoMfr);
    }
  }
  return {
    applicator:  rec && rec.applicator  ? rec.applicator  : ASSEMBLY_GEN.noDataMark,
    crimpHeight: rec ? (rec.crimpHeight  || '') : '',
    pullForce:   rec ? [rec.pullForceMin, rec.pullForceMax].filter(Boolean).join(' - ') : '',
  };
}

function buildPlaceholderMap_(op, config, prevResult, thisResult, wireData, terData) {
  const p  = ASSEMBLY_GEN.placeholders;
  const wd = wireData || {};
  // Плейсхолдеры стороны (легаси): для endpoint-операции берём op.ep в слот её стороны.
  const ep = op && op.ep;
  const epSide = ep ? {
    termName: ep.termName || '', termArt: ep.termArt || '', termQty: opTermCount_(op, config),
    connName: ep.connName || '', connArt: ep.connArt || '', connQty: opTermCount_(op, config),
  } : null;
  const sA = (epSide && op.side !== 'B') ? epSide : (config.sideA || {});
  const sB = (epSide && op.side === 'B') ? epSide : (config.sideB || {});

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
  const rate = ASSEMBLY_GEN.toleranceMmPerM;
  if (!rate) return '';
  const raw = len / 1000 * rate;
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
  if (lastRow < 1 || lastCol < 1) return null;

  const range    = sheet.getRange(1, 1, lastRow, lastCol);
  const values   = range.getValues();
  const formulas = range.getFormulas();
  const mergeData = buildMergeData_(sheet); // читаем мёрджи ДО записей: переиспользуются в structural fill (−1 чтение, −1 flush)
  const mergeMap  = mergeData.map;

  const phEntries = Object.entries(phMap); // один раз, не на каждую ячейку
  const dirtyRows = {};
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) continue;
      const val = values[r][c];
      if (typeof val !== 'string' || val === '') continue;
      if (val.indexOf('{{') < 0 && val.indexOf('[') < 0) continue; // быстрый отсев ячеек без плейсхолдеров
      let nv = val;
      for (const [token, repl] of phEntries) {
        if (nv.includes(token)) nv = nv.split(token).join(String(repl));
      }
      if (nv !== val) {
        values[r][c] = nv; // держим in-memory копию в актуальном состоянии для повторного использования
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

  // Отдаём пост-замены значения, формулы и мёрджи, чтобы fillTechCardStructurally_
  // не перечитывал лист повторно (экономия 3 полных чтений на операцию).
  return { values, formulas, mergeMap, merges: mergeData.list, lastRow, lastCol };
}

// ── Structural fill helpers ───────────────────────────────────

// Scans values for section anchor rows.
function detectSections_(values) {
  let kp = -1, sfI = -1, res = -1, sfO = -1, tm = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (!cell) continue;
      if (tm  < 0 && /расс?ч[её]?тное\s*врем/i.test(cell))       { tm  = r; break; }
      if (res < 0 && tm < 0 && cell === 'результат')              { res = r; break; }
      if (kp  < 0 && cell.includes('комплектующ'))                { kp  = r; break; }
      if (/полуфабрикат|^п\/ф/.test(cell)) {
        if (res >= 0 || tm >= 0) { if (sfO < 0) { sfO = r; break; } }
        else                     { if (sfI < 0) { sfI = r; break; } }
      }
    }
  }
  return { kompRow: kp, sfInRow: sfI, resultRow: res, sfOutRow: sfO, timeRow: tm };
}

// Creates a mutable sheet context object; ctx.refresh() re-reads all sheet data.
// Опциональный seed {values, formulas} (из replacePlaceholders_) позволяет
// пропустить начальные getValues/getFormulas — лист уже прочитан до этого.
function makeSheetCtx_(sheet, seed) {
  const ctx = { sheet };
  ctx.refresh = function() {
    ctx.lastRow = sheet.getLastRow();
    ctx.lastCol = sheet.getLastColumn();
    if (ctx.lastRow < 1 || ctx.lastCol < 1) { ctx.values = []; ctx.formulas = []; ctx.mergeMap = {}; ctx.sections = {}; return; }
    ctx.values   = sheet.getRange(1, 1, ctx.lastRow, ctx.lastCol).getValues();
    ctx.formulas = sheet.getRange(1, 1, ctx.lastRow, ctx.lastCol).getFormulas();
    const md = buildMergeData_(sheet);
    ctx.mergeMap = md.map; ctx.merges = md.list;
    ctx.sections = detectSections_(ctx.values);
  };

  // Лёгкое обновление после вставки count строк после 1-based afterRow.
  // Значения и объединения читаем с листа (корректно под объединениями —
  // новые строки под вертикальной меткой возвращаются пустыми). Формулы НЕ
  // перечитываем (в слотах данных их нет) — вставляем пустые, экономя 1 чтение.
  ctx.applyInsert = function(afterRow, count) {
    if (!ctx.lastCol || !ctx.values) { ctx.refresh(); return; }
    // Вставляем пустые строки в память на позицию afterRow (0-based индекс = afterRow):
    // существующие строки сдвигаются (значения не меняются), новые слоты пусты —
    // совпадает с листом для логики заполнения (якоря секций + пустота слотов).
    // Не читаем getValues/getFormulas с листа — экономия 2 чтений на вставку.
    const emptyRow = new Array(ctx.lastCol).fill('');
    for (let i = 0; i < count; i++) {
      ctx.values.splice(afterRow, 0, emptyRow.slice());
      ctx.formulas.splice(afterRow, 0, emptyRow.slice());
    }
    ctx.lastRow = ctx.values.length;
    // Объединения: если API-путь посчитал их в памяти — берём их (0 чтений с листа),
    // иначе (legacy/сбой) читаем с листа.
    if (_lastPostMerges) {
      ctx.merges = _lastPostMerges;
      ctx.mergeMap = buildMergeMapFromList_(_lastPostMerges);
    } else {
      const md = buildMergeData_(sheet);
      ctx.mergeMap = md.map; ctx.merges = md.list;
    }
    ctx.sections = detectSections_(ctx.values);
  };

  if (seed && seed.values && seed.values.length) {
    ctx.lastRow  = seed.values.length;
    ctx.lastCol  = seed.values[0] ? seed.values[0].length : 0;
    ctx.values   = seed.values;
    ctx.formulas = seed.formulas;
    ctx.mergeMap = seed.mergeMap || buildMergeMap_(sheet); // мёрджи уже прочитаны в replacePlaceholders_
    ctx.merges   = seed.merges || [];
    ctx.sections = detectSections_(ctx.values);
  } else {
    ctx.refresh();
  }
  return ctx;
}

// Scans ctx.values for global column header positions.
function detectGlobalColumns_(values) {
  let artCol = -1, grnCol = -1, normCol = -1, seqCol = -1, unitCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] || '').toLowerCase().trim();
      if (artCol  < 0 && (cell === 'артикул'  || cell === 'art' || cell === 'обозначение'))         artCol  = c;
      if (grnCol  < 0 && (cell === 'грн'       || cell === 'наименование' || cell === 'название'
                          || (cell.includes('грн') && cell.length < 6)))                             grnCol  = c;
      if (normCol < 0 && (cell === 'норма'     || cell === 'кол-во' || cell === 'qty'
                          || (cell.includes('норма') && cell.length < 10)))                          normCol = c;
      if (seqCol  < 0 && cell === '№')                                                               seqCol  = c;
      if (unitCol < 0 && (cell.indexOf('ед.изм') === 0 || cell === 'ед' || cell === 'единица'))      unitCol = c;
    }
    if (artCol >= 0 && grnCol >= 0 && normCol >= 0) break;
  }
  return { artCol, grnCol, normCol, seqCol, unitCol };
}

// Writes a value to a cell, skipping real formulas and resolving merges.
function setCell_(ctx, r, col, val) {
  if (col < 0 || col >= ctx.lastCol) return;
  const v = val == null ? '' : String(val);
  if (!v) return;
  const fml = ctx.formulas[r] && ctx.formulas[r][col];
  if (fml && !/^=""$|^=''$/.test(fml.trim())) return;
  let wr = r + 1, wc = col + 1;
  if (ctx.mergeMap[wr] && ctx.mergeMap[wr][wc]) { const p = ctx.mergeMap[wr][wc]; wr = p.r; wc = p.c; }
  try { ctx.sheet.getRange(wr, wc).setValue(v); } catch (e) {}
}

// Resolves art/name/norm column indices for a data row (looks for header above, falls back to globals).
function resolveCols_(ctx, colMap, dataRow) {
  const loc = findColHeadersAbove_(ctx.values, dataRow);
  // Нашли локальную шапку секции (есть «Наименование/ГРН») → доверяем ей ЦЕЛИКОМ, включая
  // отсутствие «Артикул» (art=-1). Иначе откат к global artCol ловит «Обозначение» из шапки
  // операций и пишет комплектующее в колонку-ярлык (затирая «Комплектующие»).
  if (loc.name >= 0) {
    return { art: loc.art, name: loc.name, norm: loc.norm >= 0 ? loc.norm : colMap.normCol };
  }
  return { art: colMap.artCol, name: colMap.grnCol, norm: colMap.normCol };
}

// Quantity multiplier per assembly; defaults to 1 when not specified/invalid.
function partQty_(config) {
  const q = config && config.partQty;
  return q > 0 ? q : 1;
}

// Base wire label: "<артикул|имя> <длина>мм" (пустые части отбрасываются).
function wireBaseName_(w) {
  return [w.art || w.name, w.length ? w.length + 'мм' : ''].filter(Boolean).join(' ');
}

// Finds the first cell matching `pattern` (tested against trimmed text).
// Returns 0-based {row, col}; {row:-1, col:-1} if not found.
function findColByText_(values, pattern) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (pattern.test(String(row[c] || '').trim())) return { row: r, col: c };
    }
  }
  return { row: -1, col: -1 };
}

// Норма ВЫХОДА операции = норма ВХОДА следующей (результат перетекает в след. лист).
// Для одиночной заготовки (ins/prsB) — (кол-во контактов на провод × кол-во сборок).
// Для пооперационных (cut/prsA) выход пооводной — одиночной нормы нет (следующая
// операция берёт вход пооводно), возвращаем ''.
function computeOutputNorm_(opOrType, config) {
  const op = (opOrType && typeof opOrType === 'object') ? opOrType : { type: opOrType };
  const opType = op.type;
  // Коакс: один кабель → один полуфабрикат на изделие; норма входа след. карты = шт×партия.
  if (isCoaxOp_(opType)) return String(partQty_(config));
  const pQty = partQty_(config);
  // Монтаж = собранный полуфабрикат (норма = контактов×партия). Опрессовка стороны A —
  // пооводный выход (нормы нет); стороны B — единая заготовка (норма = партия).
  const kind = termKind_(op);
  if (kind === 'ins' || kind === 'sdr') {
    const w = opRoutedWires_(op, config);
    return w.length ? String((w[0].qty || 1) * pQty) : '';
  }
  if (kind === 'prs') return opIsBuildSide_(op) ? '' : String(pQty);
  // Свивка/лужение/усадка дают один полуфабрикат на изделие → норма в шт.
  if (opType === 'twist' || opType === 'tin' || opType === 'heatTut') return String(pQty);
  return '';
}

// Перезаписывает ярлык «Полуфабрикат»/«П/Ф» в строке rowIdx (0-based) на newLabel.
function relabelSemifinished_(sheet, ctx, rowIdx, newLabel) {
  const row = ctx.values[rowIdx] || [];
  for (let c = 0; c < row.length; c++) {
    if (/полуфабрикат|^п\/ф/i.test(String(row[c] || '').trim())) {
      fillMergedCell_(sheet, rowIdx + 1, c + 1, newLabel, ctx.mergeMap);
      return;
    }
  }
}

// Ячейка считается «пустой» для заполнения, если реально пуста ИЛИ содержит
// любой видимый плейсхолдер ручной сборки в угловых кавычках (‹наименование›,
// ‹норма›, ‹артикул›, ‹…›) — генератор его перезаписывает.
function isPlaceholderOrEmpty_(value) {
  const s = String(value == null ? '' : value).trim();
  return s === '' || /^‹.*›$/.test(s);
}

// Группирует провода резки по артикулу, суммируя метраж (одинаковые провода = одна строка).
function groupCutKompl_(wires, pQty) {
  const map = {}, order = [];
  (wires || []).forEach(function (w) {
    const key = String(w.art || w.name || '').trim();
    const meters = (w.qty > 0 && w.length > 0) ? (w.qty * w.length * pQty / 1000)
                 : (w.length > 0 ? (w.length * pQty / 1000) : 0);
    if (!map[key]) { map[key] = { art: w.art || w.name || '', name: w.name || '', meters: 0 }; order.push(key); }
    map[key].meters += meters;
  });
  return order.map(function (k) {
    return { art: map[k].art, name: map[k].name, norm: map[k].meters > 0 ? formatDecimalComma_(map[k].meters, 4) : '' };
  });
}

// Builds the wire result name for terminal operations (used in Результат and Время sections).
function buildWireResultName_(w, termArt, connArt, sideLabel) {
  const base = wireBaseName_(w);
  const withTerm = termArt ? base + ', обжатый терминал ' + termArt : base;
  return connArt ? withTerm + ' → ' + connArt + ' сторона ' + sideLabel : withTerm;
}

// Двойной обжим: маршрут-провода группируются ПО ПАРАМ (2 провода в 1 терминал, соседние в
// порядке роутинга). Возвращает по строке на терминал: «{пров1} + {пров2}, обж. терм. {арт}».
function doubleCrimpRows_(wires, termArt, loc) {
  const rows = [];
  for (let i = 0; i < wires.length; i += 2) {
    const a = wireBaseName_(wires[i]);
    const b = wires[i + 1] ? wireBaseName_(wires[i + 1]) : '';
    const pair = b ? (a + ' + ' + b) : a;
    rows.push(pair + ', обж. терм. ' + termArt + ' (двойной обжим, ' + loc + ')');
  }
  return rows;
}

// Finds empty slots starting at anchorRow up to bound; expands by inserting rows if needed.
// nameCol — 0-based column used to test emptiness. srcRow — 1-based source row for copy.
// Returns array of 0-based row indices for the slots.
function findAndExpandSlots_(sheet, ctx, anchorRow, bound, nameCol, neededCount, srcRow) {
  const slots = [anchorRow];
  for (let r = anchorRow + 1; r < bound; r++) {
    const fc = String((ctx.values[r] || [])[0] || '').toLowerCase().trim();
    if (fc && !/полуфабрикат|^п\/ф/.test(fc) && !/^\d+$/.test(fc)) break;
    const nameV = nameCol >= 0 ? String((ctx.values[r] || [])[nameCol] || '').trim() : '';
    if (isPlaceholderOrEmpty_(nameV)) slots.push(r);
    else break;
  }
  if (neededCount > slots.length) {
    const insertCount = neededCount - slots.length;
    const lastSlot = slots[slots.length - 1];
    if (insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, srcRow, ctx.merges)) {
      for (let i = 0; i < insertCount; i++) slots.push(lastSlot + 1 + i);
      ctx.applyInsert(lastSlot + 1, insertCount);
    }
  }
  return slots;
}

// Writes sequence number to seqCol, respecting merges.
function writeSeqNum_(sheet, rowNum, seqCol, idx, mergeMap) {
  if (seqCol < 0) return;
  let sr = rowNum, sc = seqCol + 1;
  if (mergeMap[sr] && mergeMap[sr][sc]) { const p = mergeMap[sr][sc]; sr = p.r; sc = p.c; }
  try { sheet.getRange(sr, sc).setValue(idx + 1); } catch (e) {}
}

// ── Section fill functions ────────────────────────────────────

function fillKompl_(sheet, ctx, colMap, op, config, wireData) {
  const { kompRow, sfInRow, resultRow, timeRow } = ctx.sections;
  const cols  = resolveCols_(ctx, colMap, kompRow);
  const pQty  = partQty_(config);
  const opType = op.type;

  // Коакс: кабель — на карте резки (норма в метрах), разъём — на пайке пина и монтаже корпуса.
  if (isCoaxOp_(opType)) {
    const cx   = config.coax || {};
    const side = op.side === 'B' ? (cx.sideB || {}) : (cx.sideA || {});
    let comp = null;
    // ГРН берём из BOM КАК ЕСТЬ (если пусто — пусто, НЕ дублируем артикул).
    const cableArt  = cx.cableArt  || (cx.sideA && cx.sideA.wire) || '';
    const connName  = side.connName || '';   // имя разъёма из BOM (общий ГРН подузлов стороны)
    if (opType === 'coaxCut') {
      const len = parseFloat(String(cx.cableLength || '').replace(',', '.')) || 0;
      comp = { art: cableArt, name: cx.cableName || '', norm: len > 0 ? formatDecimalComma_(len * pQty / 1000, 4) : '' };
    } else if (opType === 'coaxStrip' && side.article) {
      // Разделка идёт ПОД разъём этой стороны — он и есть комплектующее (под что зачищаем).
      comp = { art: side.article, name: connName, norm: String(pQty) };
    } else if (opType === 'coaxCutTut') {
      // Расход ТУТ = длина × число концов (op.units) × партия / 1000 (метры). Для 2 сторон — 2 отрезка.
      // У CUT_TUT нет колонки «Артикул» — материал по «Наименованию», поэтому имя = арт как fallback.
      const len   = parseFloat(String(cx.tutLength || '').replace(',', '.')) || 0;
      const units = Math.max(1, Number(op.units) || 1);
      comp = { art: cx.tutArt || '', name: cx.tutName || cx.tutArt || '', norm: len > 0 ? formatDecimalComma_(len * units * pQty / 1000, 4) : '' };
    } else if ((opType === 'coaxPin' || opType === 'coaxHousing') && side.article) {
      // Артикул разъёма + подпись подузла (пин/корпус); ГРН — имя разъёма из BOM (ОБЩИЙ
      // для всех подузлов стороны), пусто если в BOM имени нет.
      const part = opType === 'coaxPin' ? ' (пин)' : ' (корпус)';
      comp = { art: side.article + part, name: connName, norm: String(pQty) };
    } else if (opType === 'coaxInsSleeve' && side.article) {
      // Гильза экрана — из того же разъёма (артикул + «(гильза)»); норма = число концов × партия.
      const units = Math.max(1, Number(op.units) || 1);
      comp = { art: side.article + ' (гильза)', name: connName, norm: String(units * pQty) };
    }
    // Крышка (coaxCapCrimp/coaxCapScrew) — операция над полуфабрикатом, без отдельного
    // комплектующего: разъём (вместе с крышкой, один артикул) уже посчитан на coaxPin.
    // Артикул разъёма виден в наименовании результата. Двойной учёт не нужен.
    if (comp) {
      setCell_(ctx, kompRow, cols.art,  comp.art);
      setCell_(ctx, kompRow, cols.name, comp.name);
      setCell_(ctx, kompRow, cols.norm, comp.norm);
    }
    return;
  }

  // Резка ТУТ: комплектующее = материал ТУТ. Задана длина отрезков → норма в МЕТРАХ
  // (Σ длина×кол-во × партия / 1000, как коакс); длины нет → fallback в штуках.
  if (opType === 'cutTut') {
    const segs = Array.isArray(op.tutSegs) ? op.tutSegs : [];
    const totalMm = segs.reduce((s, x) =>
      s + (parseFloat(String(x.len || '').replace(',', '.')) || 0) * (Number(x.count) || 0), 0);
    setCell_(ctx, kompRow, cols.art,  op.tutArt || '');
    setCell_(ctx, kompRow, cols.name, op.tutName || op.tutArt || '');
    if (totalMm > 0) {
      setCell_(ctx, kompRow, cols.norm, formatDecimalComma_(totalMm * pQty / 1000, 4));
      if (colMap.unitCol >= 0) setCell_(ctx, kompRow, colMap.unitCol, 'М.');
    } else {
      setCell_(ctx, kompRow, cols.norm, String((op.tutCount || 0) * pQty));
      if (colMap.unitCol >= 0) setCell_(ctx, kompRow, colMap.unitCol, 'Шт.');
    }
    return;
  }

  // Резка: одинаковые провода (один артикул) складываем в одну строку с суммой метража.
  const _cutW = isCutOp_(opType) ? cutWiresOf_(op, config) : [];
  const wires = (_cutW.length > 0) ? groupCutKompl_(_cutW, pQty) : null;
  // Опрессовка → комплектующее = ТЕРМИНАЛ, норма = число терминалов (двойной обжим = пары).
  // Монтаж/пайка → комплектующее = РАЗЪЁМ (корпус), норма = 1 корпус × партия.
  const ep   = opEndpoint_(op);
  const kind = termKind_(op);
  // ГРН (name) — из BOM КАК ЕСТЬ (пусто, если в BOM имени нет; НЕ дублируем артикул).
  const comp =
      kind === 'prs' ? { art: ep.termArt || ep.termName || '', name: ep.termName || '', norm: String(opTermCount_(op, config) * pQty) }
    : (kind === 'ins' || kind === 'sdr') ? { art: ep.connArt || ep.connName || '', name: ep.connName || '', norm: String(pQty) }
    : null;

  if (wires) {
    const bound = Math.min(...[sfInRow, resultRow, timeRow].filter(x => x > kompRow).concat([ctx.values.length]));
    const slots = [kompRow];
    for (let r = kompRow + 1; r < bound; r++) {
      const lbl = String((ctx.values[r] || [])[0] || '').trim();
      if (lbl && !/^\d+$/.test(lbl)) break;
      const artV  = cols.art  >= 0 ? String((ctx.values[r] || [])[cols.art]  || '').trim() : '';
      const grnV  = cols.name >= 0 ? String((ctx.values[r] || [])[cols.name] || '').trim() : '';
      if (isPlaceholderOrEmpty_(artV) && isPlaceholderOrEmpty_(grnV)) slots.push(r);
      else break;
    }
    if (wires.length > slots.length) {
      const insertCount = wires.length - slots.length;
      const lastSlot = slots[slots.length - 1];
      if (insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, kompRow + 1, ctx.merges)) {
        for (let i = 0; i < insertCount; i++) slots.push(lastSlot + 1 + i);
        ctx.applyInsert(lastSlot + 1, insertCount);
      }
    }
    for (let i = 0; i < Math.min(wires.length, slots.length); i++) {
      const w = wires[i];   // {art, name, norm} — уже сгруппировано/суммировано
      const r = slots[i];
      writeSeqNum_(sheet, r + 1, colMap.seqCol, i, ctx.mergeMap);
      setCell_(ctx, r, cols.art,  w.art  || w.name || '');
      setCell_(ctx, r, cols.name, w.name || '');
      setCell_(ctx, r, cols.norm, w.norm);
    }
  } else if (comp) {
    setCell_(ctx, kompRow, cols.art,  comp.art);
    setCell_(ctx, kompRow, cols.name, comp.name);
    setCell_(ctx, kompRow, cols.norm, comp.norm);
    // Терминал/разъём считаются в штуках — перетираем шаблонную «М.» (метры провода).
    if (colMap.unitCol >= 0) setCell_(ctx, kompRow, colMap.unitCol, 'Шт.');
  }
}

function fillSfIn_(sheet, ctx, colMap, config, prevResult, op, prevResultNorm) {
  const { sfInRow, resultRow } = ctx.sections;
  const cols  = resolveCols_(ctx, colMap, sfInRow);
  const nameC = cols.name >= 0 ? cols.name : cols.art;
  const normC = cols.norm >= 0 ? cols.norm : colMap.normCol;
  if (nameC < 0) return;

  const wires = opRoutedWires_(op, config);   // маршрут-провода эндпоинта (или все, легаси)
  const ep    = opEndpoint_(op);
  const kind  = termKind_(op);
  const pQty  = partQty_(config);

  // Уникальные ветви-источники (собранные ранее заготовки, приходящие на этот разъём).
  const branches = [];
  (Array.isArray(op.wireSources) ? op.wireSources : []).forEach(s => {
    if (s && branches.indexOf(s) < 0) branches.push(s);
  });

  // Имена и нормы строк входа (параллельные массивы). Приоритет:
  //   опрессовка разъёма-слияния → по ВЕТВЯМ (входящие заготовки);
  //   монтаж с двойным обжимом → по спаренным терминалам;
  //   сторона A → провода ПОСТРОЧНО (опрессовка нарезанные / монтаж с терминалом);
  //   иначе → одна заготовка стороны (prevResult).
  let wireNames = null, wireNorms = null;
  if (kind === 'prs' && branches.length > 0) {
    wireNames = branches.map(b => 'Заготовка ветви ' + b);
    wireNorms = branches.map(() => String(pQty));
  } else if (kind === 'ins' && ep.double && wires.length >= 2) {
    wireNames = doubleCrimpRows_(wires, ep.termArt || ep.termName || '', opLoc_(op, config));
    wireNorms = wireNames.map(() => String(pQty));
  } else if (opIsBuildSide_(op) && kind === 'prs' && wires.length > 0) {
    wireNames = wires.map(wireBaseName_);
    wireNorms = wires.map(w => formatDecimalComma_(w.qty * pQty, 4));
  } else if (opIsBuildSide_(op) && (kind === 'ins' || kind === 'sdr') && wires.length > 0) {
    const t = ep.termArt || ep.termName || '';   // пайка → терминала нет, имя = провод
    wireNames = wires.map(w => buildWireResultName_(w, t, '', opSideLabel_(op)));
    wireNorms = wires.map(w => formatDecimalComma_(w.qty * pQty, 4));
  }

  if (wireNames && wireNames.length > 0) {
    const bound = resultRow > sfInRow ? resultRow : ctx.values.length;
    const slots = findAndExpandSlots_(sheet, ctx, sfInRow, bound, nameC, wireNames.length, sfInRow + 1);
    for (let i = 0; i < Math.min(wireNames.length, slots.length); i++) {
      const rowNum = slots[i] + 1;
      writeSeqNum_(sheet, rowNum, colMap.seqCol, i, ctx.mergeMap);
      fillMergedCell_(sheet, rowNum, nameC + 1, wireNames[i], ctx.mergeMap);
      if (normC >= 0 && wireNorms) fillMergedCell_(sheet, rowNum, normC + 1, wireNorms[i], ctx.mergeMap);
    }
  } else {
    setCell_(ctx, sfInRow, nameC, prevResult);
    // Норма входа = норма выхода предыдущей операции (тянется с прошлого листа).
    if (normC >= 0 && prevResultNorm) setCell_(ctx, sfInRow, normC, prevResultNorm);
    // Параллельный вход (нарезанный ТУТ) → СЛЕДУЮЩИЙ ряд «Полуфабрикат» секции входа.
    const aux = op && op._auxInput;
    if (aux && aux.name) {
      const bound = resultRow > sfInRow ? resultRow : ctx.values.length;
      for (let r = sfInRow + 1; r < bound; r++) {
        const isSf = (ctx.values[r] || []).some(c => /полуфабрикат|^п\/ф/i.test(String(c || '').trim()));
        if (!isSf) continue;
        writeSeqNum_(sheet, r + 1, colMap.seqCol, 1, ctx.mergeMap);
        setCell_(ctx, r, nameC, aux.name);
        if (normC >= 0 && aux.norm) setCell_(ctx, r, normC, aux.norm);
        break;
      }
    }
  }
}

function fillSfOut_(sheet, ctx, colMap, config, thisResult, op, isLast) {
  // На последнем листе результат — это готовое ИЗДЕЛИЕ: имя берём из наименования
  // изделия, а ярлык «Полуфабрикат» меняем на «Изделие».
  const opType   = op && op.type;
  const singleName = (isLast && config.assemblyName) ? config.assemblyName : thisResult;
  const wires    = opRoutedWires_(op, config);   // маршрут-провода (или все, легаси)
  const ep       = opEndpoint_(op);
  const kind     = termKind_(op);
  const pQty     = partQty_(config);
  const isCutWire = isCutOp_(opType);
  const isTermOp  = kind === 'prs' || kind === 'ins';

  let fRes = -1, fSfOut = -1;
  for (let r = 0; r < ctx.values.length; r++) {
    for (let c = 0; c < (ctx.values[r] || []).length; c++) {
      const cell = String((ctx.values[r] || [])[c] || '').toLowerCase().trim();
      if (!cell) continue;
      if (fRes < 0 && cell === 'результат')                            { fRes   = r; break; }
      if (fRes >= 0 && fSfOut < 0 && /полуфабрикат|^п\/ф/.test(cell)) { fSfOut = r; break; }
    }
  }
  if (fSfOut < 0) return;

  const hdr   = findColHeadersAbove_(ctx.values, fSfOut);
  const nameC = hdr.name >= 0 ? hdr.name : (hdr.art >= 0 ? hdr.art : colMap.grnCol);
  const normC = hdr.norm >= 0 ? hdr.norm : colMap.normCol;

  // Двойной обжим (опрессовка) → по терминалам-парам (2 провода в 1 терминал).
  const isDouble = kind === 'prs' && ep.double && wires.length >= 2 && nameC >= 0;
  // Одной строкой: монтаж/пайка и опрессовка стороны B (единая заготовка). Резка и
  // опрессовка стороны A — по проводу.
  const isSingle = kind === 'ins' || kind === 'sdr' || (kind === 'prs' && !opIsBuildSide_(op));

  if (isDouble) {
    let resTimeBound = ctx.values.length;
    for (let r = fSfOut + 1; r < ctx.values.length; r++) {
      if ((ctx.values[r] || []).some(c => /расс?ч[её]?тное\s*врем/i.test(String(c || '')))) { resTimeBound = r; break; }
    }
    const rows  = doubleCrimpRows_(wires, ep.termArt || ep.termName || '', opLoc_(op, config));
    const slots = findAndExpandSlots_(sheet, ctx, fSfOut, resTimeBound, nameC, rows.length, fSfOut + 1);
    for (let i = 0; i < Math.min(rows.length, slots.length); i++) {
      const rowNum = slots[i] + 1;
      writeSeqNum_(sheet, rowNum, colMap.seqCol, i, ctx.mergeMap);
      fillMergedCell_(sheet, rowNum, nameC + 1, rows[i], ctx.mergeMap);
      if (normC >= 0) fillMergedCell_(sheet, rowNum, normC + 1, String(pQty), ctx.mergeMap);
    }
  } else if (isSingle && wires.length > 0 && nameC >= 0) {
    const insNorm = String((wires[0].qty || 1) * pQty);
    writeSeqNum_(sheet, fSfOut + 1, colMap.seqCol, 0, ctx.mergeMap);
    fillMergedCell_(sheet, fSfOut + 1, nameC + 1, singleName, ctx.mergeMap);
    if (normC >= 0) fillMergedCell_(sheet, fSfOut + 1, normC + 1, insNorm, ctx.mergeMap);
  } else if ((isCutWire || (kind === 'prs' && opIsBuildSide_(op))) && wires.length > 0 && nameC >= 0) {
    let resTimeBound = ctx.values.length;
    for (let r = fSfOut + 1; r < ctx.values.length; r++) {
      if ((ctx.values[r] || []).some(c => /расс?ч[её]?тное\s*врем/i.test(String(c || '')))) { resTimeBound = r; break; }
    }
    const slots    = findAndExpandSlots_(sheet, ctx, fSfOut, resTimeBound, nameC, wires.length, fSfOut + 1);
    const termArt  = ep.termArt || ep.termName || '';
    const sideLbl  = opSideLabel_(op);
    for (let i = 0; i < Math.min(wires.length, slots.length); i++) {
      const w      = wires[i];
      const rowNum = slots[i] + 1;
      const wName  = isTermOp
        ? buildWireResultName_(w, termArt, '', sideLbl)
        : wireBaseName_(w);
      writeSeqNum_(sheet, rowNum, colMap.seqCol, i, ctx.mergeMap);
      fillMergedCell_(sheet, rowNum, nameC + 1, wName, ctx.mergeMap);
      if (normC >= 0) fillMergedCell_(sheet, rowNum, normC + 1, formatDecimalComma_(w.qty * pQty, 4), ctx.mergeMap);
    }
  } else if (nameC >= 0) {
    fillMergedCell_(sheet, fSfOut + 1, nameC + 1, singleName, ctx.mergeMap);
    // Один полуфабрикат на изделие → норма результата = шт (иначе в шаблоне останется ‹норма›/#REF!).
    if ((opType === 'twist' || opType === 'tin' || opType === 'heatTut' || isCoaxOp_(opType)) && normC >= 0)
      fillMergedCell_(sheet, fSfOut + 1, normC + 1, String(pQty), ctx.mergeMap);
    if ((opType === 'cutTut' || opType === 'insTut') && normC >= 0)
      fillMergedCell_(sheet, fSfOut + 1, normC + 1, String((op.tutCount || 0) * pQty), ctx.mergeMap);
  }

  if (isLast) relabelSemifinished_(sheet, ctx, fSfOut, 'Изделие');
}

function fillTime_(sheet, ctx, colMap, op, config, thisResult, opType, isLast) {
  // На последнем листе наименование в Рассч. времени — это изделие, ярлык → «Изделие».
  const singleName = (isLast && config.assemblyName) ? config.assemblyName : thisResult;
  const { timeRow } = ctx.sections;
  const kind       = termKind_(op);
  const isCutWire  = isCutOp_(opType);
  // Пайка (sdr) считается как монтаж: одной строкой, норма = время × число паяных
  // соединений × партия. Без sdr она проваливалась в else и писала сырое tOp без
  // множителя (0,4 мин на 200 шт).
  const isTermOp   = kind === 'prs' || kind === 'ins' || kind === 'sdr';
  const wires      = opRoutedWires_(op, config);
  const ep         = opEndpoint_(op);
  const pQty       = partQty_(config);

  const timeDataRows = [];
  for (let r = timeRow + 1; r < ctx.values.length; r++) {
    const rc = (ctx.values[r] || []).map(c => String(c || '').toLowerCase().trim());
    if (!rc.some(v => v)) { if (timeDataRows.length > 0) break; continue; }
    if (rc.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение' || v === 'факт')) continue;
    timeDataRows.push(r);
  }
  if (!timeDataRows.length) return;

  const cols      = resolveCols_(ctx, colMap, timeDataRows[0]);
  const tNormCol  = cols.norm >= 0 ? cols.norm : colMap.normCol;
  const tNameCol  = cols.name >= 0 ? cols.name : colMap.grnCol;

  if ((isCutWire || isTermOp) && wires.length > 0 && tNameCol >= 0) {
    const tOpSec   = parseFloat(String(op.tOp   || '').replace(',', '.')) || 0;
    const tPrepSec = parseFloat(String(op.tPrep || '').replace(',', '.')) || 0;
    const tOpMin   = tOpSec  / 60;
    const tPrepMin = tPrepSec / 60;
    // Одной строкой: монтаж/пайка и опрессовка стороны B. Резка/опрессовка A — по проводу.
    const isIns    = kind === 'ins' || kind === 'sdr' || (kind === 'prs' && !opIsBuildSide_(op));

    if (isIns) {
      // INS: одна строка — суммарное время на все провода
      const totalQty = wires.reduce((s, w) => s + (w.qty || 1), 0);
      const rawNorm  = tOpMin * totalQty * pQty + tPrepMin;
      const wNorm    = formatDecimalComma_(rawNorm, 2);
      writeSeqNum_(sheet, timeDataRows[0] + 1, colMap.seqCol, 0, ctx.mergeMap);
      fillMergedCell_(sheet, timeDataRows[0] + 1, tNameCol + 1, singleName, ctx.mergeMap);
      if (tNormCol >= 0) fillMergedCell_(sheet, timeDataRows[0] + 1, tNormCol + 1, wNorm, ctx.mergeMap);
    } else {
      // CUT / PRS: по строке на каждый провод
      const timeSlots = [...timeDataRows];
      for (let r = timeDataRows[timeDataRows.length - 1] + 1; r < ctx.values.length; r++) {
        const rc = (ctx.values[r] || []).map(c => String(c || '').toLowerCase().trim());
        if (!rc.some(v => v)) break;
        if (rc.some(v => v === 'наименование' || v === 'норма' || v === 'обозначение' || v === 'факт')) continue;
        if (isPlaceholderOrEmpty_(String((ctx.values[r] || [])[tNameCol] || '').trim())) timeSlots.push(r);
        else break;
      }
      if (wires.length > timeSlots.length) {
        const insertCount = wires.length - timeSlots.length;
        const lastSlot    = timeSlots[timeSlots.length - 1];
        if (insertRowsAfterSafe_(sheet, lastSlot + 1, insertCount, timeDataRows[0] + 1, ctx.merges)) {
          for (let i = 0; i < insertCount; i++) timeSlots.push(lastSlot + 1 + i);
          ctx.applyInsert(lastSlot + 1, insertCount);
        }
      }
      const termArt  = isTermOp ? (ep.termArt || ep.termName || '') : '';
      const sideLbl  = opSideLabel_(op);
      const tPrepPerWire = wires.length > 1 ? tPrepMin / wires.length : tPrepMin;
      for (let i = 0; i < Math.min(wires.length, timeSlots.length); i++) {
        const w       = wires[i];
        const rowNum  = timeSlots[i] + 1;
        const wName   = isTermOp
          ? buildWireResultName_(w, termArt, '', sideLbl)
          : wireBaseName_(w);
        const rawNorm = (tOpMin > 0 ? tOpMin * w.qty * pQty : w.qty * pQty) + tPrepPerWire;
        const wNorm   = formatDecimalComma_(rawNorm, 2);
        writeSeqNum_(sheet, rowNum, colMap.seqCol, i, ctx.mergeMap);
        fillMergedCell_(sheet, rowNum, tNameCol + 1, wName, ctx.mergeMap);
        if (tNormCol >= 0) fillMergedCell_(sheet, rowNum, tNormCol + 1, wNorm, ctx.mergeMap);
      }
    }
  } else if (opType === 'twist') {
    // Время свивки = (опер. время на 1 свивку) × число свивок (пар) × партия + подгот.
    const tOpMin   = (parseFloat(String(op.tOp   || '').replace(',', '.')) || 0) / 60;
    const tPrepMin = (parseFloat(String(op.tPrep || '').replace(',', '.')) || 0) / 60;
    const cnt      = twistCount_(config);
    const norm     = formatDecimalComma_(tOpMin * cnt * pQty + tPrepMin, 2);
    writeSeqNum_(sheet, timeDataRows[0] + 1, colMap.seqCol, 0, ctx.mergeMap);
    fillMergedCell_(sheet, timeDataRows[0] + 1, tNameCol + 1, singleName, ctx.mergeMap);
    if (tNormCol >= 0) fillMergedCell_(sheet, timeDataRows[0] + 1, tNormCol + 1, norm, ctx.mergeMap);
  } else if (opType === 'tin') {
    // Время лужения = (опер. время на 1 конец) × число лужёных концов × партия + подгот.
    // Концы — только выбранные провода (config.tin.wireIndices); пусто → все.
    const tOpMin   = (parseFloat(String(op.tOp   || '').replace(',', '.')) || 0) / 60;
    const tPrepMin = (parseFloat(String(op.tPrep || '').replace(',', '.')) || 0) / 60;
    const tinIdx   = (config.tin && Array.isArray(config.tin.wireIndices)) ? config.tin.wireIndices : null;
    const tinned   = (tinIdx && tinIdx.length) ? tinIdx.map(i => wires[i]).filter(Boolean) : wires;
    const ends     = tinned.reduce((s, w) => s + (w.qty || 1), 0) || 1;
    const norm     = formatDecimalComma_(tOpMin * ends * pQty + tPrepMin, 2);
    writeSeqNum_(sheet, timeDataRows[0] + 1, colMap.seqCol, 0, ctx.mergeMap);
    fillMergedCell_(sheet, timeDataRows[0] + 1, tNameCol + 1, singleName, ctx.mergeMap);
    if (tNormCol >= 0) fillMergedCell_(sheet, timeDataRows[0] + 1, tNormCol + 1, norm, ctx.mergeMap);
  } else if (opType === 'cutTut' || opType === 'insTut' || opType === 'heatTut') {
    // ТУТ-операции: норма = время × число отрезков (паяных концов) × партия + подгот.
    // Без этой ветки уходили в else и писали сырое tOp без множителя.
    const tOpMin   = (parseFloat(String(op.tOp   || '').replace(',', '.')) || 0) / 60;
    const tPrepMin = (parseFloat(String(op.tPrep || '').replace(',', '.')) || 0) / 60;
    const cnt      = (op.tutCount || 0) || 1;
    const norm     = formatDecimalComma_(tOpMin * cnt * pQty + tPrepMin, 2);
    writeSeqNum_(sheet, timeDataRows[0] + 1, colMap.seqCol, 0, ctx.mergeMap);
    fillMergedCell_(sheet, timeDataRows[0] + 1, tNameCol + 1, singleName, ctx.mergeMap);
    if (tNormCol >= 0) fillMergedCell_(sheet, timeDataRows[0] + 1, tNormCol + 1, norm, ctx.mergeMap);
  } else if (isCoaxOp_(opType)) {
    // Коакс: норма = время × число концов (op.units) × партия + подгот. Однолистовые
    // операции (ТУТ/втулка/крышка/усадка) идут на оба конца за один лист → units=число
    // сторон; по-сторонние (резка/разделка/пайка/корпус) — units=1.
    const tOpMin   = (parseFloat(String(op.tOp   || '').replace(',', '.')) || 0) / 60;
    const tPrepMin = (parseFloat(String(op.tPrep || '').replace(',', '.')) || 0) / 60;
    const units    = Math.max(1, Number(op.units) || 1);
    const norm     = formatDecimalComma_(tOpMin * units * pQty + tPrepMin, 2);
    writeSeqNum_(sheet, timeDataRows[0] + 1, colMap.seqCol, 0, ctx.mergeMap);
    fillMergedCell_(sheet, timeDataRows[0] + 1, tNameCol + 1, singleName, ctx.mergeMap);
    if (tNormCol >= 0) fillMergedCell_(sheet, timeDataRows[0] + 1, tNormCol + 1, norm, ctx.mergeMap);
  } else {
    const tVals = timeDataRows.length <= 1
      ? [secToMin_(op.tOp)]
      : timeDataRows.length === 2
        ? [secToMin_(op.tPrep), secToMin_(op.tOp)]
        : [secToMin_(op.tPrep), secToMin_(op.tOp), secToMin_(op.tMachine)];
    for (let i = 0; i < timeDataRows.length && i < tVals.length; i++) {
      const r = timeDataRows[i];
      setCell_(ctx, r, tNormCol, tVals[i]);
      if (i === 0 && singleName) fillMergedCell_(sheet, r + 1, tNameCol + 1, singleName, ctx.mergeMap);
    }
  }

  if (isLast && timeDataRows.length) relabelSemifinished_(sheet, ctx, timeDataRows[0], 'Изделие');
}

function fillDopusk_(sheet, ctx, wireData) {
  const wd = wireData || {};
  const dopuskCol = findColByText_(ctx.values, /^допуск$/i).col;
  if (dopuskCol < 0) return;
  const tolStr = buildCutTolerance_(wd);
  const lenStr = buildCutLengthKd_(wd);
  for (let r = 0; r < ctx.values.length; r++) {
    const rowText = (ctx.values[r] || []).map(c => String(c || '').toLowerCase()).join(' ');
    const cur     = String((ctx.values[r] || [])[dopuskCol] || '').trim().toLowerCase();
    if (/тестовый\s*рез/.test(rowText) && (!cur || cur === '-' || cur === '—')) {
      fillMergedCell_(sheet, r + 1, dopuskCol + 1, tolStr, ctx.mergeMap);
    }
    if (lenStr && (cur === '[l кд]' || /измерить\s*длину|длина\s*должна/.test(rowText) && (!cur || cur === '[l кд]'))) {
      fillMergedCell_(sheet, r + 1, dopuskCol + 1, lenStr, ctx.mergeMap);
    }
  }
}

function fillTerminalFields_(sheet, ctx, terData) {
  const dopCol = findColByText_(ctx.values, /^допуск$/i).col;
  if (dopCol >= 0) {
    for (let r = 0; r < ctx.values.length; r++) {
      const rowText     = (ctx.values[r] || []).map(c => String(c || '').toLowerCase()).join(' ');
      const cur         = String((ctx.values[r] || [])[dopCol] || '').trim();
      const isControlRow = (ctx.values[r] || []).slice(0, 4).some(c => /^контроль$/i.test(String(c || '').trim()));
      if (!isControlRow) continue;
      if (!cur || cur === '-' || cur === '—') {
        // Контрольная строка ожидает значение из БД — если его нет, ставим маркер «нет данных».
        if (/datasheet|высота обжима|геометр/i.test(rowText)) {
          fillMergedCell_(sheet, r + 1, dopCol + 1, terData.crimpHeight || ASSEMBLY_GEN.noDataMark, ctx.mergeMap);
        }
        if (/pull[\s-]?test|разрыв|усилие обрыва/i.test(rowText)) {
          fillMergedCell_(sheet, r + 1, dopCol + 1, terData.pullForce || ASSEMBLY_GEN.noDataMark, ctx.mergeMap);
        }
      }
    }
  }
  if (!terData.applicator) return;
  const prog = findColByText_(ctx.values, /№\s*прог/i);
  const progRow = prog.row, progCol = prog.col;
  if (progRow >= 0 && progCol >= 0) {
    for (let dr = progRow + 1; dr < Math.min(progRow + 8, ctx.values.length); dr++) {
      if (!String((ctx.values[dr] || [])[progCol] || '').trim()) {
        fillMergedCell_(sheet, dr + 1, progCol + 1, terData.applicator, ctx.mergeMap);
        break;
      }
    }
  }
}

// Заполняет поля свивки по текстовым маркерам в шаблоне:
//   «вывести заданный шаг»            → значение шага свивки
//   «записать сюда какие провода …»   → список выбранных проводов
function fillTwistFields_(sheet, ctx, config) {
  const t        = config.twist || {};
  const wiresStr = twistWiresLabel_(config, true);   // пары по строкам, нумерованные
  const pitchStr = t.pitch ? `${t.pitch} мм` : '';
  for (let r = 0; r < ctx.values.length; r++) {
    const row = ctx.values[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').toLowerCase();
      if (!cell.trim()) continue;
      // Матч по вхождению ключевых слов (устойчив к написанию/пробелам/переносам).
      if (cell.includes('вывести') && cell.includes('шаг')) {
        fillMergedCell_(sheet, r + 1, c + 1, pitchStr, ctx.mergeMap);
      } else if (cell.includes('записать') && cell.includes('провод')) {
        fillMergedCell_(sheet, r + 1, c + 1, wiresStr, ctx.mergeMap);
      }
    }
  }
}

// Заполняет маркеры «Зачистка А/B: Внести данные» длиной зачистки соответствующей
// стороны (st.lenA/lenB). Только реально зачищаемые стороны (st.a/st.b).
// Цель — ячейки с «зачистка» И «внести» (узко, чтобы не задеть описания/дефекты).
function fillStripFields_(sheet, ctx, st) {
  st = st || {};
  for (let r = 0; r < ctx.values.length; r++) {
    const row = ctx.values[r] || [];
    for (let c = 0; c < row.length; c++) {
      const raw  = String(row[c] || '');
      const cell = raw.toLowerCase();
      const zi   = cell.indexOf('зачистка');
      if (zi < 0) continue;
      const tail = cell.slice(zi + 'зачистка'.length);
      let side = '';
      if (/^[\s:]*[aа]/.test(tail)) side = 'a';
      else if (/^[\s:]*[bв]/.test(tail)) side = 'b';
      else continue;

      const need = side === 'a' ? st.a : st.b;
      const len  = side === 'a' ? st.lenA : st.lenB;
      // Сторона не зачищается → прочерк; зачищается и длина есть → длина;
      // зачищается, но длина не подтянулась из БД → видимый маркер «нет данных».
      const val = !need ? '—' : (len ? formatStripLen_(len) : ASSEMBLY_GEN.noDataMark);

      if (cell.indexOf('внести') >= 0) {
        // Метка и плейсхолдер в одной ячейке — заменяем только «внести данные».
        fillMergedCell_(sheet, r + 1, c + 1, raw.replace(/внести\s*данные/ig, val), ctx.mergeMap);
      } else {
        // Плейсхолдер «Внести данные» в отдельной ячейке справа (метка и поле разнесены).
        for (let cc = c + 1; cc < row.length; cc++) {
          if (String(row[cc] || '').toLowerCase().indexOf('внести') >= 0) {
            fillMergedCell_(sheet, r + 1, cc + 1, val, ctx.mergeMap);
            break;
          }
        }
      }
    }
  }
}

// Длина зачистки → строка с «мм»; если значение уже содержит буквы (единицы) — как есть.
function formatStripLen_(v) {
  const s = String(v == null ? '' : v).trim();
  return /[a-zа-яё]/i.test(s) ? s : s + ' мм';
}

// ── Structural fill orchestrator ──────────────────────────────

function fillTechCardStructurally_(sheet, op, opType, config, prevResult, thisResult, wireData, terData, sheetState, prevResultNorm, isLast) {
  const ctx    = makeSheetCtx_(sheet, sheetState);
  if (ctx.lastRow < 1) return;
  const colMap = detectGlobalColumns_(ctx.values);
  const isCutWire = isCutOp_(opType);
  const wires = isCutWire && cutWiresOf_(op, config).length > 0 ? cutWiresOf_(op, config) : null;

  if (ctx.sections.kompRow >= 0)
    fillKompl_(sheet, ctx, colMap, op, config, wireData);
  const isTermOp_ = !!termKind_(op);
  // Коакс-резка (CUT_WIRE/CUT_TUT) — заготовка из сырья, входного полуфабриката нет:
  // первый «Полуфабрикат» — это уже выход (его заполнит fillSfOut_), вход не трогаем.
  const isCoaxCut_ = opType === 'coaxCut' || opType === 'coaxCutTut';
  if (ctx.sections.sfInRow >= 0 && (prevResult || isTermOp_) && !isCoaxCut_)
    fillSfIn_(sheet, ctx, colMap, config, prevResult, op, prevResultNorm);
  if (thisResult && (ctx.sections.sfOutRow >= 0 || ctx.sections.resultRow >= 0))
    fillSfOut_(sheet, ctx, colMap, config, thisResult, op, isLast);
  if (ctx.sections.timeRow >= 0)
    fillTime_(sheet, ctx, colMap, op, config, thisResult, opType, isLast);
  if (isCutWire && wires)
    fillDopusk_(sheet, ctx, wireData);
  if (termKind_(op) && terData)
    fillTerminalFields_(sheet, ctx, terData);
  if (opType === 'twist')
    fillTwistFields_(sheet, ctx, config);
  if (opType === 'cutWire' && config.strip)
    fillStripFields_(sheet, ctx, config.strip);
  // Резка под двойной обжим — зачистка одного конца (CUT_WIRE+1strip), длина из БД.ТЕР.
  if (opType === 'cutWireStrip' && op.strip)
    fillStripFields_(sheet, ctx, op.strip);
  if (isCoaxOp_(opType))
    fillCoaxHeader_(sheet, ctx, config);
  if (opType === 'coaxStrip')
    fillCoaxDims_(sheet, ctx, config, op);
}

// Размеры разделки коакса в STRIP-карте: ячейка-метка D1-D3/L1-L3/L+/L- → значение из
// матч-записи СТОРОНЫ операции (config.coax.sideA/sideB). «D1» → «1,3» (что есть что —
// показывает диаграмма). Нет данных по метке → оставляем метку как есть.
function fillCoaxDims_(sheet, ctx, config, op) {
  const cx   = config.coax || {};
  const side = op.side === 'B' ? (cx.sideB || {}) : (cx.sideA || {});
  const map  = { 'D1': side.d1, 'D2': side.d2, 'D3': side.d3,
                 'L1': side.l1, 'L2': side.l2, 'L3': side.l3,
                 'L+': side.lPlus, 'L-': side.lMinus, 'L−': side.lMinus };
  const vals = ctx.values;
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      const key = String(vals[r][c] || '').trim();
      if (!Object.prototype.hasOwnProperty.call(map, key)) continue;
      const val = map[key];
      if (val == null || String(val).trim() === '') continue;
      fillMergedCell_(sheet, r + 1, c + 1, String(val).replace('.', ','), ctx.mergeMap);
    }
  }
}

// Шапка коакс-карты: коакс-шаблоны не содержат {{плейсхолдеров}} — наименование изделия
// и код проекта зашиты как образец. Перезаписываем ячейку ПОД ярлыком-меткой.
function fillCoaxHeader_(sheet, ctx, config) {
  const name  = config.assemblyName  || '';
  const index = config.assemblyIndex || '';
  if (!name && !index) return;
  const setBelow = (pattern, val) => {
    if (!val) return;
    const hit = findColByText_(ctx.values, pattern);
    if (hit.row < 0 || hit.row + 1 >= ctx.values.length) return;
    fillMergedCell_(sheet, hit.row + 2, hit.col + 1, val, ctx.mergeMap); // строка ниже (1-based +2)
  };
  setBelow(/наименование\s+издели/i, name);
  setBelow(/номер\s+проекта|код\s+издели/i, index);
}

// ── Row insertion ─────────────────────────────────────────────

// Список объединений ПОСЛЕ последней API-вставки (посчитан в памяти из pre-merges).
// applyInsert использует его, чтобы не читать getMergedRanges с листа. null → legacy.
var _lastPostMerges = null;

// Строит mergeMap[row][col]={r,c} (top-left) из списка прямоугольников [{r1,r2,c1,c2}].
function buildMergeMapFromList_(list) {
  const map = {};
  (list || []).forEach((m) => {
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r !== m.r1 || c !== m.c1) {
          if (!map[r]) map[r] = {};
          map[r][c] = { r: m.r1, c: m.c1 };
        }
      }
    }
  });
  return map;
}

// Inserts `count` rows after `afterRow` (1-based), copying srcRow as the template.
// Быстрый путь — один Sheets API batchUpdate (insertDimension сам сдвигает все
// объединения ниже, copyPaste тиражирует строку-шаблон). При сбое/недоступности
// API откатывается на проверенный SpreadsheetApp-путь (legacy).
// knownMerges (опц.) — уже прочитанный список объединений [{r1,r2,c1,c2}] из ctx,
// чтобы не читать getMergedRanges повторно на каждую вставку.
function insertRowsAfterSafe_(sheet, afterRow, count, srcRow, knownMerges) {
  _lastPostMerges = null; // сбрасываем: только успешный API-путь выставит его
  if (typeof Sheets !== 'undefined') {
    try {
      return insertRowsViaSheetsApi_(sheet, afterRow, count, srcRow, knownMerges);
    } catch (e) {
      // API недоступен или вызов упал — падаем на проверенный legacy-путь ниже
    }
  }
  return insertRowsAfterLegacy_(sheet, afterRow, count, srcRow, knownMerges);
}

/**
 * Вставка строк одним вызовом Sheets API. Последовательность в одном batchUpdate
 * (атомарно, без мигания):
 *   1) insertDimension — вставляет строки, сам сдвигает содержимое/объединения ниже;
 *   2) unmergeCells по новым строкам — снимает объединения, которые insertDimension
 *      растянул на новые строки (вертикальные метки секций, пересекающие точку вставки),
 *      иначе copyPaste падает «нельзя вставить в диапазон с объединениями»;
 *   3) copyPaste строки-шаблона в каждую новую строку;
 *   4) mergeCells — пересоздаёт пересекающие объединения уже растянутыми на count.
 */
function insertRowsViaSheetsApi_(sheet, afterRow, count, srcRow, knownMerges) {
  // КРИТИЧНО: сбросить отложенные записи SpreadsheetApp ДО серверной правки,
  // иначе они лягут на дореставочные позиции строк после сдвига.
  SpreadsheetApp.flush();

  const sheetId     = sheet.getSheetId();
  const lastCol     = sheet.getLastColumn();
  const adjustedSrc = srcRow > afterRow ? srcRow + count : srcRow; // позиция шаблона после вставки
  const srcIdx      = adjustedSrc - 1; // 0-based

  // Объединения, ПЕРЕСЕКАЮЩИЕ точку вставки (r1 ≤ afterRow < r2) — их insertDimension
  // растянет на новые строки. Снимем по новым строкам и пересоздадим растянутыми.
  // Берём уже прочитанный список из ctx (knownMerges), иначе читаем сами.
  const allMerges = knownMerges || sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getMergedRanges()
    .map(r => ({ r1: r.getRow(), r2: r.getLastRow(), c1: r.getColumn(), c2: r.getLastColumn() }));
  const crossing = allMerges.filter(m => m.r1 <= afterRow && m.r2 > afterRow);

  const requests = [{
    insertDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: afterRow, endIndex: afterRow + count },
      inheritFromBefore: afterRow > 0,
    },
  }];

  // Снимаем каждое пересекающее объединение по ПОЛНОМУ (растянутому) диапазону —
  // unmergeCells требует выделить весь диапазон объединения, не часть.
  crossing.forEach(m => {
    requests.push({
      unmergeCells: {
        range: { sheetId, startRowIndex: m.r1 - 1, endRowIndex: m.r2 + count, startColumnIndex: m.c1 - 1, endColumnIndex: m.c2 },
      },
    });
  });

  for (let i = 0; i < count; i++) {
    const destIdx = afterRow + i; // 0-based новая строка
    requests.push({
      copyPaste: {
        source:      { sheetId, startRowIndex: srcIdx,  endRowIndex: srcIdx + 1,  startColumnIndex: 0, endColumnIndex: lastCol },
        destination: { sheetId, startRowIndex: destIdx, endRowIndex: destIdx + 1, startColumnIndex: 0, endColumnIndex: lastCol },
        pasteType: 'PASTE_NORMAL',
        pasteOrientation: 'NORMAL',
      },
    });
  }

  // Восстанавливаем пересекающие объединения уже растянутыми на count строк.
  crossing.forEach(m => {
    requests.push({
      mergeCells: {
        mergeType: 'MERGE_ALL',
        range: { sheetId, startRowIndex: m.r1 - 1, endRowIndex: m.r2 + count, startColumnIndex: m.c1 - 1, endColumnIndex: m.c2 },
      },
    });
  });

  Sheets.Spreadsheets.batchUpdate({ requests }, SpreadsheetApp.getActive().getId());

  // Считаем НОВЫЙ список объединений в памяти (insertDimension сдвигает/растягивает
  // по предсказуемым правилам) — чтобы applyInsert не читал getMergedRanges с листа:
  //   выше точки вставки (r2 ≤ afterRow) — без изменений;
  //   ниже (r1 > afterRow) — сдвиг на +count;
  //   пересекающие — растянуты до r2+count;
  //   + горизонтальные объединения строки-шаблона тиражируются в новые строки.
  const post = [];
  allMerges.forEach((m) => {
    if (m.r2 <= afterRow)      post.push({ r1: m.r1, r2: m.r2, c1: m.c1, c2: m.c2 });
    else if (m.r1 > afterRow)  post.push({ r1: m.r1 + count, r2: m.r2 + count, c1: m.c1, c2: m.c2 });
    else                       post.push({ r1: m.r1, r2: m.r2 + count, c1: m.c1, c2: m.c2 });
  });
  const srcHoriz = allMerges.filter((m) => m.r1 === srcRow && m.r2 === srcRow && m.c2 > m.c1);
  for (let i = 0; i < count; i++) {
    const nr = afterRow + 1 + i;
    srcHoriz.forEach((m) => post.push({ r1: nr, r2: nr, c1: m.c1, c2: m.c2 }));
  }
  _lastPostMerges = post;
  return true;
}

// Проверенный SpreadsheetApp-путь: разрыв затронутых объединений → insertRowsAfter →
// построчный copyTo → восстановление. Используется как fallback.
function insertRowsAfterLegacy_(sheet, afterRow, count, srcRow, knownMerges) {
  const lastCol = sheet.getLastColumn();

  const allMerges = knownMerges || sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getMergedRanges().map(r => ({
      r1: r.getRow(), r2: r.getLastRow(),
      c1: r.getColumn(), c2: r.getLastColumn()
    }));
  // insertRowsAfter не трогает объединения, целиком расположенные ВЫШЕ точки
  // вставки. Ломаем/восстанавливаем только затронутые (r2 > afterRow) — результат
  // идентичен полному перебору, но работы заметно меньше (шапка карты выше зоны вставки).
  const merges = allMerges.filter(m => m.r2 > afterRow);

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
  // (ищем среди ВСЕХ — строка-шаблон может быть выше afterRow и не попасть в merges)
  const srcRowMerges = allMerges.filter(m => m.r1 === srcRow && m.r2 === srcRow && m.c2 > m.c1);
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
  return buildMergeData_(sheet).map;
}

// Reads merged ranges ONCE and returns both:
//   map  — mergeMap[row1based][col1based] = {r, c} (top-left), для резолва записи в мёрджи;
//   list — [{r1, r2, c1, c2}] (1-based) — прямоугольники объединений, для детекта пересечений.
// Позволяет переиспользовать прочитанные объединения и не читать getMergedRanges дважды.
function buildMergeData_(sheet) {
  const map = {};
  const list = [];
  try {
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr < 1 || lc < 1) return { map, list };
    const merges = sheet.getRange(1, 1, lr, lc).getMergedRanges();
    for (const m of merges) {
      const r1 = m.getRow(), c1 = m.getColumn();
      const r2 = m.getLastRow(), c2 = m.getLastColumn();
      list.push({ r1, r2, c1, c2 });
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
  return { map, list };
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
