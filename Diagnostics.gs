// ============================================================
//  Diagnostics.gs — диагностика и живые тесты через `clasp run`
//  Dev-тулинг (не UI). Все функции возвращают строку/JSON для clasp run.
//  Запуск: `clasp run <fn>` из _local; живые тесты — `npm run check:live`.
//  Зависимости: Config/Utils/TemplateStore/OperationDatabase/AssemblyGenerator.
// ============================================================

// Живой смоук-набор против РЕАЛЬНОЙ таблицы — то, что vm-тесты (чистая логика) не ловят
// (напр. лимит insertImage 2МБ/картинки). Возвращает строки PASS/FAIL/ERR + сводку.
function liveTestSuite() {
  var results = [];
  function check(name, fn) {
    try { var r = fn() || {}; results.push((r.ok ? 'PASS' : 'FAIL') + ' | ' + name + (r.msg ? ' — ' + r.msg : '')); }
    catch (e) { results.push('ERR  | ' + name + ' — ' + (e && e.message)); }
  }

  check('каталог шаблонов не пуст', function () {
    var n = readCatalog_().length; return { ok: n > 0, msg: n + ' шаблонов' };
  });
  check('snapshot: схема актуальна', function () {
    var s = getTechOperationsSnapshot_(); var v = s.meta && s.meta.schemaVersion;
    return { ok: String(v) === String(TECHOPS_DB_APP.schemaVersion), msg: 'snapshot v' + v + ' / code v' + TECHOPS_DB_APP.schemaVersion };
  });
  check('БД.ОП: записи есть', function () { var n = readOpRecordsForGenerator_().length; return { ok: n > 0, msg: n + ' оп' }; });
  check('БД.ТЕР: записи есть', function () { var n = readTerRecordsForGenerator_().length; return { ok: n > 0, msg: n + ' тер' }; });
  check('картинки шаблонов 100% в Drive', function () {
    var tot = 0, drive = 0;
    readCatalog_().forEach(function (t) { parseJsonArray_(t.imagesJson).forEach(function (i) { tot++; if (i.driveFileId) drive++; }); });
    return { ok: tot === drive, msg: drive + '/' + tot + ' в Drive' };
  });
  check('вставка шаблона с картинкой кладёт картинку', function () {
    var t = readCatalog_().find(function (x) { return parseJsonArray_(x.imagesJson).some(function (i) { return i.driveFileId; }); });
    if (!t) return { ok: false, msg: 'нет шаблона с Drive-картинкой' };
    var ss = SpreadsheetApp.getActive(); var sheet = null;
    try {
      var res = insertTemplate(t.id); sheet = ss.getSheetByName(res.sheetName); SpreadsheetApp.flush();
      var n = sheet ? sheet.getImages().length : 0;
      return { ok: n >= 1, msg: n + ' картинок на «' + t.title + '»' };
    } finally { try { if (sheet) ss.deleteSheet(sheet); } catch (e) {} }
  });

  var fails = results.filter(function (r) { return r.indexOf('FAIL') === 0 || r.indexOf('ERR') === 0; }).length;
  var out = 'LIVE ' + (fails ? 'FAIL(' + fails + ')' : 'OK') + ' [' + results.length + ' проверок]\n' + results.join('\n');
  Logger.log(out);
  return out;
}

// ── Точечная диагностика (clasp run <fn>) ────────────────────

/** Список шаблонов: размер, число картинок (и сколько в Drive). */
function diagTemplates() {
  var lines = ['ШАБЛОНЫ (' + readCatalog_().length + '):'];
  readCatalog_().forEach(function (t) {
    var imgs = parseJsonArray_(t.imagesJson);
    lines.push('• ' + t.title + ' | ' + t.height + 'x' + t.width + ' | картинок=' + imgs.length
      + ' (Drive ' + imgs.filter(function (i) { return i.driveFileId; }).length + ')');
  });
  var out = lines.join('\n'); Logger.log(out); return out;
}

/** Покрытие Drive-кешем по шаблонам с картинками. */
function diagImages() {
  var lines = ['КАРТИНКИ:']; var tot = 0, drive = 0;
  readCatalog_().forEach(function (t) {
    var imgs = parseJsonArray_(t.imagesJson); if (!imgs.length) return;
    var d = imgs.filter(function (i) { return i.driveFileId; }).length; tot += imgs.length; drive += d;
    lines.push('• ' + t.title + ' — ' + imgs.length + ' (Drive ' + d + ')');
  });
  lines.push('ИТОГ: ' + drive + '/' + tot + ' в Drive');
  var out = lines.join('\n'); Logger.log(out); return out;
}

/** Что генератор находит как исходную таблицу (BOM) на активном листе. */
function diagBom() {
  var src = findAssemblySourceData_(SpreadsheetApp.getActive());
  var out = 'BOM: лист=«' + (src.sourceSheet || '?') + '» компонентов=' + (src.components || []).length
    + ' полей изделия=' + Object.keys(src.assemblyInfo || {}).length;
  Logger.log(out); return out;
}

/** Состояние БД техопераций: версии схемы (snapshot vs code), счётчики по вкладкам. */
function diagDbSchema() {
  var s = getTechOperationsSnapshot_(); var c = (s.meta && s.meta.countsByTab) || {};
  var out = 'БД: snapshot v' + (s.meta && s.meta.schemaVersion) + ' / code v' + TECHOPS_DB_APP.schemaVersion
    + ' | записей=' + (s.records || []).length
    + ' | ' + TECHOPS_DB_APP.tabOrder.map(function (k) { return k + ':' + (c[k] || 0); }).join(' ');
  Logger.log(out); return out;
}

// Сэмпл живых строк БД.ОП/БД.ТЕР в формате fixtures/real-db.json — для обновления фикстуры
// тестов (`npm run fixture` пишет в real-db.live.json, дальше сверить/влить вручную).
function dumpDbFixture() {
  var recs = (getTechOperationsSnapshot_().records) || [];
  var op = recs.filter(function (r) { return r.tabKey === 'op'; }).slice(0, 4).map(function (r) {
    return { opNumber: r.opNumber || '', opName: r.opName || '', tOp: r.tOp || '', tPrep: r.tPrep || '', tMachine: r.tMachine || '' };
  });
  var ter = recs.filter(function (r) { return r.tabKey === 'ter'; }).slice(0, 4).map(function (r) {
    return {
      terManufacturer: r.terManufacturer || '', terSeries: r.terSeries || '', terComponent: r.terComponent || '',
      terType: r.terType || '', terArticle: r.terArticle || '', terLPlus: r.terLPlus || '', terLMinus: r.terLMinus || '',
      terApplicator: r.terApplicator || '', terCrimpHeight: r.terCrimpHeight || '',
      terPullForceMin: r.terPullForceMin || '', terPullForceMax: r.terPullForceMax || '',
      terStep: r.terStep || '', terStrip: r.terStrip || ''
    };
  });
  return JSON.stringify({ _note: 'Сэмпл из живой БД (dumpDbFixture). Сверить/влить в real-db.json.', op: op, ter: ter });
}
