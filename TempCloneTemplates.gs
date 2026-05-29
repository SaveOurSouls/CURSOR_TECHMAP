// ============================================================
//  TempCloneTemplates.gs — ВРЕМЕННЫЙ СКРИПТ, удалить после использования.
//
//  Шаг 1: showCatalogMapping()      — логирует текущие источники и что будет создано
//  Шаг 2: cloneAllOperationTemplates() — клонирует шаблон для каждой операции из БД
//  Шаг 3: cleanupInsertionTemplates() — (из предыдущей версии) чистит некорректные
// ============================================================

// ── Шаг 1: просмотр маппинга ─────────────────────────────────
/**
 * Логирует: какие категории есть в каталоге, какие операции из БД к ним привяжутся,
 * и для каких операций источника нет. Запусти первым — проверь Logs перед клонированием.
 */
function showCatalogMapping() {
  const ss = SpreadsheetApp.getActive();
  const catalog = readCatalog_();

  // Группировка шаблонов по категории
  const catMap = buildCatMap_(catalog);
  Logger.log('=== КАТЕГОРИИ В КАТАЛОГЕ ===');
  Object.keys(catMap).sort().forEach(cat => {
    Logger.log(`  "${cat}" — ${catMap[cat].length} шаблон(а): ` +
      catMap[cat].map(t => t.title).join(', '));
  });

  // Чтение операций из снимка БД
  const ops = readAllOps_(ss);
  if (!ops) return;

  Logger.log('\n=== МАППИНГ ОПЕРАЦИЙ → ИСТОЧНИК ===');
  const noSource = [];
  ops.forEach(({ opNumber, opName }) => {
    const { variants, found } = resolveVariants_(opName, catMap);
    if (!found) {
      noSource.push(`  ${opNumber} | ${opName}`);
    } else {
      variants.forEach(v => {
        const title = (opNumber ? opNumber + ' | ' : '') + opName + v.suffix;
        Logger.log(`  [OK] "${title}"  ← источник: "${v.src.title}"`);
      });
    }
  });

  if (noSource.length) {
    Logger.log('\n=== НЕТ ИСТОЧНИКА (' + noSource.length + ') ===');
    noSource.forEach(s => Logger.log(s));
  } else {
    Logger.log('\nВсе операции покрыты источниками.');
  }
}

// ── Шаг 2: клонирование всех операций ───────────────────────
/**
 * Для каждой операции из _TC_TECHOPS_DB создаёт шаблон, клонируя соответствующий
 * источник из _TC_LIBRARY по категории. Уже существующие по title пропускает.
 */
function cloneAllOperationTemplates() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();

  const catalog = readCatalog_();
  const catMap  = buildCatMap_(catalog);

  const ops = readAllOps_(ss);
  if (!ops) return;

  const storeSheet   = ensureStoreSheet_(ss);
  const catalogSheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  const now = new Date().toISOString();

  let nextStoreRow = catalog.reduce((max, item) =>
    Math.max(max, item.storeRow + item.height - 1), 0) + 1;

  const existingTitles = new Set(catalog.map(t => t.title.toLowerCase().trim()));
  let created = 0, skipped = 0;
  const noSourceList = [];

  ops.forEach(({ opNumber, opName }) => {
    const { variants, found } = resolveVariants_(opName, catMap);

    if (!found) {
      noSourceList.push(opNumber + ' | ' + opName);
      return;
    }

    variants.forEach(({ suffix, src, category }) => {
      const title = (opNumber ? opNumber + ' | ' : '') + opName + suffix;

      if (existingTitles.has(title.toLowerCase())) {
        skipped++;
        return;
      }

      runWithSheetVisible_(storeSheet, () => {
        const srcRange = storeSheet.getRange(src.storeRow, src.storeColumn, src.height, src.width);
        ensureSheetCapacity_(storeSheet, nextStoreRow + src.height - 1, src.width);
        const dstRange = storeSheet.getRange(nextStoreRow, 1, src.height, src.width);
        dstRange.breakApart();
        clearStoreSlotForWrite_(dstRange);
        SpreadsheetApp.flush();
        copyRangePreservingFormulas_(srcRange, dstRange);
        dstRange.getCell(1, 1).setNote('techmap-template-store');
      });

      upsertCatalogRecord_(catalogSheet, {
        id:               slugify_(title),
        title,
        category,
        description:      '',
        storeRow:         nextStoreRow,
        storeColumn:      1,
        height:           src.height,
        width:            src.width,
        sourceSheet:      src.sourceSheet,
        sourceRange:      src.sourceRange,
        updatedAt:        now,
        rowHeightsJson:   JSON.stringify(src.rowHeights || []),
        columnWidthsJson: JSON.stringify(src.columnWidths || []),
        imagesJson:       src.imagesJson || '[]',
      });

      existingTitles.add(title.toLowerCase());
      nextStoreRow += src.height;
      created++;
      Logger.log('Создан: ' + title);
      SpreadsheetApp.flush();
    });
  });

  bumpCatalogVersion_();

  const noSrcText = noSourceList.length
    ? '\n\nНет источника (' + noSourceList.length + '):\n' + noSourceList.join('\n')
    : '\n\nВсе операции покрыты.';

  ui.alert(
    'Готово!\nСоздано: ' + created +
    '\nПропущено (уже есть): ' + skipped +
    noSrcText
  );
}

// ── Шаг 3: очистка некорректных шаблонов (монтаж) ───────────
/**
 * Удаляет "монтаж терминалов"-шаблоны без суффикса (А)/(В) в заголовке.
 * У правильных исправляет категорию и очищает описание.
 */
function cleanupInsertionTemplates() {
  const ss           = SpreadsheetApp.getActive();
  const ui           = SpreadsheetApp.getUi();
  const catalogSheet = ss.getSheetByName(TECHMAP_APP.librarySheetName);
  if (!catalogSheet) { ui.alert('_TC_LIBRARY не найден.'); return; }

  const lastRow = catalogSheet.getLastRow();
  if (lastRow < 2) { ui.alert('Каталог пуст.'); return; }

  const values = catalogSheet
    .getRange(2, 1, lastRow - 1, TECHMAP_APP.catalogHeaders.length)
    .getValues();

  const toDelete = [];
  let fixed = 0;

  values.forEach((row, i) => {
    const title       = String(row[1] || '');
    const description = String(row[3] || '');
    const sheetRow    = i + 2;
    if (!/монтаж терминал/i.test(title)) return;
    if (!/клон/i.test(description)) return;
    const hasAB = /\([АВ]\)/u.test(title);
    if (!hasAB) {
      toDelete.push(sheetRow);
    } else {
      const correctCat = /\(В\)/u.test(title) ? 'Монтаж терминалов (В)' : 'Монтаж терминалов (А)';
      catalogSheet.getRange(sheetRow, 3).setValue(correctCat);
      catalogSheet.getRange(sheetRow, 4).setValue('');
      fixed++;
    }
  });

  toDelete.sort((a, b) => b - a);
  toDelete.forEach(row => catalogSheet.deleteRow(row));
  bumpCatalogVersion_();
  ui.alert('Удалено: ' + toDelete.length + '\nИсправлено: ' + fixed);
}

// ── Helpers ───────────────────────────────────────────────────

/**
 * Читает все op-записи из _TC_TECHOPS_DB, дедублирует по opNumber|opName.
 * Возвращает [{opNumber, opName}] или null при ошибке.
 */
function readAllOps_(ss) {
  const ui = SpreadsheetApp.getUi();
  const dbSheet = ss.getSheetByName(TECHOPS_DB_APP.dataSheetName);
  if (!dbSheet) {
    ui.alert('_TC_TECHOPS_DB не найден. Синхронизируйте базу и повторите.');
    return null;
  }
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) { ui.alert('База техопераций пуста.'); return null; }

  const rows = dbSheet.getRange(2, 1, lastRow - 1, TECHOPS_DB_APP.dataHeaders.length).getValues();
  const seen = {};
  rows
    .filter(r => r[0] === 'op' && String(r[6] || '').trim())
    .forEach(r => {
      const opNumber = String(r[6] || '').trim();
      const opName   = String(r[7] || '').trim();
      const key = opNumber + '|' + opName;
      if (!seen[key]) seen[key] = { opNumber, opName };
    });
  return Object.values(seen);
}

/**
 * Строит map категория → [шаблоны].
 * Для каждой категории в качестве источника предпочитает шаблон без " | " в title
 * (базовый), иначе берёт первый попавшийся.
 */
function buildCatMap_(catalog) {
  const map = {};
  catalog.forEach(t => {
    const cat = (t.category || '').trim();
    if (!cat) return;
    if (!map[cat]) map[cat] = [];
    map[cat].push(t);
  });
  return map;
}

/**
 * Для opName определяет список вариантов для создания ({suffix, src, category}).
 * - Если в каталоге есть категория "opName (А)" или "opName (В)" → создаём пару А+В.
 * - Если есть "opName" → создаём один шаблон.
 * - Иначе → {found: false}.
 */
function resolveVariants_(opName, catMap) {
  const catA    = opName + ' (А)';
  const catB    = opName + ' (В)';
  const hasA    = !!(catMap[catA] && catMap[catA].length);
  const hasB    = !!(catMap[catB] && catMap[catB].length);
  const hasExact = !!(catMap[opName] && catMap[opName].length);

  if (hasA || hasB) {
    const srcA = pickSource_(catMap[catA]) || pickSource_(catMap[catB]);
    const srcB = pickSource_(catMap[catB]) || pickSource_(catMap[catA]);
    return {
      found: true,
      variants: [
        { suffix: ' (А)', src: srcA, category: catA },
        { suffix: ' (В)', src: srcB, category: catB },
      ],
    };
  }

  if (hasExact) {
    return {
      found: true,
      variants: [{ suffix: '', src: pickSource_(catMap[opName]), category: opName }],
    };
  }

  return { found: false, variants: [] };
}

/**
 * Выбирает "лучший" источник из списка шаблонов одной категории:
 * предпочитает тот, у которого нет " | " в title (базовый без кода операции).
 */
function pickSource_(templates) {
  if (!templates || !templates.length) return null;
  return templates.find(t => t.title.indexOf(' | ') === -1) || templates[0];
}
