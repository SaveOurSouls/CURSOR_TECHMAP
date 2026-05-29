// ============================================================
//  Utils.gs — shared utilities (единый источник для всех модулей)
// ============================================================

// ── String helpers ───────────────────────────────────────────

/** Приводит значение к строке и обрезает пробелы. */
function normalizeString_(value) {
  return String(value || '').trim();
}

/**
 * Нормализует заголовок колонки: lowercase, схлопывание пробелов.
 * Заменяет normalizeTechOperationsHeader_ и normalizeMaterialHeader_ (идентичная логика).
 */
function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/** Нормализует строку для полнотекстового поиска: lowercase, без пробелов. */
function normalizeSearch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, '');
}

/** Преобразует значение в целое число; возвращает 0 при ошибке. */
function toInt_(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
}

/** Конвертирует секунды в минуты; возвращает '' если значение равно нулю. */
function secToMin_(value) {
  const s = parseFloat(String(value || '').replace(',', '.'));
  if (!s) return '';
  const m = s / 60;
  return String(m % 1 === 0 ? m : m.toFixed(2)).replace('.', ',');
}

/** Парсит JSON-массив; возвращает [] при ошибке или пустом значении. */
function parseJsonArray_(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

/** Создаёт slug из произвольной строки (для идентификаторов шаблонов). */
function slugify_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9а-яё]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-+/g, '-');
}

// ── Sheet helpers ────────────────────────────────────────────

/**
 * Расширяет лист до нужного числа строк/колонок если текущего размера недостаточно.
 */
function ensureSheetCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

/**
 * Создаёт лист с уникальным именем на основе baseName.
 * Если имя занято — добавляет суффикс -2, -3, … до свободного.
 */
function createUniqueSheet_(ss, baseName) {
  const MAX_LEN = 100;
  const base = String(baseName || 'Лист').substring(0, MAX_LEN);
  if (!ss.getSheetByName(base)) return ss.insertSheet(base);
  for (let i = 2; i <= 999; i++) {
    const suffix = '-' + i;
    const candidate = base.substring(0, MAX_LEN - suffix.length) + suffix;
    if (!ss.getSheetByName(candidate)) return ss.insertSheet(candidate);
  }
  return ss.insertSheet();
}

/** Возвращает true если имя листа является служебным (_TC_* или _TPL_*). */
function isSystemSheet_(sheetName) {
  return (
    sheetName === TECHMAP_APP.librarySheetName ||
    sheetName === TECHMAP_APP.storeSheetName   ||
    sheetName.indexOf(TECHMAP_APP.legacyTemplatePrefix) === 0
  );
}

/**
 * Временно показывает скрытый лист на время выполнения callback fn.
 * Служба Sheets возвращает ошибку для ряда операций (clear, copyTo)
 * на полностью скрытом листе (_TC_STORE).
 */
function runWithSheetVisible_(sheet, fn) {
  if (!sheet) return fn();
  const wasHidden = sheet.isSheetHidden();
  if (wasHidden) sheet.showSheet();
  try {
    return fn();
  } finally {
    if (wasHidden) sheet.hideSheet();
  }
}

// ── Catalog version ───────────────────────────────────────────

/** Обновляет метку версии каталога (timestamp) в UserProperties. */
function bumpCatalogVersion_() {
  try {
    PropertiesService.getUserProperties()
      .setProperty('tc_catalog_version', Date.now().toString());
  } catch (e) {}
}

// ── ChunkCache_ — универсальная chunk-кеш утилита ─────────────
/**
 * Возвращает объект с методами load/save/clear для хранения
 * произвольного JSON-объекта в CacheService чанками.
 *
 * @param {string} prefix    — уникальный префикс ключа кеша для модуля
 * @param {number} chunkSize — максимальный размер одного чанка в символах
 * @param {number} ttl       — TTL в секундах
 */
function ChunkCache_(prefix, chunkSize, ttl) {
  const cache      = CacheService.getDocumentCache();
  const countKey   = prefix + ':count';
  const chunkKey   = (i) => prefix + ':chunk:' + i;

  return {
    /** Загружает и парсит объект из кеша. Возвращает null при промахе или ошибке. */
    load() {
      const countVal = cache.get(countKey);
      const count = toInt_(countVal);
      if (!count) return null;
      let serialized = '';
      for (let i = 0; i < count; i++) {
        const chunk = cache.get(chunkKey(i));
        if (chunk === null || chunk === undefined) return null;
        serialized += chunk;
      }
      try { return JSON.parse(serialized); } catch (e) { return null; }
    },

    /** Сериализует объект и записывает в кеш чанками. */
    save(data) {
      const serialized  = JSON.stringify(data);
      const chunkCount  = Math.ceil(serialized.length / chunkSize) || 1;
      cache.put(countKey, String(chunkCount), ttl);
      for (let i = 0; i < chunkCount; i++) {
        cache.put(
          chunkKey(i),
          serialized.slice(i * chunkSize, (i + 1) * chunkSize),
          ttl
        );
      }
    },

    /** Удаляет все чанки и счётчик из кеша. */
    clear() {
      const countVal = cache.get(countKey);
      const count = toInt_(countVal) || 0;
      cache.remove(countKey);
      for (let i = 0; i < count; i++) cache.remove(chunkKey(i));
    },
  };
}
