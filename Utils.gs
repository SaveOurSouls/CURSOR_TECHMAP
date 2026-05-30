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

/**
 * Форматирует число для записи в техкарту: округляет до maxDecimals знаков,
 * срезает хвостовые нули и использует запятую как десятичный разделитель
 * (конвенция проекта). Единый источник форматирования дробных значений —
 * устраняет float-мусор вида "0,31100000000000005" в нормах расхода.
 *
 * @param {number|string} value      число или строка (с . или , разделителем)
 * @param {number}        [maxDecimals=2] максимум знаков после запятой
 * @returns {string} форматированная строка или '' для нечислового/пустого ввода
 */
function formatDecimalComma_(value, maxDecimals) {
  const n = typeof value === 'number'
    ? value
    : parseFloat(String(value === null || value === undefined ? '' : value).replace(',', '.'));
  if (!isFinite(n)) return '';
  const d = maxDecimals === undefined ? 2 : maxDecimals;
  return String(Number(n.toFixed(d))).replace('.', ',');
}

/** Конвертирует секунды в минуты; возвращает '' если значение равно нулю. */
function secToMin_(value) {
  const s = parseFloat(String(value || '').replace(',', '.'));
  if (!s) return '';
  const m = s / 60;
  return String(m % 1 === 0 ? m : m.toFixed(2)).replace('.', ',');
}

/**
 * Нормализует десятичный разделитель технического числа к запятой (конвенция
 * проекта для L+/L-/шага/обжима и т.п.). Чистая функция, не зависит от Sheets.
 *
 * Правило: если значение целиком парсится как число (опц. знак, цифры, один
 * разделитель . или ,) — заменяет точку на запятую. Любое нечисловое значение
 * (буквы, единицы измерения, диапазоны) возвращается без изменений — данные
 * не теряются и не искажаются.
 *
 * @param {*} value входное значение (строка/число/пусто)
 * @returns {string} нормализованная строка или '' для пустого ввода
 */
function normalizeTechnicalDecimal_(value) {
  const str = String(value === null || value === undefined ? '' : value).trim();
  if (str === '') return '';
  return /^[+-]?\d+([.,]\d+)?$/.test(str) ? str.replace('.', ',') : str;
}

/**
 * Сериализует объект в JSON для безопасного встраивания в <script> через
 * GAS-скриптлет <?!= ... ?> (неэкранированный вывод). Экранирует символы,
 * которыми пользовательские данные (напр. название шаблона) могли бы
 * вырваться из script-контекста или сломать JS-литерал:
 *   <  >  → \uXXXX (предотвращает </script> breakout — stored XSS),
 *   U+2028 / U+2029 → \uXXXX (иначе обрывают строковый литерал JS).
 *
 * @param {*} obj любой JSON-сериализуемый объект
 * @returns {string} безопасная для встраивания JSON-строка
 */
function embedJsonForHtml_(obj) {
  return JSON.stringify(obj).replace(/[<>\u2028\u2029]/g, function (ch) {
    return '\\u' + ch.charCodeAt(0).toString(16).padStart(4, '0');
  });
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

/**
 * Безопасно парсит JSON; возвращает fallback при пустом/битом значении.
 * Заменяет россыпь inline try/catch в путях загрузки снапшота.
 *
 * @param {*} value      сырая строка JSON
 * @param {*} fallback   значение по умолчанию при ошибке/пустом вводе
 * @returns {*} распарсенный объект или fallback
 */
function safeJsonParse_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  try {
    const parsed = JSON.parse(value);
    return parsed === null || parsed === undefined ? fallback : parsed;
  } catch (e) {
    return fallback;
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
 * Записывает строку заголовков в row 1 листа единым батч-вызовом
 * с фирменным оформлением (жирный + фон). Единый источник стиля шапки
 * для всех служебных листов.
 *
 * @param {Sheet}    sheet   целевой лист
 * @param {string[]} headers массив заголовков колонок
 * @param {string}   [background='#f3f6fc'] цвет фона шапки
 */
function writeSheetHeader_(sheet, headers, background) {
  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(background || '#f3f6fc');
}

/**
 * Применяет значения размеров (высоты строк / ширины колонок) к листу,
 * группируя подряд идущие одинаковые значения в один батч-вызов сеттера
 * (run-length encoding). Нулевые/пустые значения пропускаются.
 *
 * @param {number[]} values массив размеров (px), индекс 0 = start
 * @param {number}   start  1-based начальная строка/колонка
 * @param {function(number, number, number)} setter (startIndex, count, size)
 */
function applyRunLengthDimensions_(values, start, setter) {
  let i = 0;
  while (i < values.length) {
    const size = values[i];
    if (!size) { i += 1; continue; }
    let j = i + 1;
    while (j < values.length && values[j] === size) j += 1;
    setter(start + i, j - i, size);
    i = j;
  }
}

/**
 * Проверяет наличие колонки в карте заголовков. Корректно обрабатывает
 * индекс 0 (первая колонка) в отличие от простого truthy-чека.
 *
 * @param {Object} headerMap карта normalizedHeader → columnIndex
 * @param {string} key       нормализованный ключ заголовка
 * @returns {boolean}
 */
function hasHeader_(headerMap, key) {
  const idx = headerMap[key];
  return idx === 0 || idx > 0;
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
