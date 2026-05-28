// ============================================================
//  Config.gs — все константы проекта (единственный источник истины)
// ============================================================

// ── Модуль шаблонов ──────────────────────────────────────────
const TECHMAP_APP = {
  menuTitle:            'Техкарты',
  librarySheetName:     '_TC_LIBRARY',
  storeSheetName:       '_TC_STORE',
  imageCacheFolderName: '_TC_IMAGE_CACHE',
  legacyTemplatePrefix: '_TPL_',
  catalogHeaders: [
    'id', 'title', 'category', 'description',
    'storeRow', 'storeColumn', 'height', 'width',
    'sourceSheet', 'sourceRange', 'updatedAt',
    'rowHeightsJson', 'columnWidthsJson', 'imagesJson',
  ],
};

// ── База техопераций ─────────────────────────────────────────
const TECHOPS_DB_APP = {
  sourceSpreadsheetId: '1W3VK9Fw71lYdw1Klcsn_za5-2EhvLoXIAKZVYOCnKcs',
  metaSheetName:       '_TC_TECHOPS_META',
  dataSheetName:       '_TC_TECHOPS_DB',
  metaHeaders:  ['key', 'value'],
  dataHeaders: [
    'tabKey', 'displayText', 'normalizedSearch', 'exportJson',
    'sourceSheet', 'sortKey',
    'extra1', 'extra2', 'extra3', 'extra4', 'extra5', 'extra6', 'extra7',
  ],
  cacheKeyPrefix:   'techmap-techops-db-v8',
  schemaVersion:    8,
  cacheChunkSize:   80000,
  cacheTtlSeconds:  21600,
  tabs: {
    ob: {
      key:               'ob',
      label:             'БД.ОБ',
      sourceSheetName:   'БД.ОБ',
      headerRowNumber:   2,
      searchPlaceholder: 'Поиск по полю "Для базы"...',
      outputLabels:      ['Для базы'],
    },
    op: {
      key:               'op',
      label:             'БД.ОП',
      sourceSheetName:   'БД.ОП',
      headerRowNumber:   2,
      searchPlaceholder: 'Поиск по номеру или названию операции...',
      outputLabels: [
        'Номер | Название',
        'Время Операции',
        'Время подготовки, сек',
        'Расход на настройку м; шт;',
        'Время машины, сек/оп; сек/м',
      ],
    },
    ter: {
      key:               'ter',
      label:             'БД.ТЕР',
      sourceSheetName:   'БД.ТЕР',
      headerRowNumber:   1,
      searchPlaceholder: 'Поиск по производителю, серии, product name...',
      outputLabels: [
        'Тип', 'Производитель', 'Product Name', 'Series',
        'Шаг', 'Тип конт.', 'Арт. ISL', 'Арт. SAG', 'Аппликатор',
      ],
    },
    coax: {
      key:               'coax',
      label:             'БД.КОАКС',
      sourceSheetName:   'БД.КОАКС',
      headerRowNumber:   2,
      searchPlaceholder: 'Поиск по артикулам, сериям, проводу, размерам...',
      outputLabels: ['Артикул', 'Программа', 'D1', 'D2', 'D3', 'L1', 'L2', 'L3', 'L+', 'L-'],
    },
  },
  tabOrder: ['ob', 'op', 'ter', 'coax'],
};

// ── Генератор сборок ──────────────────────────────────────────
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
    tolerance:    '{{ДОПУСК}}',
    lengthKd:     '[L КД]',
  },
  // Linear tolerance: at 1000mm → ±toleranceMmPerM mm; scales proportionally
  toleranceMmPerM: 8,
};
