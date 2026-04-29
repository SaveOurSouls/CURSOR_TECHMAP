/**
 * Явное описание структур данных проекта.
 *
 * Файл нужен для двух задач:
 * 1) держать схему в формате `.gs`, который можно напрямую перенести в Apps Script;
 * 2) иметь единый словарь полей для каталога, сохранения и вставки шаблонов.
 */
const TECHMAP_DATA_MODEL = {
  materialsSource: {
    spreadsheetId: '1NExDzeG-vw3zY_ooeXoxRffIARh2wbLJNAHT3FU_Ig8',
    sheets: ['COMPCON', 'COMPCOAX', 'COMPTERM', 'COMPWIRE', 'COMPACCESS'],
    searchTagHeader: 'Поисковый тег',
  },

  operationsSource: {
    spreadsheetId: '1W3VK9Fw71lYdw1Klcsn_za5-2EhvLoXIAKZVYOCnKcs',
    sheets: ['БД.ОБ', 'БД.ОП', 'БД.ТЕР', 'БД.КОАКС'],
  },

  materialRecord: {
    sourceSheet: 'COMPCON',
    searchTag: '123-ABC | Разъем SMA male | Rosenberger | ELITAN',
    article: '123-ABC',
    type: 'Разъем SMA male',
    manufacturer: 'Rosenberger',
    supplier: 'ELITAN',
    normalized: '123-abc|разъемsmamale|rosenberger|elitan',
  },

  operationRecord: {
    sourceSheet: 'БД.ОП',
    searchKey: '010.005 | Опрессовка наконечника',
    number: '010.005',
    title: 'Опрессовка наконечника',
    operationTime: '12',
    setupTimeSec: '90',
    setupConsumption: '1 шт',
    machineTime: '0',
    normalizedSearch: '010.005|опрессовканаконечника',
    payload: {
      numberAndTitle: '010.005 | Опрессовка наконечника',
      operationTime: '12',
      setupTimeSec: '90',
      setupConsumption: '1 шт',
      machineTime: '0',
    },
  },

  catalogRecord: {
    id: 'wire-cutting',
    title: 'Резка провода',
    category: 'COAX',
    description: 'Подготовка заготовок кабеля RG58 по длине.',
    storeRow: 1,
    storeColumn: 1,
    height: 32,
    width: 12,
    sourceSheet: 'Рабочий лист',
    sourceRange: 'A1:L32',
    updatedAt: '2026-04-28T22:25:00.000Z',
    rowHeightsJson: '[28,28,24]',
    columnWidthsJson: '[155,95,95]',
  },

  saveDialogState: {
    selection: {
      sheetName: 'Лист1',
      rangeA1: 'A1:L32',
      height: 32,
      width: 12,
    },
    templates: [
      {
        id: 'wire-cutting',
        title: 'Резка провода',
        category: 'COAX',
        description: 'Подготовка заготовок кабеля RG58 по длине.',
        sizeLabel: '32 x 12',
      },
    ],
  },

  sidebarCatalogItem: {
    id: 'wire-cutting',
    title: 'Резка провода',
    category: 'COAX',
    description: 'Подготовка заготовок кабеля RG58 по длине.',
    sizeLabel: '32 x 12',
    updatedAt: '2026-04-28T22:25:00.000Z',
  },

  insertResult: {
    title: 'Резка провода',
    sheetName: 'Проект 001',
    insertedRange: 'B5:M36',
  },
};

/**
 * Возвращает короткую памятку по структурам данных проекта.
 * Можно вызывать из редактора Apps Script при внедрении или отладке.
 */
function getTechmapDataStructure() {
  return {
    description: 'Схемы данных каталога шаблонов и UI состояния',
    structures: TECHMAP_DATA_MODEL,
  };
}
