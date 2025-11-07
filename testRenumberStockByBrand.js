const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows || 1;
    this.numCols = numCols || 1;
  }

  _ensureRow(rowIdx) {
    while (this.sheet.rows.length <= rowIdx) {
      const newRow = new Array(this.sheet.getLastColumn()).fill('');
      this.sheet.rows.push(newRow);
    }
  }

  getValues() {
    const values = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.row + r;
      if (rowIdx === 1) {
        const slice = [];
        for (let c = 0; c < this.numCols; c++) {
          const headerIdx = this.col + c - 1;
          slice.push(this.sheet.header[headerIdx] || '');
        }
        values.push(slice);
      } else {
        const dataIdx = rowIdx - 2; // data rows start at 2
        const dataRow = this.sheet.rows[dataIdx] || [];
        const slice = [];
        for (let c = 0; c < this.numCols; c++) {
          const colIdx = this.col + c - 1;
          slice.push(colIdx < dataRow.length ? dataRow[colIdx] : '');
        }
        values.push(slice);
      }
    }
    return values;
  }

  getValue() {
    const values = this.getValues();
    return values.length && values[0].length ? values[0][0] : '';
  }

  getDisplayValue() {
    const value = this.getValue();
    return value === null || value === undefined ? '' : String(value);
  }

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      const rowIdx = this.row + r;
      if (rowIdx === 1) {
        for (let c = 0; c < values[r].length; c++) {
          const colIdx = this.col + c - 1;
          if (colIdx >= this.sheet.header.length) {
            this.sheet.header.length = colIdx + 1;
          }
          this.sheet.header[colIdx] = values[r][c];
        }
        continue;
      }

      const targetIdx = rowIdx - 2;
      while (this.sheet.rows.length <= targetIdx) {
        this.sheet.rows.push([]);
      }
      for (let c = 0; c < values[r].length; c++) {
        const colIdx = this.col + c - 1;
        if (colIdx >= this.sheet.header.length) {
          this.sheet.header.length = colIdx + 1;
        }
        this.sheet.rows[targetIdx][colIdx] = values[r][c];
      }
    }
    return this;
  }

  setValue(value) {
    return this.setValues([[value]]);
  }

  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.row + r;
      const targetIdx = rowIdx - 2;
      if (!this.sheet.rows[targetIdx]) {
        continue;
      }
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.col + c - 1;
        this.sheet.rows[targetIdx][colIdx] = '';
      }
    }
    return this;
  }

  setBackground() {
    return this;
  }

  setFontColor() {
    return this;
  }

  setNumberFormat() {
    return this;
  }

  clearDataValidations() {
    return this;
  }

  sort() {
    return this;
  }

  getRow() {
    return this.row;
  }

  getColumn() {
    return this.col;
  }
}

class MockSheet {
  constructor(name, header, rows) {
    this.name = name;
    this.header = header.slice();
    this.rows = rows.map(r => r.slice());
  }

  getName() {
    return this.name;
  }

  getSheetId() {
    return 123;
  }

  getLastRow() {
    return this.rows.length + 1;
  }

  getLastColumn() {
    return this.header.length;
  }

  getRange(row, col, numRows, numCols) {
    return new MockRange(this, row, col, numRows, numCols);
  }

  deleteRow(rowNumber) {
    const idx = rowNumber - 2;
    if (idx >= 0 && idx < this.rows.length) {
      this.rows.splice(idx, 1);
    }
  }
}

let stockSheet = null;
let achatsSheet = null;

const spreadsheet = {
  getSheetByName(name) {
    if (name === 'Stock') return stockSheet;
    if (name === 'Achats') return achatsSheet;
    return null;
  }
};

let activeSpreadsheet = spreadsheet;

const sandbox = {
  console,
  SpreadsheetApp: {
    getActive() {
      return activeSpreadsheet;
    }
  },
  PropertiesService: {
    getDocumentProperties() {
      return {
        setProperty() {},
        getProperty() { return null; },
        deleteProperty() {}
      };
    }
  },
  Session: {
    getScriptTimeZone() {
      return 'Etc/GMT';
    }
  },
  Utilities: {
    formatDate(date) {
      return date.toISOString();
    }
  }
};

const code = fs.readFileSync('onEdit_Main.gs', 'utf8');
vm.runInNewContext(code, sandbox);

const HEADERS = vm.runInNewContext('HEADERS', sandbox);

stockSheet = new MockSheet('Stock', [
  HEADERS.STOCK.ID,
  HEADERS.STOCK.OLD_SKU,
  HEADERS.STOCK.SKU
], [
  ['ID-1', 'PCF-1', 'PCF-1'],
  ['ID-2', 'PCF-2', 'PCF-0'],
  ['ID-3', 'PCF-5', 'PCF-0'],
  ['ID-4', '',      'PCF-0']
]);

achatsSheet = new MockSheet('Achats', [
  HEADERS.ACHATS.ID,
  HEADERS.ACHATS.REFERENCE
], [
  ['ID-1', 'PCF'],
  ['ID-2', 'PCF'],
  ['ID-3', 'PCF'],
  ['ID-4', 'PCF']
]);

console.log('Initial old SKUs:', stockSheet.rows.map(row => row[1]));

sandbox.renumberStockByBrand_();

const newValues = stockSheet.rows.map(row => row[2]);
console.log('Renumbered SKUs:', newValues);

const suffixes = newValues.filter(Boolean).map(v => parseInt(v.split('-').pop(), 10));
const isStrictlyIncreasing = suffixes.every((val, idx) => idx === 0 || val > suffixes[idx - 1]);
console.log('Strictly increasing:', isStrictlyIncreasing);

// Test: existing numbered SKUs should remain untouched and new entries pick the next suffix.
const stockSheetExisting = new MockSheet('Stock', [
  'ID',
  'SKU(ancienne nomenclature)',
  'SKU'
], [
  ['ID-10', '', 'JLF-54'],
  ['ID-11', '', 'JLF-55'],
  ['ID-12', '', 'JLF-0']
]);

const achatsSheetExisting = new MockSheet('Achats', [
  'ID',
  'REFERENCE'
], [
  ['ID-10', 'JLF'],
  ['ID-11', 'JLF'],
  ['ID-12', 'JLF']
]);

const spreadsheetExisting = {
  getSheetByName(name) {
    if (name === 'Stock') return stockSheetExisting;
    if (name === 'Achats') return achatsSheetExisting;
    return null;
  }
};

activeSpreadsheet = spreadsheetExisting;

sandbox.renumberStockByBrand_();

const preservedValues = stockSheetExisting.rows.map(row => row[2]);
console.log('Existing numbered SKUs scenario:', preservedValues);
assert.deepStrictEqual(preservedValues, ['JLF-54', 'JLF-55', 'JLF-56']);
console.log('Existing SKUs preserved and new suffix assigned correctly.');

// --- handleAchats FRAIS DE COLISSAGE → TOTAL TTC recalculation ---
(() => {
  const headers = [
    HEADERS.ACHATS.ID,
    HEADERS.ACHATS.REFERENCE,
    HEADERS.ACHATS.QUANTITE_RECUE,
    HEADERS.ACHATS.PRIX_UNITAIRE_TTC,
    HEADERS.ACHATS.TOTAL_TTC,
    HEADERS.ACHATS.FRAIS_COLISSAGE
  ];

  const rows = [[
    'ID-20',
    'PCF',
    2,
    50,
    100,
    5
  ]];

  const sheet = new MockSheet('Achats', headers, rows);
  const fraisCol = headers.indexOf(HEADERS.ACHATS.FRAIS_COLISSAGE) + 1;
  const totalCol = headers.indexOf(HEADERS.ACHATS.TOTAL_TTC);
  const prixCol = headers.indexOf(HEADERS.ACHATS.PRIX_UNITAIRE_TTC);

  sheet.rows[0][fraisCol - 1] = 10;

  const event = {
    source: {
      getActiveSheet() { return sheet; },
      getSheetByName(name) {
        if (name === 'Achats') return sheet;
        return null;
      }
    },
    range: sheet.getRange(2, fraisCol),
    value: '10',
    oldValue: '5'
  };

  sandbox.handleAchats(event);

  assert.strictEqual(sheet.rows[0][totalCol], 110);
  assert.strictEqual(sheet.rows[0][prixCol], 50);
  console.log('handleAchats recalculates TOTAL TTC when frais increase.');
})();

(() => {
  const headers = [
    HEADERS.ACHATS.ID,
    HEADERS.ACHATS.REFERENCE,
    HEADERS.ACHATS.QUANTITE_RECUE,
    HEADERS.ACHATS.PRIX_UNITAIRE_TTC,
    HEADERS.ACHATS.TOTAL_TTC,
    HEADERS.ACHATS.FRAIS_COLISSAGE
  ];

  const rows = [[
    'ID-21',
    'PCF',
    2,
    50,
    110,
    10
  ]];

  const sheet = new MockSheet('Achats', headers, rows);
  const fraisCol = headers.indexOf(HEADERS.ACHATS.FRAIS_COLISSAGE) + 1;
  const totalCol = headers.indexOf(HEADERS.ACHATS.TOTAL_TTC);

  sheet.rows[0][fraisCol - 1] = '';

  const event = {
    source: {
      getActiveSheet() { return sheet; },
      getSheetByName(name) {
        if (name === 'Achats') return sheet;
        return null;
      }
    },
    range: sheet.getRange(2, fraisCol),
    value: '',
    oldValue: '10'
  };

  sandbox.handleAchats(event);

  assert.strictEqual(sheet.rows[0][totalCol], 100);
  console.log('handleAchats removes frais from TOTAL TTC when cleared.');
})();

(() => {
  const headers = [
    HEADERS.ACHATS.ID,
    HEADERS.ACHATS.REFERENCE,
    HEADERS.ACHATS.QUANTITE_RECUE,
    HEADERS.ACHATS.PRIX_UNITAIRE_TTC,
    HEADERS.ACHATS.TOTAL_TTC,
    HEADERS.ACHATS.FRAIS_COLISSAGE
  ];

  const rows = [[
    'ID-22',
    'PCF',
    '',
    '',
    150,
    20
  ]];

  const sheet = new MockSheet('Achats', headers, rows);
  const fraisCol = headers.indexOf(HEADERS.ACHATS.FRAIS_COLISSAGE) + 1;
  const totalCol = headers.indexOf(HEADERS.ACHATS.TOTAL_TTC);

  sheet.rows[0][totalCol] = 150;
  sheet.rows[0][fraisCol - 1] = 30;

  const event = {
    source: {
      getActiveSheet() { return sheet; },
      getSheetByName(name) {
        if (name === 'Achats') return sheet;
        return null;
      }
    },
    range: sheet.getRange(2, fraisCol),
    value: '30',
    oldValue: '20'
  };

  sandbox.handleAchats(event);

  assert.strictEqual(sheet.rows[0][totalCol], 160);
  console.log('handleAchats keeps existing total baseline when price/qty are missing.');
})();

(() => {
  const saleDate = vm.runInNewContext('new Date("2023-01-02T00:00:00Z")', sandbox);
  const stockHeaders = [
    HEADERS.STOCK.ID,
    HEADERS.STOCK.SKU,
    HEADERS.STOCK.TAILLE_COLIS,
    HEADERS.STOCK.LOT,
    HEADERS.STOCK.VENDU,
    HEADERS.STOCK.DATE_VENTE,
    HEADERS.STOCK.PRIX_VENTE,
    HEADERS.STOCK.VALIDER_SAISIE
  ];
  const stockRows = [[
    'ACHAT-1',
    'SKU-1',
    'M',
    'LOT1',
    true,
    saleDate,
    120,
    false
  ]];

  const achatsHeaders = [
    HEADERS.ACHATS.ID,
    HEADERS.ACHATS.QUANTITE_RECUE,
    HEADERS.ACHATS.PRIX_UNITAIRE_TTC,
    HEADERS.ACHATS.TOTAL_TTC,
    HEADERS.ACHATS.FRAIS_COLISSAGE
  ];
  const achatsRows = [[
    'ACHAT-1',
    2,
    100,
    200,
    0
  ]];

  const fraisHeaders = ['TAILLE DU COLIS', 'LOT', 'FRAIS DE COLISSAGE'];
  const fraisRows = [[
    'M',
    'LOT1',
    5
  ]];

  const ventesHeaders = [
    HEADERS.VENTES.ID,
    HEADERS.VENTES.DATE_VENTE,
    HEADERS.VENTES.ARTICLE,
    HEADERS.VENTES.SKU,
    HEADERS.VENTES.PRIX_VENTE,
    HEADERS.VENTES.DELAI_IMMOBILISATION,
    HEADERS.VENTES.DELAI_MISE_EN_LIGNE,
    HEADERS.VENTES.DELAI_PUBLICATION,
    HEADERS.VENTES.DELAI_VENTE,
    HEADERS.VENTES.FRAIS_COLISSAGE,
    HEADERS.VENTES.TAILLE_COLIS,
    HEADERS.VENTES.LOT
  ];

  const stock = new MockSheet('Stock', stockHeaders, stockRows);
  const achats = new MockSheet('Achats', achatsHeaders, achatsRows);
  const frais = new MockSheet('Frais', fraisHeaders, fraisRows);
  const ventes = new MockSheet('Ventes', ventesHeaders, []);

  const mockSpreadsheet = {
    sheets: { Stock: stock, Achats: achats, Frais: frais, Ventes: ventes },
    toastMessages: [],
    getSheetByName(name) {
      return this.sheets[name] || null;
    },
    insertSheet(name) {
      const sheet = new MockSheet(name, [], []);
      this.sheets[name] = sheet;
      return sheet;
    },
    toast(message) {
      this.toastMessages.push(message);
    },
    getActiveSheet() {
      return stock;
    },
    getUi() {
      return { alert() {}, ButtonSet: { OK: 'OK' } };
    }
  };

  activeSpreadsheet = mockSpreadsheet;

  const validerCol = stockHeaders.indexOf(HEADERS.STOCK.VALIDER_SAISIE) + 1;
  stock.rows[0][validerCol - 1] = true;

  const event = {
    source: mockSpreadsheet,
    range: stock.getRange(2, validerCol),
    value: 'TRUE',
    oldValue: 'FALSE'
  };

  const dateCol = stockHeaders.indexOf(HEADERS.STOCK.DATE_VENTE) + 1;
  const rawDate = stock.getRange(2, dateCol).getValue();
  assert.strictEqual(Object.prototype.toString.call(rawDate), '[object Date]');
  assert.strictEqual(Object.prototype.toString.call(sandbox.getDateOrNull_(rawDate)), '[object Date]');

  sandbox.handleStock(event);

  assert.strictEqual(mockSpreadsheet.toastMessages.length, 0);
  assert.strictEqual(stock.rows.length, 0);
  const ventesFraisCol = ventes.header.indexOf(HEADERS.VENTES.FRAIS_COLISSAGE);
  assert.strictEqual(ventes.rows[0][ventesFraisCol], 5);
  const ventesTailleCol = ventes.header.indexOf(HEADERS.VENTES.TAILLE_COLIS);
  assert.strictEqual(ventes.rows[0][ventesTailleCol], 'M');
  const ventesLotCol = ventes.header.indexOf(HEADERS.VENTES.LOT);
  assert.strictEqual(ventes.rows[0][ventesLotCol], 'LOT1');

  const totalCol = achatsHeaders.indexOf(HEADERS.ACHATS.TOTAL_TTC);
  assert.strictEqual(achats.rows[0][totalCol], 205);
  const fraisCol = achatsHeaders.indexOf(HEADERS.ACHATS.FRAIS_COLISSAGE);
  assert.strictEqual(achats.rows[0][fraisCol], 5);
  console.log('handleStock applies shipping fees to ventes and achats totals.');
})();

(() => {
  const saleDate = vm.runInNewContext('new Date("2023-01-03T00:00:00Z")', sandbox);
  const stockHeaders = [
    HEADERS.STOCK.ID,
    HEADERS.STOCK.SKU,
    HEADERS.STOCK.TAILLE_COLIS,
    HEADERS.STOCK.LOT,
    HEADERS.STOCK.VENDU,
    HEADERS.STOCK.DATE_VENTE,
    HEADERS.STOCK.PRIX_VENTE,
    HEADERS.STOCK.VALIDER_SAISIE
  ];
  const stockRows = [[
    'ACHAT-2',
    'SKU-2',
    '',
    '',
    true,
    saleDate,
    80,
    false
  ]];

  const achatsHeaders = [
    HEADERS.ACHATS.ID,
    HEADERS.ACHATS.QUANTITE_RECUE,
    HEADERS.ACHATS.PRIX_UNITAIRE_TTC,
    HEADERS.ACHATS.TOTAL_TTC,
    HEADERS.ACHATS.FRAIS_COLISSAGE
  ];
  const achatsRows = [[
    'ACHAT-2',
    2,
    80,
    160,
    0
  ]];

  const fraisHeaders = ['TAILLE DU COLIS', 'LOT', 'FRAIS DE COLISSAGE'];
  const fraisRows = [[
    'M',
    'LOT1',
    5
  ]];

  const ventesHeaders = [
    HEADERS.VENTES.ID,
    HEADERS.VENTES.DATE_VENTE,
    HEADERS.VENTES.ARTICLE,
    HEADERS.VENTES.SKU,
    HEADERS.VENTES.PRIX_VENTE,
    HEADERS.VENTES.DELAI_IMMOBILISATION,
    HEADERS.VENTES.DELAI_MISE_EN_LIGNE,
    HEADERS.VENTES.DELAI_PUBLICATION,
    HEADERS.VENTES.DELAI_VENTE,
    HEADERS.VENTES.FRAIS_COLISSAGE,
    HEADERS.VENTES.TAILLE_COLIS,
    HEADERS.VENTES.LOT
  ];

  const stock = new MockSheet('Stock', stockHeaders, stockRows);
  const achats = new MockSheet('Achats', achatsHeaders, achatsRows);
  const frais = new MockSheet('Frais', fraisHeaders, fraisRows);
  const ventes = new MockSheet('Ventes', ventesHeaders, []);

  const mockSpreadsheet = {
    sheets: { Stock: stock, Achats: achats, Frais: frais, Ventes: ventes },
    toastMessages: [],
    getSheetByName(name) {
      return this.sheets[name] || null;
    },
    insertSheet(name) {
      const sheet = new MockSheet(name, [], []);
      this.sheets[name] = sheet;
      return sheet;
    },
    toast(message) {
      this.toastMessages.push(message);
    },
    getActiveSheet() {
      return stock;
    },
    getUi() {
      return { alert() {}, ButtonSet: { OK: 'OK' } };
    }
  };

  activeSpreadsheet = mockSpreadsheet;

  const validerCol = stockHeaders.indexOf(HEADERS.STOCK.VALIDER_SAISIE) + 1;
  stock.rows[0][validerCol - 1] = true;

  const event = {
    source: mockSpreadsheet,
    range: stock.getRange(2, validerCol),
    value: 'TRUE',
    oldValue: 'FALSE'
  };

  sandbox.handleStock(event);

  assert.strictEqual(stock.rows.length, 1);
  assert.strictEqual(stock.rows[0][validerCol - 1], false);
  assert.strictEqual(ventes.rows.length, 0);
  const totalCol = achatsHeaders.indexOf(HEADERS.ACHATS.TOTAL_TTC);
  assert.strictEqual(achats.rows[0][totalCol], 160);
  const fraisCol = achatsHeaders.indexOf(HEADERS.ACHATS.FRAIS_COLISSAGE);
  assert.strictEqual(achats.rows[0][fraisCol], 0);
  console.log('handleStock blocks validation when shipping size is missing.');
})();
