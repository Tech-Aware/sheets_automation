const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
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

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      const rowIdx = this.row + r;
      const targetIdx = rowIdx - 2;
      if (!this.sheet.rows[targetIdx]) {
        this.sheet.rows[targetIdx] = [];
      }
      for (let c = 0; c < values[r].length; c++) {
        const colIdx = this.col + c - 1;
        this.sheet.rows[targetIdx][colIdx] = values[r][c];
      }
    }
    return this;
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

[
  'config.gs',
  'sheet_utils.gs',
  'achats_workflow.gs',
  'stock_workflow.gs',
  'onEdit_Main.gs'
].forEach(path => {
  const source = fs.readFileSync(path, 'utf8');
  vm.runInNewContext(source, sandbox);
});

const HEADERS = vm.runInNewContext('HEADERS', sandbox);

stockSheet = new MockSheet('Stock', [
  HEADERS.STOCK.ID,
  HEADERS.STOCK.SKU
], [
  ['ID-1', 'PCF-1'],
  ['ID-2', 'PCF-0'],
  ['ID-3', 'PCF-0'],
  ['ID-4', 'PCF-0']
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

console.log('Initial SKUs:', stockSheet.rows.map(row => row[1]));

sandbox.renumberStockByBrand_();

const newValues = stockSheet.rows.map(row => row[1]);
console.log('Renumbered SKUs:', newValues);

const suffixes = newValues.filter(Boolean).map(v => parseInt(v.split('-').pop(), 10));
const isStrictlyIncreasing = suffixes.every((val, idx) => idx === 0 || val > suffixes[idx - 1]);
console.log('Strictly increasing:', isStrictlyIncreasing);

// Test: existing numbered SKUs should remain untouched and new entries pick the next suffix.
const stockSheetExisting = new MockSheet('Stock', [
  'ID',
  'SKU'
], [
  ['ID-10', 'JLF-54'],
  ['ID-11', 'JLF-55'],
  ['ID-12', 'JLF-0']
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

const preservedValues = stockSheetExisting.rows.map(row => row[1]);
console.log('Existing numbered SKUs scenario:', preservedValues);
assert.deepStrictEqual(preservedValues, ['JLF-54', 'JLF-55', 'JLF-56']);
console.log('Existing SKUs preserved and new suffix assigned correctly.');
