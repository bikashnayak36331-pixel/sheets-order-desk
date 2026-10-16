'use strict';
/**
 * A small in-memory stand-in for the Google Apps Script services Order Desk uses.
 *
 * It runs the real src/*.js files in a Node `vm` context, the same way Apps Script runs them in
 * one shared global scope. Only the methods Order Desk calls are implemented, and an unknown
 * method throws, so a typo in the source fails the tests instead of failing in Google.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC_ORDER = ['Config.js', 'Core.js', 'Sheets.js', 'Setup.js', 'Orders.js', 'Documents.js', 'Reports.js'];

function colLetterToNumber(letters) {
  return letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

class FakeRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    if (row < 1 || col < 1 || numRows < 1 || numCols < 1) {
      throw new Error(`Range out of bounds: row ${row}, col ${col}, ${numRows}x${numCols}`);
    }
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const row = [];
      for (let c = 0; c < this.numCols; c++) row.push(this.sheet.cell(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  getValue() { return this.sheet.cell(this.row, this.col); }
  setValues(values) {
    if (values.length !== this.numRows || values.some((r) => r.length !== this.numCols)) {
      throw new Error(`The data has ${values.length} rows but the range has ${this.numRows}`);
    }
    values.forEach((r, i) => r.forEach((v, j) => this.sheet.setCell(this.row + i, this.col + j, v)));
    return this;
  }
  setValue(v) {
    for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet.setCell(this.row + r, this.col + c, v);
    return this;
  }
  setFormula(f) {
    this.sheet.formulas[`${this.row},${this.col}`] = f;
    this.sheet.setCell(this.row, this.col, f);
    return this;
  }
  getFormula() { return this.sheet.formulas[`${this.row},${this.col}`] || ''; }
  setDataValidation(rule) { this.sheet.validations.push({ range: this, rule }); return this; }
}
// Formatting calls that really exist on Apps Script's Range: recorded, otherwise no-ops.
['setNumberFormat', 'setFontWeight', 'setFontSize', 'setFontColor', 'setBackground', 'setHorizontalAlignment',
  'setWrap'].forEach((m) => {
  FakeRange.prototype[m] = function (...args) { this.sheet.formats.push([m, this.row, this.col, ...args]); return this; };
});

class FakeSheet {
  constructor(spreadsheet, name) {
    Object.assign(this, { spreadsheet, name, data: [], formulas: {}, validations: [], formats: [], charts: [],
      rules: [], frozen: 0, widths: {} });
  }
  getName() { return this.name; }
  getParent() { return this.spreadsheet; }
  getFormUrl() { return this.formUrl || null; }
  cell(r, c) { return (this.data[r - 1] && this.data[r - 1][c - 1] !== undefined) ? this.data[r - 1][c - 1] : ''; }
  setCell(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() {
    for (let r = this.data.length; r > 0; r--) if (this.data[r - 1].some((v) => v !== '' && v !== null)) return r;
    return 0;
  }
  getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxRows() { return 1000; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (!m) throw new Error('Bad A1 notation ' + a);
      const c1 = colLetterToNumber(m[1]); const r1 = Number(m[2]);
      const c2 = m[3] ? colLetterToNumber(m[3]) : c1; const r2 = m[4] ? Number(m[4]) : r1;
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(values) { const r = this.getLastRow() + 1; values.forEach((v, i) => this.setCell(r, i + 1, v)); return this; }
  deleteRows(start, count) { this.data.splice(start - 1, count); return this; }
  clearContents() { this.data = []; this.formulas = {}; return this; }
  setFrozenRows(n) { this.frozen = n; return this; }
  setColumnWidth(col, w) { this.widths[col] = w; return this; }
  setConditionalFormatRules(rules) { this.rules = rules; return this; }
  getCharts() { return this.charts; }
  insertChart(chart) { this.charts.push(chart); return this; }
  newChart() {
    const spec = { options: {}, ranges: [] };
    const builder = {
      setChartType(t) { spec.type = t; return builder; },
      addRange(r) { spec.ranges.push(r); return builder; },
      setPosition(...p) { spec.position = p; return builder; },
      setOption(k, v) { spec.options[k] = v; return builder; },
      build() { return spec; },
    };
    return builder;
  }
  /** Test helper: the sheet as objects keyed by the header row. */
  records() {
    const [header, ...rows] = this.data;
    return rows.filter((r) => r.some((v) => v !== '')).map((r) => Object.fromEntries(header.map((h, i) => [h, r[i] === undefined ? '' : r[i]])));
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = []; this.toasts = []; this.activeRange = null; }
  getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`A sheet named "${name}" already exists`);
    const s = new FakeSheet(this, name); this.sheets.push(s); return s;
  }
  getSheets() { return this.sheets.slice(); }
  getSpreadsheetTimeZone() { return 'UTC'; }
  getActiveRange() { return this.activeRange; }
  toast(msg, title) { this.toasts.push({ msg, title }); }
}

function builder(kind) {
  const spec = { kind };
  const proxy = new Proxy({}, {
    get(_, prop) {
      if (prop === 'build') return () => spec;
      const allowed = {
        validation: ['requireValueInList', 'setAllowInvalid', 'setHelpText'],
        rule: ['whenTextEqualTo', 'setBackground', 'setRanges'],
      }[kind];
      if (!allowed.includes(prop)) throw new TypeError(`${kind} builder has no method ${String(prop)}`);
      return (...args) => { spec[prop] = args; return proxy; };
    },
  });
  return proxy;
}

function pad(n) { return String(n).padStart(2, '0'); }

function createEnvironment({ now } = {}) {
  const ss = new FakeSpreadsheet();
  const mail = [];
  const drive = { folders: [] };
  const triggers = [];
  let fileSeq = 0;

  class FakeBlob {
    constructor(content, type, name) { Object.assign(this, { content, type, name: name || '' }); }
    getAs(type) { return new FakeBlob(this.content, type, this.name); }
    setName(n) { this.name = n; return this; }
    getName() { return this.name; }
    getContentType() { return this.type; }
    getDataAsString() { return this.content; }
  }
  const makeFolder = (name) => {
    const folder = {
      name, files: [],
      getName: () => name,
      createFile(blob) {
        const id = 'file' + (++fileSeq);
        const file = { id, blob, getId: () => id, getName: () => blob.name, getUrl: () => `https://drive.google.com/file/d/${id}/view`,
          getAs: (t) => blob.getAs(t) };
        folder.files.push(file); return file;
      },
    };
    drive.folders.push(folder); return folder;
  };
  const iterator = (items) => { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; };
  const triggerBuilder = (fn) => {
    const t = { fn, getHandlerFunction: () => fn };
    const b = {
      forSpreadsheet() { return b; }, onFormSubmit() { t.event = 'formSubmit'; return b; },
      onEdit() { t.event = 'edit'; return b; }, timeBased() { t.event = 'time'; return b; },
      everyDays(n) { t.everyDays = n; return b; }, atHour(h) { t.atHour = h; return b; },
      everyHours(n) { t.everyHours = n; return b; }, create() { triggers.push(t); return t; },
    };
    return b;
  };

  const RealDate = Date;
  const fixed = now ? new RealDate(now) : null;
  class ClockDate extends RealDate {
    constructor(...args) { if (args.length === 0 && fixed) super(fixed.getTime()); else super(...args); }
    static now() { return fixed ? fixed.getTime() : RealDate.now(); }
  }

  const globals = {
    console,
    Date: ClockDate,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getActive: () => ss,
      getUi: () => {
        const menu = { items: [], addItem(label, fn) { menu.items.push([label, fn]); return menu; },
          addSeparator() { return menu; }, addToUi() { ss.menu = menu; return menu; } };
        return { createMenu: (title) => { menu.title = title; return menu; } };
      },
      newDataValidation: () => builder('validation'),
      newConditionalFormatRule: () => builder('rule'),
    },
    Charts: { ChartType: { BAR: 'BAR', PIE: 'PIE', COLUMN: 'COLUMN' } },
    Utilities: {
      formatDate(date, tz, pattern) {
        if (tz !== 'UTC') throw new Error('fake formatDate only supports UTC');
        const d = new RealDate(date.getTime());
        return pattern
          .replace('yyyy', d.getUTCFullYear()).replace('MM', pad(d.getUTCMonth() + 1)).replace('dd', pad(d.getUTCDate()))
          .replace('HH', pad(d.getUTCHours())).replace('mm', pad(d.getUTCMinutes()));
      },
      parseDate(text, tz, pattern) {
        if (tz !== 'UTC' || pattern !== 'yyyy-MM-dd') throw new Error('fake parseDate only supports UTC yyyy-MM-dd');
        const [y, m, d] = text.split('-').map(Number);
        return new RealDate(RealDate.UTC(y, m - 1, d));
      },
    },
    Session: { getScriptTimeZone: () => 'UTC', getEffectiveUser: () => ({ getEmail: () => 'owner@shop.example' }) },
    HtmlService: { createHtmlOutput: (html) => ({ getBlob: () => new FakeBlob(html, 'text/html') }) },
    DriveApp: {
      getFoldersByName: (name) => iterator(drive.folders.filter((f) => f.name === name)),
      createFolder: (name) => makeFolder(name),
    },
    MailApp: { sendEmail: (msg) => { mail.push(msg); }, getRemainingDailyQuota: () => 100 },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => { triggers.splice(triggers.indexOf(t), 1); },
      newTrigger: triggerBuilder,
    },
  };
  const context = vm.createContext(globals);
  return { context, ss, mail, drive, triggers, FakeRange };
}

/** Load Order Desk into a fresh fake environment. `bundle` loads dist/OrderDesk.gs instead of src/. */
function loadOrderDesk(options = {}) {
  const env = createEnvironment(options);
  const root = path.join(__dirname, '..');
  const files = options.bundle ? [path.join(root, 'dist', 'OrderDesk.gs')] : SRC_ORDER.map((f) => path.join(root, 'src', f));
  files.forEach((file) => vm.runInContext(fs.readFileSync(file, 'utf8'), env.context, { filename: file }));
  env.run = (code) => vm.runInContext(code, env.context);
  /** Simulate a user editing one cell on a sheet (fires the installable onEdit handler). */
  env.edit = (sheetName, a1, value) => {
    const sheet = env.ss.getSheetByName(sheetName);
    const range = sheet.getRange(a1);
    const oldValue = range.getValue();
    range.setValue(value);
    env.context.handleEdit({ range, value, oldValue: oldValue === '' ? undefined : oldValue });
  };
  return env;
}

module.exports = { loadOrderDesk, SRC_ORDER };
