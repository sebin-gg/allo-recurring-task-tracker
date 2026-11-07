// In-memory mock of the Google Apps Script services Code.gs uses.
// Load Code.gs with loadGas() — returns its public functions bound to the mock.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));

/* ---------- mutable clock: tests freeze/advance "now" ---------- */
export const clock = { now: new Date('2026-11-07T19:00:00Z') };
class FakeDate extends Date {
  constructor(...args) {
    super(...(args.length ? args : [clock.now.getTime()]));
  }
  static now() { return clock.now.getTime(); }
}

/* ---------- range / sheet mocks ---------- */
const colToNum = (s) => s.toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const numToCol = (n) => { let s = ''; while (n > 0) { s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s; n = Math.floor((n - 1) / 26); } return s; };
const parseA1 = (a1) => { const m = /^([A-Za-z]+)(\d+)$/.exec(a1); return m ? { row: +m[2], col: colToNum(m[1]) } : null; };

class MockRange {
  constructor(sheet, row, col, nRows = 1, nCols = 1) {
    Object.assign(this, { sheet, row, col, nRows, nCols });
  }
  _grid() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const src = this.sheet.data[this.row - 1 + r] || [];
      const row = [];
      for (let c = 0; c < this.nCols; c++) row.push(src[this.col - 1 + c] ?? '');
      out.push(row);
    }
    return out;
  }
  getValues() { return this._grid(); }
  getValue() { return this._grid()[0][0]; }
  setValues(vals) {
    vals.forEach((vr, r) => vr.forEach((v, c) => {
      const ri = this.row - 1 + r, ci = this.col - 1 + c;
      this.sheet.data[ri] = this.sheet.data[ri] || [];
      this.sheet.data[ri][ci] = v;
    }));
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  getA1Notation() { return `${numToCol(this.col)}${this.row}`; }
  getCell(r, c) { return new MockRange(this.sheet, this.row + r - 1, this.col + c - 1); }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getSheet() { return this.sheet; }
  // no-ops used by renderers/setup
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
  setFormula(f) { return this.setValues([[f]]); }
  setDataValidation() { return this; }
  insertCheckboxes() { return this; }
  clearDataValidations() { return this; }
  clear() { return this.setValues(this._grid().map((r) => r.map(() => ''))); }
}

class MockSheet {
  constructor(name) { this.name = name; this.data = []; }
  getName() { return this.name; }
  getLastRow() {
    let n = 0;
    this.data.forEach((r, i) => { if (r && r.some((c) => c !== '' && c != null)) n = i + 1; });
    return n;
  }
  _maxCols() { return this.data.reduce((m, r) => Math.max(m, r ? r.length : 0), 1); }
  getDataRange() { return new MockRange(this, 1, 1, Math.max(this.getLastRow(), 1), this._maxCols()); }
  appendRow(vals) { this.data.push([...vals]); }
  getRange(...args) {
    if (typeof args[0] === 'string') {
      const a1 = args[0];
      if (a1.includes(':')) { // A:A or A1:B10
        const [a, b] = a1.split(':');
        if (/^[A-Za-z]+$/.test(a)) { const c = colToNum(a); return new MockRange(this, 1, c, Math.max(this.getLastRow(), 1), 1); }
        const p = parseA1(a), q = parseA1(b);
        return new MockRange(this, p.row, p.col, q.row - p.row + 1, q.col - p.col + 1);
      }
      const p = parseA1(a1);
      return new MockRange(this, p.row, p.col);
    }
    return new MockRange(this, args[0], args[1], args[2] ?? 1, args[3] ?? 1);
  }
}

export function makeWorld() {
  const sheets = new Map();
  const mail = [];
  const toasts = [];
  const ss = {
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new MockSheet(n); sheets.set(n, s); return s; },
    getSpreadsheetTimeZone: () => 'UTC',
    toast: (m) => toasts.push(m),
  };
  const SpreadsheetApp = {
    getActive: () => ss,
    getUi: () => ({ createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) }),
    newDataValidation: () => ({ requireValueInRange: () => ({ build: () => ({}) }) }),
  };
  const LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  // GAS formats in the spreadsheet TZ; our mock's "local" is the spreadsheet local.
  const pad = (n) => String(n).padStart(2, '0');
  const Utilities = { formatDate: (d, tz, fmt) => fmt === 'yyyy-MM-dd'
    ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : String(d) };
  const MailApp = { sendEmail: (m) => mail.push(m) };
  const ScriptApp = {
    getProjectTriggers: () => [], deleteTrigger() {},
    newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => ({ getHandlerFunction: () => 'runDailyChecks' }) }) }) }) }),
  };
  return { ss, sheets, mail, toasts, SpreadsheetApp, LockService, Utilities, MailApp, ScriptApp };
}

export function loadGas(world = makeWorld()) {
  const src = readFileSync(join(here, '..', 'Code.gs'), 'utf8');
  const fns = ['setupSheets', 'installTriggers', 'onOpen', 'ensureCycles', 'runDailyChecks',
    'completeCycle', 'acknowledge', 'onEdit', 'renderMyTasks', 'renderDashboard',
    'demoDueToday', 'demoForceEscalation'];
  const factory = new Function(
    'SpreadsheetApp', 'LockService', 'Utilities', 'MailApp', 'ScriptApp', 'Date', 'console',
    `${src}\nreturn { ${fns.join(', ')} };`
  );
  const api = factory(world.SpreadsheetApp, world.LockService, world.Utilities,
    world.MailApp, world.ScriptApp, FakeDate, console);
  return { api, world };
}

/* ---------- test helpers ---------- */
export function sheetRows(world, name) {
  return world.sheets.get(name).getDataRange().getValues();
}
export function cycleRows(world) {
  return sheetRows(world, 'Cycles').slice(1).filter((r) => r[0]);
}
export function logKeys(world) {
  return sheetRows(world, 'Log').slice(1).filter((r) => r[0]).map((r) => r[0]);
}
export function setConfig(world, key, value) {
  const cfg = world.sheets.get('Config');
  const rows = cfg.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { cfg.getRange(i + 1, 2).setValue(value); return; }
  }
  cfg.appendRow([key, value]);
}
