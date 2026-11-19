// Synchronous, in-memory emulation of Google Sheets API for Apps Script code.
// Throws NeedSheet error if data is missing, triggering lazy reload by runtime.

export class NeedSheet extends Error {
  constructor(name) {
    super(`Sheet "${name}" is not loaded.`);
    this.sheetName = name;
    this.isNeedSheet = true;
  }
}

const COL_RE = /^([A-Z]+)(\d*)$/i;

function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

function blank(v) {
  return v === undefined || v === null ? "" : v;
}

function cloneCell(v) {
  return v instanceof Date ? new Date(v.getTime()) : v;
}

export class Range {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }

  getValues() {
    const rows = this.sheet._rows();
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const src = rows[this.row - 1 + r];
      const line = [];
      for (let c = 0; c < this.numCols; c++) {
        line.push(src ? cloneCell(blank(src[this.col - 1 + c])) : "");
      }
      out.push(line);
    }
    return out;
  }

  getDisplayValues() {
    return this.getValues().map((row) =>
      row.map((v) => (v instanceof Date ? v.toISOString() : String(v)))
    );
  }

  getValue() {
    return this.getValues()[0][0];
  }

  setValues(values) {
    this.sheet._write(this.row, this.col, values);
    return this;
  }

  setValue(value) {
    const values = [];
    for (let r = 0; r < this.numRows; r++) {
      values.push(new Array(this.numCols).fill(value));
    }
    this.sheet._write(this.row, this.col, values);
    return this;
  }

  clearContent() {
    const values = [];
    for (let r = 0; r < this.numRows; r++) {
      values.push(new Array(this.numCols).fill(""));
    }
    this.sheet._write(this.row, this.col, values, true);
    return this;
  }

  clear() {
    return this.clearContent();
  }

  setFormula(formula) { return this.setValue(formula); }
  setNumberFormat() { return this; }
  setBackground() { return this; }
  setFontWeight() { return this; }
  setFontColor() { return this; }
  setHorizontalAlignment() { return this; }
  setDataValidation() { return this; }
  setNote() { return this; }
  setWrap() { return this; }
  sort() { return this; }
}

export class Sheet {
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
  }

  getName() { return this.name; }
  getSheetId() { return this.name; }

  _state() {
    return this.workbook._state(this.name);
  }

  _rows() {
    return this._state().rows;
  }

  _own() {
    const state = this._state();
    if (!state.owned) {
      state.rows = state.rows.map((r) => r.slice());
      state.owned = true;
    }
    state.dirty = true;
    return state.rows;
  }

  _write(row, col, values, allowShrinkOnly) {
    const rows = this._own();
    for (let r = 0; r < values.length; r++) {
      const idx = row - 1 + r;
      while (rows.length <= idx) rows.push([]);
      const target = rows[idx];
      for (let c = 0; c < values[r].length; c++) {
        const cIdx = col - 1 + c;
        while (target.length <= cIdx) target.push("");
        target[cIdx] = cloneCell(blank(values[r][c]));
      }
    }
    this.workbook._touched(this.name, allowShrinkOnly);
  }

  getLastRow() {
    const rows = this._rows();
    let last = rows.length;
    while (last > 0 && rows[last - 1].every((v) => blank(v) === "")) last--;
    return last;
  }

  getLastColumn() {
    let max = 0;
    for (const row of this._rows()) {
      for (let c = row.length; c > max; c--) {
        if (blank(row[c - 1]) !== "") {
          max = Math.max(max, c);
          break;
        }
      }
    }
    return max;
  }

  getMaxRows() { return Math.max(this._rows().length, 1000); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 26); }

  insertRowsAfter() { return this; }
  insertColumnsAfter() { return this; }

  getDataRange() {
    return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  getRange(a, b, c, d) {
    if (typeof a === "string") return this._a1(a);
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }

  _a1(text) {
    const parts = text.split(":");
    const parse = (s) => {
      const m = COL_RE.exec(s.trim());
      if (!m) throw new Error(`Unsupported range format: ${text}`);
      return { col: colToNum(m[1]), row: m[2] ? Number(m[2]) : null };
    };
    const start = parse(parts[0]);
    const end = parts[1] ? parse(parts[1]) : start;
    const lastRow = Math.max(this.getLastRow(), 1);
    const r1 = start.row || 1;
    const r2 = end.row || lastRow;
    return new Range(this, r1, start.col, r2 - r1 + 1, end.col - start.col + 1);
  }

  appendRow(values) {
    const rows = this._own();
    let last = this.getLastRow();
    rows.length = last;
    rows.push(values.map((v) => cloneCell(blank(v))));
    this.workbook._touched(this.name);
    return this;
  }

  deleteRow(n) {
    const rows = this._own();
    if (n >= 1 && n <= rows.length) rows.splice(n - 1, 1);
    this.workbook._touched(this.name);
    return this;
  }

  deleteRows(n, count) {
    const rows = this._own();
    rows.splice(n - 1, count);
    this.workbook._touched(this.name);
    return this;
  }

  insertRowAfter(n) {
    const rows = this._own();
    rows.splice(n, 0, []);
    return this;
  }

  insertRowBefore(n) {
    const rows = this._own();
    rows.splice(n - 1, 0, []);
    return this;
  }

  clearContents() {
    const rows = this._own();
    rows.length = 0;
    this.workbook._touched(this.name);
    return this;
  }

  clear() { return this.clearContents(); }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  setTabColor() { return this; }
  hideSheet() { return this; }
  autoResizeColumns() { return this; }
}

export class Workbook {
  constructor({ known, loaded, timezone, virtualSheets }) {
    this.timezone = timezone;
    this.known = new Map(known);
    this.states = new Map();
    this.created = new Set();
    this.deleted = new Set();
    this.virtual = virtualSheets || new Map();
    for (const [name, data] of loaded || []) {
      this.states.set(name, {
        rows: data.rows,
        version: data.version,
        original: data.rows,
        owned: false,
        dirty: false,
      });
    }
    this.sheets = new Map();
    this.missing = new Set();
  }

  _state(name) {
    const state = this.states.get(name);
    if (!state) {
      this.missing.add(name);
      throw new NeedSheet(name);
    }
    return state;
  }

  _touched() {}

  exists(name) {
    return this.known.has(name) && !this.deleted.has(name);
  }

  getId() { return "cloudflare-d1-workbook"; }
  getName() { return "School Management System SaaS"; }
  getSpreadsheetTimeZone() { return this.timezone; }

  getSheetByName(name) {
    name = String(name);
    if (!this.exists(name)) return null;
    return this._sheetObject(name);
  }

  _sheetObject(name) {
    let sheet = this.sheets.get(name);
    if (!sheet) {
      sheet = new Sheet(this, name);
      this.sheets.set(name, sheet);
    }
    return sheet;
  }

  getSheets() {
    return [...this.known.keys()]
      .filter((n) => !this.deleted.has(n) && !n.startsWith("__"))
      .map((n) => this._sheetObject(n));
  }

  insertSheet(name) {
    name = String(name);
    if (this.exists(name)) throw new Error(`A sheet with the name "${name}" already exists.`);
    this.known.set(name, 0);
    this.deleted.delete(name);
    this.created.add(name);
    this.states.set(name, { rows: [], version: 0, original: [], owned: true, dirty: true, isNew: true });
    return this._sheetObject(name);
  }

  deleteSheet(sheet) {
    const name = typeof sheet === "string" ? sheet : sheet.getName();
    this.deleted.add(name);
    this.created.delete(name);
  }

  flush() {}

  collectChanges() {
    const changes = [];
    for (const name of this.deleted) {
      if (this.virtual.has(name)) continue;
      const state = this.states.get(name);
      changes.push({ sheet: name, op: "drop", expected: state ? state.version : this.known.get(name) });
    }
    for (const [name, state] of this.states) {
      if (this.deleted.has(name) || !state.dirty || this.virtual.has(name)) continue;
      const rows = state.rows;
      let length = rows.length;
      while (length > 0 && rows[length - 1].every((v) => blank(v) === "")) length--;
      const set = [];
      const original = state.original || [];
      for (let i = 0; i < length; i++) {
        if (!state.isNew && original[i] && sameRow(original[i], rows[i])) continue;
        set.push([i + 1, rows[i]]);
      }
      const truncated = state.isNew ? false : original.length > length;
      if (!set.length && !truncated && !state.isNew) continue;
      changes.push({
        sheet: name,
        op: state.isNew ? "create" : "update",
        expected: state.version,
        length,
        set,
      });
    }
    return changes;
  }
}

function sameRow(a, b) {
  if (a === b) return true;
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i++) {
    const x = blank(a[i]);
    const y = blank(b[i]);
    if (x === y) continue;
    if (x instanceof Date && y instanceof Date && x.getTime() === y.getTime()) continue;
    return false;
  }
  return true;
}