/* Mô hình dữ liệu bảng tính + chấm điểm bài tập Excel (không phụ thuộc DOM). */
(function (global) {
  'use strict';
  const XL = global.XL || require('./formula.js');

  const key = (r, c) => r + ',' + c;

  // Chuyển chuỗi người dùng gõ thành nội dung ô (giống Excel: số, %, ngày dd/mm/yyyy, TRUE/FALSE, chữ).
  function parseInput(text) {
    if (text === null || text === undefined) return { lit: null };
    const s = String(text);
    if (s.trim() === '') return { lit: null };
    if (s.startsWith('=') && s.length > 1) return { formula: s };
    if (s.startsWith("'")) return { lit: s.slice(1) };
    const t = s.trim();
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(t)) return { lit: parseFloat(t) };
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { lit: parseFloat(t.replace(/,/g, '')) };
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)%$/.test(t)) return { lit: parseFloat(t) / 100, fmt: 'percent' };
    const d = XL.parseDateText(t);
    if (d !== null) return { lit: d, fmt: 'date' };
    if (/^(true|false)$/i.test(t)) return { lit: t.toUpperCase() === 'TRUE' };
    return { lit: s };
  }

  class Sheet {
    constructor(rows, cols) {
      this.nRows = rows;
      this.nCols = cols;
      this.cells = new Map();
      this.values = new Map();
      this.dirty = true;
    }
    cell(r, c) { return this.cells.get(key(r, c)) || null; }
    ensure(r, c) {
      let x = this.cells.get(key(r, c));
      if (!x) { x = {}; this.cells.set(key(r, c), x); }
      return x;
    }
    // Đặt dữ liệu cho sẵn (từ đề bài).
    put(r, c, v, style) {
      const x = this.ensure(r, c);
      if (v && typeof v === 'object' && v.date) {
        const [y, m, d] = v.date.split('-').map(Number);
        x.lit = XL.dateSerial(y, m, d);
        x.fmt = 'date';
      } else if (typeof v === 'string' && v.startsWith('=')) {
        x.formula = v;
      } else x.lit = v === '' ? null : v;
      if (style) Object.assign(x, style);
      this.dirty = true;
    }
    // Ghi nội dung người dùng gõ; trả về false nếu ô bị khóa.
    setInput(r, c, text) {
      const x = this.ensure(r, c);
      if (x.locked) return false;
      const p = parseInput(text);
      delete x.formula; delete x.tree; delete x.parseError;
      x.lit = null;
      if (p.formula) x.formula = p.formula;
      else x.lit = p.lit;
      if (p.fmt && (!x.fmt || x.fmt === 'general')) x.fmt = p.fmt;
      this.dirty = true;
      return true;
    }
    input(r, c) {
      const x = this.cell(r, c);
      if (!x) return '';
      if (x.formula) return x.formula;
      if (x.lit === null || x.lit === undefined) return '';
      if (typeof x.lit === 'number' && x.fmt === 'date') return XL.formatDate(x.lit);
      if (typeof x.lit === 'number' && x.fmt === 'percent') return XL.fmtGeneral(x.lit * 100) + '%';
      if (typeof x.lit === 'boolean') return x.lit ? 'TRUE' : 'FALSE';
      if (typeof x.lit === 'string' && /^[+-]?\d/.test(x.lit) && parseInput(x.lit).lit !== x.lit) return "'" + x.lit;
      return typeof x.lit === 'number' ? XL.fmtGeneral(x.lit) : String(x.lit);
    }
    isFormula(r, c) { const x = this.cell(r, c); return !!(x && x.formula); }
    compute() {
      this.values = new Map();
      const visiting = new Set();
      const self = this;
      const ctx = {
        get(r, c) {
          const k = key(r, c);
          if (self.values.has(k)) return self.values.get(k);
          const x = self.cells.get(k);
          if (!x) return null;
          if (!x.formula) return x.lit === undefined ? null : x.lit;
          if (visiting.has(k)) return XL.E.CIRC;
          visiting.add(k);
          let v;
          try {
            if (!x.tree) x.tree = XL.parse(x.formula.slice(1));
            v = XL.evaluate(x.tree, ctx);
            if (v instanceof XL.Range) v = v.h === 1 && v.w === 1 ? v.rows[0][0] : XL.E.VALUE;
            if (v === null || v === undefined) v = 0;
            if (typeof v === 'number' && !isFinite(v)) v = XL.E.NUM;
            delete x.parseError;
          } catch (e) {
            x.parseError = e.message;
            v = new XL.XLError('#LỖI');
          }
          visiting.delete(k);
          self.values.set(k, v);
          return v;
        },
      };
      for (const k of this.cells.keys()) {
        const [r, c] = k.split(',').map(Number);
        ctx.get(r, c);
      }
      this.ctx = ctx;
      this.dirty = false;
    }
    value(r, c) {
      if (this.dirty) this.compute();
      const k = key(r, c);
      if (this.values.has(k)) return this.values.get(k);
      const x = this.cells.get(k);
      return x && !x.formula ? (x.lit === undefined ? null : x.lit) : null;
    }
    display(r, c) {
      const v = this.value(r, c);
      if (v === null || v === undefined) return '';
      if (XL.isErr(v)) return v.code;
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'number') {
        const x = this.cell(r, c);
        return XL.formatNumber(v, (x && x.fmt) || 'general');
      }
      return String(v);
    }
    clone() {
      const s = new Sheet(this.nRows, this.nCols);
      for (const [k, x] of this.cells) { const y = Object.assign({}, x); delete y.tree; s.cells.set(k, y); }
      s.dirty = true;
      return s;
    }
    // Chỉ lưu phần người dùng nhập (ô không khóa) để khôi phục tiến độ.
    userData() {
      const out = {};
      for (const [k, x] of this.cells) {
        if (x.locked) continue;
        if (x.formula || (x.lit !== null && x.lit !== undefined) || (x.fmt && x.fmt !== 'general')) {
          out[k] = { f: x.formula, l: x.lit, fmt: x.fmt };
        }
      }
      return out;
    }
    loadUserData(data) {
      for (const k in data) {
        const [r, c] = k.split(',').map(Number);
        const x = this.ensure(r, c);
        if (x.locked) continue;
        const d = data[k];
        delete x.tree;
        x.formula = d.f || undefined;
        if (!x.formula) delete x.formula;
        x.lit = d.l === undefined ? null : d.l;
        if (d.fmt) x.fmt = d.fmt;
      }
      this.dirty = true;
    }
  }

  // ---------- Dựng bảng tính từ đề bài ----------
  const STYLE = {
    title: { locked: true, cls: 'x-title' },
    head: { locked: true, cls: 'x-head' },
    data: { locked: true, cls: 'x-data' },
    label: { locked: true, cls: 'x-label' },
    note: { locked: true, cls: 'x-note' },
  };
  function buildSheet(ex) {
    const s = new Sheet(ex.rows || 30, ex.cols || 10);
    for (const b of ex.blocks) {
      const p = XL.parseAddr(b.at);
      b.rows.forEach((row, i) => {
        (Array.isArray(row) ? row : [row]).forEach((v, j) => {
          if (v === undefined) return;
          const st = Object.assign({}, STYLE[b.style || 'data']);
          if (b.fmt) st.fmt = b.fmt;
          if (b.merge && i === 0 && j === 0) st.span = b.merge;
          s.put(p.r + i, p.c + j, v, st);
        });
      });
    }
    for (const t of ex.tasks || []) {
      const rg = XL.parseRange(t.range);
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const x = s.ensure(r, c);
        x.locked = false;
        x.task = t.id;
        if (t.cellFmt && !x.fmt) x.fmt = t.cellFmt;
      }
    }
    s.compute();
    return s;
  }

  function eachCell(rg, fn) {
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) fn(r, c);
  }

  function fillAnswer(s, t) {
    const rg = XL.parseRange(t.range);
    let i = 0;
    eachCell(rg, (r, c) => {
      const x = s.ensure(r, c);
      delete x.tree;
      if (t.values) { x.lit = t.values[i++]; delete x.formula; }
      else if (t.answer) { x.formula = XL.shiftFormula(t.answer, r - rg.r1, c - rg.c1); }
      if (t.fmt) x.fmt = t.fmt;
    });
    s.dirty = true;
  }

  // Bảng tính chuẩn: lấy dữ liệu đề + điền toàn bộ đáp án.
  function solvedSheet(ex) {
    const s = buildSheet(ex);
    for (const t of ex.tasks) fillAnswer(s, t);
    s.compute();
    return s;
  }

  const normStr = (v) => String(v).normalize('NFC').trim().replace(/\s+/g, ' ').toLowerCase();
  function sameValue(u, e) {
    if (e === null || e === '') return u === null || u === '' || u === 0;
    if (XL.isErr(e)) return XL.isErr(u) && u.code === e.code;
    if (XL.isErr(u)) return false;
    if (typeof e === 'number') {
      if (typeof u !== 'number') return false;
      return Math.abs(u - e) <= 1e-6 * Math.max(1, Math.abs(e));
    }
    if (typeof e === 'boolean') return u === e;
    if (u === null) return false;
    return normStr(u) === normStr(e);
  }

  // Chấm từng câu: trả về trạng thái + kết quả từng ô.
  function grade(sheet, ex, solved) {
    solved = solved || solvedSheet(ex);
    sheet.compute();
    const results = {};
    for (const t of ex.tasks) {
      const rg = XL.parseRange(t.range);
      const cells = [];
      let filled = 0, okCount = 0, constCount = 0, fmtBad = 0, missingFn = null;
      eachCell(rg, (r, c) => {
        const x = sheet.cell(r, c) || {};
        const hasContent = !!x.formula || (x.lit !== null && x.lit !== undefined && x.lit !== '');
        const u = sheet.value(r, c), e = solved.value(r, c);
        let ok = hasContent && sameValue(u, e);
        let msg = '';
        if (t.fmt && x.fmt !== t.fmt) { fmtBad++; msg = 'Chưa đúng định dạng'; }
        if (hasContent) filled++;
        if (ok && !t.allowValue && !x.formula) { constCount++; msg = 'Đúng số nhưng phải dùng công thức'; }
        if (ok && x.formula && t.mustUse) {
          const used = XL.usedFunctions(x.formula);
          const miss = t.mustUse.find((f) => !used.includes(f));
          if (miss) { missingFn = miss; msg = 'Cần dùng hàm ' + miss; }
        }
        if (t.type === 'format') ok = x.fmt === t.fmt;
        if (ok) okCount++;
        cells.push({ r, c, ok, filled: hasContent || t.type === 'format', expected: e, msg });
      });
      const total = cells.length;
      let status;
      if (t.type === 'format') status = okCount === total ? 'ok' : okCount ? 'partial' : 'empty';
      else if (filled === 0) status = 'empty';
      else if (okCount === total && constCount === 0 && !missingFn && (!t.fmt || fmtBad === 0)) status = 'ok';
      else if (okCount === total && constCount > 0) status = 'const';
      else if (okCount === total && missingFn) status = 'fn';
      else if (okCount === total && fmtBad) status = 'fmt';
      else if (okCount > 0) status = 'partial';
      else status = 'wrong';
      results[t.id] = { status, cells, okCount, total, filled, missingFn };
    }
    return results;
  }

  function score(ex, results) {
    let got = 0, max = 0;
    for (const t of ex.tasks) {
      const p = t.pts || 1;
      max += p;
      const r = results[t.id];
      if (!r) continue;
      if (r.status === 'ok') got += p;
      else if (r.status === 'partial' || r.status === 'const' || r.status === 'fn' || r.status === 'fmt') got += p * (r.okCount / r.total) * 0.5;
    }
    return { got: Math.round(got * 100) / 100, max };
  }

  const api = { Sheet, parseInput, buildSheet, solvedSheet, fillAnswer, grade, score, sameValue };
  global.SheetModel = api;
  if (typeof module !== 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
