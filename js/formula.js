/* Bộ máy tính công thức kiểu Excel: tokenizer, parser, evaluator, các hàm thường gặp trong đề thi. */
(function (global) {
  'use strict';

  class XLError {
    constructor(code) { this.code = code; }
    toString() { return this.code; }
  }
  const E = {
    DIV0: new XLError('#DIV/0!'), NA: new XLError('#N/A'), VALUE: new XLError('#VALUE!'),
    NAME: new XLError('#NAME?'), REF: new XLError('#REF!'), NUM: new XLError('#NUM!'), CIRC: new XLError('#CIRC!'),
  };
  const isErr = (v) => v instanceof XLError;

  // ---------- Địa chỉ ô ----------
  function colToNum(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function numToCol(n) {
    let s = '';
    n += 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  const addr = (r, c) => numToCol(c) + (r + 1);
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim());
    if (!m) return null;
    return { r: parseInt(m[2], 10) - 1, c: colToNum(m[1]) };
  }
  function parseRange(s) {
    const [a, b] = s.split(':');
    const p = parseAddr(a), q = b ? parseAddr(b) : p;
    if (!p || !q) return null;
    return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
  }

  // ---------- Ngày tháng (serial kiểu Excel) ----------
  const EPOCH = Date.UTC(1899, 11, 30);
  const dateSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
  const serialToDate = (n) => new Date(EPOCH + Math.floor(n) * 86400000);
  function formatDate(n) {
    const d = serialToDate(n);
    const p = (x) => String(x).padStart(2, '0');
    return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  function todaySerial() {
    const t = new Date();
    return dateSerial(t.getFullYear(), t.getMonth() + 1, t.getDate());
  }

  // ---------- Tokenizer ----------
  const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/;
  function tokenize(src) {
    const out = [];
    let i = 0;
    while (i < src.length) {
      const ch = src[i];
      const rest = src.slice(i);
      if (/\s/.test(ch)) { i++; continue; }
      let m;
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < src.length) {
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        if (j >= src.length) throw new Error('Thiếu dấu " đóng chuỗi');
        out.push({ t: 'str', v: s, s: i, e: j + 1 });
        i = j + 1; continue;
      }
      if ((m = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(rest))) {
        out.push({ t: 'num', v: parseFloat(m[0]), s: i, e: i + m[0].length });
        i += m[0].length; continue;
      }
      if ((m = REF_RE.exec(rest)) && parseInt(m[4], 10) > 0) {
        out.push({ t: 'ref', absC: !!m[1], col: colToNum(m[2]), absR: !!m[3], row: parseInt(m[4], 10) - 1, s: i, e: i + m[0].length });
        i += m[0].length; continue;
      }
      if ((m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest))) {
        const name = m[0].toUpperCase();
        let j = i + m[0].length;
        while (j < src.length && src[j] === ' ') j++;
        if (src[j] === '(') out.push({ t: 'func', v: name, s: i, e: i + m[0].length });
        else if (name === 'TRUE' || name === 'FALSE') out.push({ t: 'bool', v: name === 'TRUE', s: i, e: i + m[0].length });
        else out.push({ t: 'name', v: m[0], s: i, e: i + m[0].length });
        i += m[0].length; continue;
      }
      if ((m = /^(<=|>=|<>)/.exec(rest))) { out.push({ t: 'op', v: m[0], s: i, e: i + 2 }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { out.push({ t: 'op', v: ch, s: i, e: i + 1 }); i++; continue; }
      if (ch === '(' || ch === ')') { out.push({ t: ch, s: i, e: i + 1 }); i++; continue; }
      if (ch === ',' || ch === ';') { out.push({ t: 'sep', s: i, e: i + 1 }); i++; continue; }
      if (ch === ':') { out.push({ t: ':', s: i, e: i + 1 }); i++; continue; }
      throw new Error('Ký tự không hợp lệ: ' + ch);
    }
    return out;
  }

  // ---------- Parser ----------
  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (...ops) => peek() && peek().t === 'op' && ops.includes(peek().v);
    function expect(t) {
      const k = toks[p];
      if (!k || k.t !== t) throw new Error('Thiếu "' + t + '"');
      p++; return k;
    }
    function comparison() {
      let n = concat();
      while (isOp('=', '<>', '<', '>', '<=', '>=')) { const op = toks[p++].v; n = { k: 'bin', op, a: n, b: concat() }; }
      return n;
    }
    function concat() {
      let n = additive();
      while (isOp('&')) { p++; n = { k: 'bin', op: '&', a: n, b: additive() }; }
      return n;
    }
    function additive() {
      let n = mult();
      while (isOp('+', '-')) { const op = toks[p++].v; n = { k: 'bin', op, a: n, b: mult() }; }
      return n;
    }
    function mult() {
      let n = power();
      while (isOp('*', '/')) { const op = toks[p++].v; n = { k: 'bin', op, a: n, b: power() }; }
      return n;
    }
    function power() {
      let n = unary();
      while (isOp('^')) { p++; n = { k: 'bin', op: '^', a: n, b: unary() }; }
      return n;
    }
    function unary() {
      if (isOp('-')) { p++; return { k: 'neg', a: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      return postfix();
    }
    function postfix() {
      let n = primary();
      while (isOp('%')) { p++; n = { k: 'pct', a: n }; }
      return n;
    }
    function primary() {
      const k = toks[p];
      if (!k) throw new Error('Công thức chưa hoàn chỉnh');
      if (k.t === 'num') { p++; return { k: 'lit', v: k.v }; }
      if (k.t === 'str') { p++; return { k: 'lit', v: k.v }; }
      if (k.t === 'bool') { p++; return { k: 'lit', v: k.v }; }
      if (k.t === 'ref') {
        p++;
        if (peek() && peek().t === ':') {
          p++;
          const b = expect('ref');
          return { k: 'range', r1: Math.min(k.row, b.row), c1: Math.min(k.col, b.col), r2: Math.max(k.row, b.row), c2: Math.max(k.col, b.col) };
        }
        return { k: 'ref', r: k.row, c: k.col };
      }
      if (k.t === 'func') {
        p++; expect('(');
        const args = [];
        if (peek() && peek().t === ')') { p++; return { k: 'call', name: k.v, args }; }
        for (;;) {
          if (peek() && (peek().t === 'sep' || peek().t === ')')) args.push({ k: 'empty' });
          else args.push(comparison());
          if (peek() && peek().t === 'sep') { p++; continue; }
          expect(')');
          break;
        }
        return { k: 'call', name: k.v, args };
      }
      if (k.t === '(') { p++; const n = comparison(); expect(')'); return n; }
      if (k.t === 'name') { p++; return { k: 'badname', v: k.v }; }
      throw new Error('Không hiểu ký hiệu "' + (k.v || k.t) + '"');
    }
    const tree = comparison();
    if (p < toks.length) throw new Error('Thừa ký hiệu ở cuối công thức');
    return tree;
  }

  // ---------- Ép kiểu ----------
  class Range {
    constructor(rows) { this.rows = rows; }
    get h() { return this.rows.length; }
    get w() { return this.rows[0] ? this.rows[0].length : 0; }
    flat() { return this.rows.flat(); }
  }
  function scalar(v) {
    if (v instanceof Range) return v.h === 1 && v.w === 1 ? v.rows[0][0] : E.VALUE;
    return v;
  }
  function toNum(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v === null || v === undefined) return 0;
    const s = String(v).trim();
    if (s === '') return E.VALUE;
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?%?$/.test(s)) return s.endsWith('%') ? parseFloat(s) / 100 : parseFloat(s);
    const d = parseDateText(s);
    if (d !== null) return d;
    return E.VALUE;
  }
  function fmtGeneral(n) {
    if (!isFinite(n)) return E.NUM.code;
    if (Number.isInteger(n)) return String(n);
    const r = parseFloat(n.toPrecision(12));
    return String(r);
  }
  function toStr(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return fmtGeneral(v);
    return String(v);
  }
  function toBool(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === null || v === undefined) return false;
    const s = String(v).toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    return E.VALUE;
  }
  function parseDateText(s) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s.trim());
    if (!m) return null;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const d = parseInt(m[1], 10), mo = parseInt(m[2], 10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return dateSerial(y, mo, d);
  }

  const norm = (s) => String(s).normalize('NFC').toLowerCase();
  // So sánh kiểu Excel: số < chuỗi < logic; chuỗi không phân biệt hoa thường.
  function compare(a, b) {
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const rank = (x) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2);
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'string') { const x = norm(a), y = norm(b); return x < y ? -1 : x > y ? 1 : 0; }
    if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
    return a < b ? -1 : a > b ? 1 : 0;
  }

  function wildcardRe(pat) {
    let re = '';
    for (let i = 0; i < pat.length; i++) {
      const ch = pat[i];
      if (ch === '~' && i + 1 < pat.length) { re += pat[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
      if (ch === '*') re += '.*';
      else if (ch === '?') re += '.';
      else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }
  // Điều kiện kiểu SUMIF/COUNTIF: ">20", "<>Nam", "Máy*", 15 ...
  function makeCriteria(crit) {
    crit = scalar(crit);
    if (typeof crit === 'number' || typeof crit === 'boolean') return (v) => typeof v === typeof crit && v === crit;
    const s = crit === null || crit === undefined ? '' : String(crit);
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
    const op = m[1] || '=';
    const rhs = m[2];
    const num = /^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(rhs.trim()) ? parseFloat(rhs) : parseDateText(rhs);
    if (num !== null && rhs.trim() !== '') {
      return (v) => {
        if (typeof v !== 'number') return op === '<>';
        switch (op) {
          case '=': return v === num; case '<>': return v !== num;
          case '<': return v < num; case '>': return v > num;
          case '<=': return v <= num; case '>=': return v >= num;
        }
        return false;
      };
    }
    if (op === '=' || op === '<>') {
      const re = wildcardRe(rhs.normalize('NFC'));
      const hit = (v) => (rhs === '' ? v === null || v === '' : v !== null && re.test(toStr(v).normalize('NFC')));
      return op === '=' ? hit : (v) => !hit(v);
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const c = compare(v, rhs);
      return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    };
  }

  function lookupEq(a, b) {
    if (typeof a === 'string' && typeof b === 'string') return wildcardRe(a.normalize('NFC')).test(b.normalize('NFC'));
    if (typeof a !== typeof b) return false;
    return a === b;
  }

  function roundHalfUp(x, d) {
    const f = Math.pow(10, d);
    const v = Math.abs(x) * f;
    const r = Math.round(parseFloat(v.toPrecision(15)));
    return Math.sign(x) * r / f;
  }

  // Gom số từ tham số: ô rỗng/chuỗi/logic trong vùng bị bỏ qua (giống SUM của Excel).
  function collectNums(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Range) {
        for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); }
      } else if (a !== null && a !== undefined) {
        const n = toNum(a);
        if (isErr(n)) return n;
        out.push(n);
      }
    }
    return out;
  }
  const asRange = (v) => (v instanceof Range ? v : new Range([[v]]));

  // ---------- Thư viện hàm ----------
  const F = {
    SUM: (...a) => { const n = collectNums(a); return isErr(n) ? n : n.reduce((s, x) => s + x, 0); },
    AVERAGE: (...a) => { const n = collectNums(a); if (isErr(n)) return n; return n.length ? n.reduce((s, x) => s + x, 0) / n.length : E.DIV0; },
    MAX: (...a) => { const n = collectNums(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; },
    MIN: (...a) => { const n = collectNums(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; },
    COUNT: (...a) => a.reduce((s, x) => s + (x instanceof Range ? x.flat().filter((v) => typeof v === 'number').length : typeof toNum(x) === 'number' ? 1 : 0), 0),
    COUNTA: (...a) => a.reduce((s, x) => s + (x instanceof Range ? x.flat().filter((v) => v !== null && v !== '').length : 1), 0),
    COUNTBLANK: (r) => asRange(r).flat().filter((v) => v === null || v === '').length,
    ABS: (x) => { x = toNum(x); return isErr(x) ? x : Math.abs(x); },
    SQRT: (x) => { x = toNum(x); return isErr(x) ? x : x < 0 ? E.NUM : Math.sqrt(x); },
    POWER: (x, y) => { x = toNum(x); y = toNum(y); return isErr(x) ? x : isErr(y) ? y : Math.pow(x, y); },
    INT: (x) => { x = toNum(x); return isErr(x) ? x : Math.floor(x); },
    MOD: (x, y) => { x = toNum(x); y = toNum(y); if (isErr(x)) return x; if (isErr(y)) return y; if (y === 0) return E.DIV0; return x - y * Math.floor(x / y); },
    ROUND: (x, d) => { x = toNum(x); d = toNum(d ?? 0); return isErr(x) ? x : isErr(d) ? d : roundHalfUp(x, Math.trunc(d)); },
    ROUNDUP: (x, d) => { x = toNum(x); d = Math.trunc(toNum(d ?? 0)); if (isErr(x)) return x; const f = Math.pow(10, d); return Math.sign(x) * Math.ceil(parseFloat((Math.abs(x) * f).toPrecision(15))) / f; },
    ROUNDDOWN: (x, d) => { x = toNum(x); d = Math.trunc(toNum(d ?? 0)); if (isErr(x)) return x; const f = Math.pow(10, d); return Math.sign(x) * Math.floor(parseFloat((Math.abs(x) * f).toPrecision(15))) / f; },
    PRODUCT: (...a) => { const n = collectNums(a); return isErr(n) ? n : n.reduce((s, x) => s * x, 1); },

    LEFT: (s, n) => { s = toStr(s); n = n === null || n === undefined ? 1 : toNum(n); if (isErr(s)) return s; if (isErr(n)) return n; if (n < 0) return E.VALUE; return [...s].slice(0, Math.trunc(n)).join(''); },
    RIGHT: (s, n) => { s = toStr(s); n = n === null || n === undefined ? 1 : toNum(n); if (isErr(s)) return s; if (isErr(n)) return n; if (n < 0) return E.VALUE; const a = [...s]; n = Math.trunc(n); return n === 0 ? '' : a.slice(-n).join(''); },
    MID: (s, st, n) => { s = toStr(s); st = toNum(st); n = toNum(n); if (isErr(s)) return s; if (isErr(st)) return st; if (isErr(n)) return n; if (st < 1 || n < 0) return E.VALUE; return [...s].slice(Math.trunc(st) - 1, Math.trunc(st) - 1 + Math.trunc(n)).join(''); },
    LEN: (s) => { s = toStr(s); return isErr(s) ? s : [...s].length; },
    VALUE: (s) => { const v = scalar(s); if (typeof v === 'number') return v; return toNum(toStr(v)); },
    TRIM: (s) => { s = toStr(s); return isErr(s) ? s : s.trim().replace(/ +/g, ' '); },
    UPPER: (s) => { s = toStr(s); return isErr(s) ? s : s.toUpperCase(); },
    LOWER: (s) => { s = toStr(s); return isErr(s) ? s : s.toLowerCase(); },
    PROPER: (s) => { s = toStr(s); return isErr(s) ? s : s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()); },
    CONCATENATE: (...a) => { let s = ''; for (const x of a) { const t = toStr(x); if (isErr(t)) return t; s += t; } return s; },
    CONCAT: (...a) => F.CONCATENATE(...a.flatMap((x) => (x instanceof Range ? x.flat() : [x]))),
    TEXT: (v, fmt) => {
      const n = toNum(v); fmt = toStr(fmt);
      if (isErr(n)) return toStr(v);
      if (/d.*m.*y/i.test(fmt)) return formatDate(n);
      if (fmt.includes('#,##0') || fmt.includes('#.##0')) return formatNumber(n, 'number');
      if (fmt === '0') return String(Math.round(n));
      return fmtGeneral(n);
    },
    EXACT: (a, b) => toStr(a) === toStr(b),
    FIND: (f, s, st) => { f = toStr(f); s = toStr(s); st = st == null ? 1 : toNum(st); const i = s.indexOf(f, st - 1); return i < 0 ? E.VALUE : i + 1; },
    SEARCH: (f, s, st) => { f = toStr(f); s = toStr(s); st = st == null ? 1 : toNum(st); const i = s.toLowerCase().indexOf(f.toLowerCase(), st - 1); return i < 0 ? E.VALUE : i + 1; },

    AND: (...a) => { let r = true; for (const x of a.flatMap((v) => (v instanceof Range ? v.flat().filter((y) => y !== null) : [v]))) { const b = toBool(x); if (isErr(b)) return b; r = r && b; } return r; },
    OR: (...a) => { let r = false; for (const x of a.flatMap((v) => (v instanceof Range ? v.flat().filter((y) => y !== null) : [v]))) { const b = toBool(x); if (isErr(b)) return b; r = r || b; } return r; },
    NOT: (x) => { const b = toBool(x); return isErr(b) ? b : !b; },

    SUMIF: (rng, crit, sumRng) => {
      rng = asRange(rng); const s = sumRng instanceof Range ? sumRng : rng; const ok = makeCriteria(crit);
      let t = 0;
      rng.rows.forEach((row, i) => row.forEach((v, j) => { if (ok(v)) { const x = s.rows[i] && s.rows[i][j]; if (typeof x === 'number') t += x; } }));
      return t;
    },
    COUNTIF: (rng, crit) => { const ok = makeCriteria(crit); return asRange(rng).flat().filter(ok).length; },
    AVERAGEIF: (rng, crit, avgRng) => {
      rng = asRange(rng); const s = avgRng instanceof Range ? avgRng : rng; const ok = makeCriteria(crit);
      let t = 0, n = 0;
      rng.rows.forEach((row, i) => row.forEach((v, j) => { if (ok(v)) { const x = s.rows[i] && s.rows[i][j]; if (typeof x === 'number') { t += x; n++; } } }));
      return n ? t / n : E.DIV0;
    },
    SUMIFS: (sumRng, ...pairs) => multiIf(pairs, sumRng, 'sum'),
    COUNTIFS: (...pairs) => multiIf(pairs, null, 'count'),
    AVERAGEIFS: (avgRng, ...pairs) => multiIf(pairs, avgRng, 'avg'),

    VLOOKUP: (val, table, col, mode) => lookup(val, table, col, mode, false),
    HLOOKUP: (val, table, row, mode) => lookup(val, table, row, mode, true),
    INDEX: (table, r, c) => {
      table = asRange(table); r = toNum(r); c = c == null ? 1 : toNum(c);
      if (isErr(r)) return r; if (isErr(c)) return c;
      if (table.h === 1 && (c === 1 || c === 0) && r > 1) { c = r; r = 1; }
      const row = table.rows[r - 1];
      if (!row || row[c - 1] === undefined) return E.REF;
      return row[c - 1];
    },
    MATCH: (val, rng, type) => {
      val = scalar(val); const arr = asRange(rng).flat(); type = type == null ? 1 : toNum(type);
      if (type === 0) { const i = arr.findIndex((v) => lookupEq(val, v)); return i < 0 ? E.NA : i + 1; }
      let best = -1;
      for (let i = 0; i < arr.length; i++) {
        const c = compare(arr[i], val);
        if (type > 0 ? c <= 0 : c >= 0) best = i; else break;
      }
      return best < 0 ? E.NA : best + 1;
    },
    RANK: (n, rng, order) => {
      n = toNum(n); if (isErr(n)) return n;
      const nums = asRange(rng).flat().filter((v) => typeof v === 'number');
      if (!nums.includes(n)) return E.NA;
      const asc = order != null && toNum(order) !== 0;
      return 1 + nums.filter((v) => (asc ? v < n : v > n)).length;
    },

    TODAY: () => todaySerial(),
    NOW: () => { const t = new Date(); return todaySerial() + (t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds()) / 86400; },
    DATE: (y, m, d) => { y = toNum(y); m = toNum(m); d = toNum(d); if ([y, m, d].some(isErr)) return E.VALUE; if (y < 100) y += 1900; return dateSerial(y, m, d); },
    DAY: (n) => { n = toNum(n); return isErr(n) ? n : serialToDate(n).getUTCDate(); },
    MONTH: (n) => { n = toNum(n); return isErr(n) ? n : serialToDate(n).getUTCMonth() + 1; },
    YEAR: (n) => { n = toNum(n); return isErr(n) ? n : serialToDate(n).getUTCFullYear(); },
    WEEKDAY: (n, t) => { n = toNum(n); if (isErr(n)) return n; const d = serialToDate(n).getUTCDay(); t = t == null ? 1 : toNum(t); return t === 2 ? ((d + 6) % 7) + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
    DAYS: (a, b) => { a = toNum(a); b = toNum(b); return isErr(a) ? a : isErr(b) ? b : Math.floor(a) - Math.floor(b); },
    ISBLANK: (v) => scalar(v) === null,
    ISNUMBER: (v) => typeof scalar(v) === 'number',
    ISTEXT: (v) => typeof scalar(v) === 'string',
    ISERROR: (v) => isErr(scalar(v)),
    ISNA: (v) => scalar(v) === E.NA,
  };

  function multiIf(args, valRng, mode) {
    if (args.length % 2 !== 0) return E.VALUE;
    const conds = [];
    for (let i = 0; i < args.length; i += 2) conds.push([asRange(args[i]), makeCriteria(args[i + 1])]);
    const base = conds[0][0];
    const vals = valRng ? asRange(valRng) : null;
    if (conds.some(([r]) => r.h !== base.h || r.w !== base.w)) return E.VALUE;
    let t = 0, n = 0;
    for (let i = 0; i < base.h; i++) {
      for (let j = 0; j < base.w; j++) {
        if (!conds.every(([r, ok]) => ok(r.rows[i][j]))) continue;
        if (mode === 'count') { n++; continue; }
        const x = vals.rows[i] && vals.rows[i][j];
        if (typeof x === 'number') { t += x; n++; }
      }
    }
    if (mode === 'count') return n;
    if (mode === 'avg') return n ? t / n : E.DIV0;
    return t;
  }

  function lookup(val, table, idx, mode, horizontal) {
    val = scalar(val);
    if (isErr(val)) return val;
    if (!(table instanceof Range)) return E.VALUE;
    idx = toNum(idx);
    if (isErr(idx)) return idx;
    idx = Math.trunc(idx);
    const exact = mode !== null && mode !== undefined && mode !== '' && !toBool(mode);
    const n = horizontal ? table.w : table.h;
    const span = horizontal ? table.h : table.w;
    if (idx < 1) return E.VALUE;
    if (idx > span) return E.REF;
    const key = (i) => (horizontal ? table.rows[0][i] : table.rows[i][0]);
    const get = (i) => { const v = horizontal ? table.rows[idx - 1][i] : table.rows[i][idx - 1]; return v === null ? 0 : v; };
    if (exact) {
      for (let i = 0; i < n; i++) if (lookupEq(val, key(i))) return get(i);
      return E.NA;
    }
    let best = -1;
    for (let i = 0; i < n; i++) {
      const k = key(i);
      if (k === null) continue;
      if ((typeof k === 'number') !== (typeof val === 'number')) continue;
      if (compare(k, val) <= 0) best = i; else break;
    }
    return best < 0 ? E.NA : get(best);
  }

  const LAZY = {
    IF(args, ev) {
      if (args.length < 1 || args.length > 3) return E.VALUE;
      const c = toBool(ev(args[0]));
      if (isErr(c)) return c;
      if (c) return args.length > 1 ? ev(args[1]) ?? 0 : true;
      return args.length > 2 ? ev(args[2]) ?? 0 : false;
    },
    IFERROR(args, ev) {
      const v = scalar(ev(args[0]));
      return isErr(v) ? ev(args[1]) ?? 0 : v;
    },
    IFS(args, ev) {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const c = toBool(ev(args[i]));
        if (isErr(c)) return c;
        if (c) return ev(args[i + 1]);
      }
      return E.NA;
    },
  };

  // Số tham số tối thiểu/tối đa để báo lỗi giống Excel.
  const ARITY = {
    IF: [1, 3], VLOOKUP: [3, 4], HLOOKUP: [3, 4], LEFT: [1, 2], RIGHT: [1, 2], MID: [3, 3], SUMIF: [2, 3],
    COUNTIF: [2, 2], AVERAGEIF: [2, 3], RANK: [2, 3], ROUND: [2, 2], MOD: [2, 2], DATE: [3, 3], INDEX: [2, 3], MATCH: [2, 3],
  };

  function evaluate(node, ctx) {
    const ev = (n) => evaluate(n, ctx);
    switch (node.k) {
      case 'lit': return node.v;
      case 'empty': return null;
      case 'badname': return E.NAME;
      case 'ref': return ctx.get(node.r, node.c);
      case 'range': {
        const rows = [];
        for (let r = node.r1; r <= node.r2; r++) {
          const row = [];
          for (let c = node.c1; c <= node.c2; c++) row.push(ctx.get(r, c));
          rows.push(row);
        }
        return new Range(rows);
      }
      case 'neg': { const v = toNum(ev(node.a)); return isErr(v) ? v : -v; }
      case 'pct': { const v = toNum(ev(node.a)); return isErr(v) ? v : v / 100; }
      case 'bin': {
        const a = scalar(ev(node.a)), b = scalar(ev(node.b));
        if (isErr(a)) return a;
        if (isErr(b)) return b;
        if (node.op === '&') return toStr(a) + toStr(b);
        if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
          const c = compare(a, b);
          return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[node.op];
        }
        const x = toNum(a), y = toNum(b);
        if (isErr(x)) return x;
        if (isErr(y)) return y;
        switch (node.op) {
          case '+': return x + y;
          case '-': return x - y;
          case '*': return x * y;
          case '/': return y === 0 ? E.DIV0 : x / y;
          case '^': return Math.pow(x, y);
        }
        return E.VALUE;
      }
      case 'call': {
        const name = node.name;
        const ar = ARITY[name];
        if (ar && (node.args.length < ar[0] || node.args.length > ar[1])) return E.VALUE;
        if (LAZY[name]) return LAZY[name](node.args, ev);
        const fn = F[name];
        if (!fn) return E.NAME;
        const args = node.args.map(ev);
        for (const a of args) if (isErr(a) && !['ISERROR', 'ISNA', 'COUNT', 'COUNTA'].includes(name)) return a;
        const r = fn(...args);
        return r === undefined ? E.VALUE : r;
      }
    }
    return E.VALUE;
  }

  // ---------- Dịch tham chiếu khi sao chép / điền ----------
  function shiftFormula(src, dr, dc) {
    if (!src.startsWith('=')) return src;
    let toks;
    try { toks = tokenize(src.slice(1)); } catch (e) { return src; }
    let out = src.slice(1);
    for (let i = toks.length - 1; i >= 0; i--) {
      const t = toks[i];
      if (t.t !== 'ref') continue;
      const r = t.absR ? t.row : t.row + dr;
      const c = t.absC ? t.col : t.col + dc;
      const txt = r < 0 || c < 0 ? '#REF!' : (t.absC ? '$' : '') + numToCol(c) + (t.absR ? '$' : '') + (r + 1);
      out = out.slice(0, t.s) + txt + out.slice(t.e);
    }
    return '=' + out;
  }

  // F4: xoay vòng A1 -> $A$1 -> A$1 -> $A1 -> A1 cho tham chiếu tại con trỏ.
  function toggleAbs(src, caret) {
    if (!src.startsWith('=')) return null;
    let toks;
    try { toks = tokenize(src.slice(1)); } catch (e) { return null; }
    const pos = caret - 1;
    const t = toks.find((k) => k.t === 'ref' && pos >= k.s && pos <= k.e) ||
      [...toks].reverse().find((k) => k.t === 'ref' && k.e <= pos);
    if (!t) return null;
    const state = (t.absC ? 2 : 0) + (t.absR ? 1 : 0);
    const next = { 0: [true, true], 3: [false, true], 1: [true, false], 2: [false, false] }[state];
    const fmt = (k) => (next[0] ? '$' : '') + numToCol(k.col) + (next[1] ? '$' : '') + (k.row + 1);
    // Con trỏ ở vùng A1:B2 → đổi cả hai đầu vùng.
    const i = toks.indexOf(t);
    let first = t, last = t;
    if (toks[i - 1] && toks[i - 1].t === ':' && toks[i - 2] && toks[i - 2].t === 'ref') first = toks[i - 2];
    else if (toks[i + 1] && toks[i + 1].t === ':' && toks[i + 2] && toks[i + 2].t === 'ref') last = toks[i + 2];
    const body = src.slice(1);
    const txt = first === last ? fmt(t) : fmt(first) + ':' + fmt(last);
    const nb = body.slice(0, first.s) + txt + body.slice(last.e);
    return { text: '=' + nb, caret: 1 + first.s + txt.length };
  }

  // Lấy danh sách tham chiếu (để tô màu khi đang gõ công thức).
  function refsIn(src) {
    if (!src.startsWith('=')) return [];
    let toks;
    try { toks = tokenize(src.slice(1)); } catch (e) { return []; }
    const out = [];
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.t !== 'ref') continue;
      if (toks[i + 1] && toks[i + 1].t === ':' && toks[i + 2] && toks[i + 2].t === 'ref') {
        const b = toks[i + 2];
        out.push({ r1: Math.min(t.row, b.row), c1: Math.min(t.col, b.col), r2: Math.max(t.row, b.row), c2: Math.max(t.col, b.col) });
        i += 2;
      } else out.push({ r1: t.row, c1: t.col, r2: t.row, c2: t.col });
    }
    return out;
  }

  // Tìm hàm đang gõ + vị trí tham số tại con trỏ (để hiện gợi ý cú pháp).
  function callAt(src, caret) {
    const s = src.slice(0, caret);
    const stack = [];
    let inStr = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (ch === '(') {
        const m = /([A-Za-z][A-Za-z0-9.]*)\s*$/.exec(s.slice(0, i));
        stack.push({ name: m ? m[1].toUpperCase() : null, arg: 0 });
      } else if (ch === ')') stack.pop();
      else if ((ch === ',' || ch === ';') && stack.length) stack[stack.length - 1].arg++;
    }
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i];
    return null;
  }

  function usedFunctions(src) {
    if (!src || !src.startsWith('=')) return [];
    try { return tokenize(src.slice(1)).filter((t) => t.t === 'func').map((t) => t.v); } catch (e) { return []; }
  }

  function formatNumber(n, fmt) {
    const group = (x, dec) => {
      const neg = x < 0;
      const fixed = Math.abs(x).toFixed(dec);
      const [i, d] = fixed.split('.');
      return (neg ? '-' : '') + i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? '.' + d : '');
    };
    switch (fmt) {
      case 'number': return group(roundHalfUp(n, 0), 0);
      case 'number2': return group(n, 2);
      case 'vnd': return group(roundHalfUp(n, 0), 0) + ' ₫';
      case 'usd': return '$' + group(n, Number.isInteger(n) ? 0 : 2);
      case 'percent': return fmtGeneral(roundHalfUp(n * 100, 2)) + '%';
      case 'date': return formatDate(n);
      default: return fmtGeneral(n);
    }
  }

  const api = {
    XLError, E, isErr, Range, parse, tokenize, evaluate, shiftFormula, toggleAbs, refsIn, callAt, usedFunctions,
    colToNum, numToCol, addr, parseAddr, parseRange, dateSerial, formatDate, parseDateText, formatNumber, fmtGeneral,
    toStr, toNum, compare, FUNCS: F, LAZY,
  };
  global.XL = api;
  if (typeof module !== 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
