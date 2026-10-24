/* Giao diện bảng tính kiểu Excel: chọn ô, sửa, trỏ chuột chèn tham chiếu, F4, kéo điền, Ctrl+D/R/Enter, sao chép-dán, hoàn tác. */
(function (global) {
  'use strict';
  const XL = global.XL;

  const FUNC_HELP = {
    SUM: ['number1', '[number2]', '…'], AVERAGE: ['number1', '[number2]', '…'], MAX: ['number1', '[number2]', '…'], MIN: ['number1', '[number2]', '…'],
    COUNT: ['value1', '[value2]', '…'], COUNTA: ['value1', '[value2]', '…'],
    IF: ['điều_kiện', 'giá_trị_nếu_đúng', 'giá_trị_nếu_sai'], AND: ['đk1', '[đk2]', '…'], OR: ['đk1', '[đk2]', '…'], NOT: ['đk'],
    IFERROR: ['giá_trị', 'giá_trị_nếu_lỗi'],
    LEFT: ['chuỗi', '[số_ký_tự]'], RIGHT: ['chuỗi', '[số_ký_tự]'], MID: ['chuỗi', 'vị_trí_bắt_đầu', 'số_ký_tự'], LEN: ['chuỗi'],
    VALUE: ['chuỗi'], TRIM: ['chuỗi'], UPPER: ['chuỗi'], LOWER: ['chuỗi'], PROPER: ['chuỗi'], CONCATENATE: ['chuỗi1', '[chuỗi2]', '…'],
    VLOOKUP: ['giá_trị_dò', 'bảng_dò', 'cột_trả_về', '[0 = dò chính xác]'], HLOOKUP: ['giá_trị_dò', 'bảng_dò', 'hàng_trả_về', '[0 = dò chính xác]'],
    INDEX: ['vùng', 'hàng', '[cột]'], MATCH: ['giá_trị_dò', 'vùng_dò', '[0 = chính xác]'],
    SUMIF: ['vùng_điều_kiện', 'điều_kiện', '[vùng_tính_tổng]'], COUNTIF: ['vùng', 'điều_kiện'], AVERAGEIF: ['vùng_điều_kiện', 'điều_kiện', '[vùng_tính_TB]'],
    SUMIFS: ['vùng_tính_tổng', 'vùng_đk1', 'đk1', '[vùng_đk2]', '[đk2]', '…'], COUNTIFS: ['vùng_đk1', 'đk1', '[vùng_đk2]', '[đk2]', '…'],
    AVERAGEIFS: ['vùng_tính_TB', 'vùng_đk1', 'đk1', '…'],
    RANK: ['số', 'vùng_so_sánh', '[0 = giảm dần]'], ROUND: ['số', 'số_chữ_số_thập_phân'], ROUNDUP: ['số', 'số_chữ_số'], ROUNDDOWN: ['số', 'số_chữ_số'],
    INT: ['số'], MOD: ['số_bị_chia', 'số_chia'], ABS: ['số'], SQRT: ['số'], POWER: ['cơ_số', 'số_mũ'],
    DATE: ['năm', 'tháng', 'ngày'], DAY: ['ngày_tháng'], MONTH: ['ngày_tháng'], YEAR: ['ngày_tháng'], TODAY: [], NOW: [], WEEKDAY: ['ngày_tháng', '[kiểu]'],
  };
  const FUNC_NAMES = Object.keys(XL.FUNCS).concat(Object.keys(XL.LAZY)).sort();

  const el = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt !== undefined) e.textContent = txt; return e; };
  const k = (r, c) => r + ',' + c;

  class SheetView {
    constructor(host, sheet, opts) {
      this.host = host;
      this.opts = opts || {};
      this.sheet = sheet;
      this.sel = { r1: 0, c1: 0, r2: 0, c2: 0, ar: 0, ac: 0 };
      this.edit = null;
      this.point = null;
      this.undoStack = [];
      this.redoStack = [];
      this.marks = new Map();
      this.build();
      this.refresh();
    }

    // ---------- Dựng DOM ----------
    build() {
      const root = el('div', 'xl');
      this.root = root;

      const tb = el('div', 'xl-toolbar');
      const btn = (label, title, fn, cls) => { const b = el('button', 'xl-btn ' + (cls || ''), label); b.type = 'button'; b.title = title; b.addEventListener('mousedown', (e) => e.preventDefault()); b.addEventListener('click', fn); tb.appendChild(b); return b; };
      const sep = () => tb.appendChild(el('span', 'xl-sep'));
      btn('↶', 'Hoàn tác (Ctrl+Z)', () => this.undo());
      btn('↷', 'Làm lại (Ctrl+Y)', () => this.redoAct());
      sep();
      tb.appendChild(el('span', 'xl-tlabel', 'Định dạng:'));
      btn('Chung', 'General', () => this.setFormat('general'));
      btn('1,234', 'Số có phân cách hàng nghìn (#,##0)', () => this.setFormat('number'));
      btn('VNĐ', 'Tiền Việt Nam đồng', () => this.setFormat('vnd'));
      btn('$', 'Tiền USD', () => this.setFormat('usd'));
      btn('%', 'Phần trăm', () => this.setFormat('percent'));
      btn('Ngày', 'dd/mm/yyyy', () => this.setFormat('date'));
      sep();
      btn('Σ', 'AutoSum (Alt+=)', () => this.autoSum());
      btn('⬇ Điền xuống', 'Fill Down (Ctrl+D)', () => this.fillDown());
      btn('➡ Điền phải', 'Fill Right (Ctrl+R)', () => this.fillRight());
      btn('✕ Xóa', 'Xóa nội dung (Delete)', () => this.clearSel());
      root.appendChild(tb);

      const fb = el('div', 'xl-fbar');
      this.nameBox = el('div', 'xl-name', 'A1');
      fb.appendChild(this.nameBox);
      fb.appendChild(el('span', 'xl-fx', 'fx'));
      this.fInput = el('input', 'xl-finput');
      this.fInput.spellcheck = false;
      this.fInput.setAttribute('aria-label', 'Thanh công thức');
      fb.appendChild(this.fInput);
      root.appendChild(fb);

      const wrap = el('div', 'xl-wrap');
      wrap.tabIndex = 0;
      this.wrap = wrap;
      const inner = el('div', 'xl-inner');
      this.inner = inner;
      const table = el('table', 'xl-grid');
      this.table = table;
      const cg = el('colgroup');
      const c0 = el('col'); c0.style.width = '38px'; cg.appendChild(c0);
      const widths = this.opts.widths || [];
      let tw = 38;
      for (let c = 0; c < this.sheet.nCols; c++) { const w = widths[c] || 84; tw += w; const col = el('col'); col.style.width = w + 'px'; cg.appendChild(col); }
      table.style.width = tw + 'px';
      table.appendChild(cg);
      const thead = el('thead');
      const hr = el('tr');
      hr.appendChild(el('th', 'xl-corner'));
      this.colHeads = [];
      for (let c = 0; c < this.sheet.nCols; c++) { const th = el('th', 'xl-ch', XL.numToCol(c)); hr.appendChild(th); this.colHeads.push(th); }
      thead.appendChild(hr);
      table.appendChild(thead);
      const tbody = el('tbody');
      this.tds = [];
      this.rowHeads = [];
      for (let r = 0; r < this.sheet.nRows; r++) {
        const tr = el('tr');
        const th = el('th', 'xl-rh', String(r + 1));
        tr.appendChild(th);
        this.rowHeads.push(th);
        const row = [];
        for (let c = 0; c < this.sheet.nCols; c++) {
          const td = el('td');
          td.dataset.r = r; td.dataset.c = c;
          tr.appendChild(td);
          row.push(td);
        }
        this.tds.push(row);
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      inner.appendChild(table);
      this.selBox = el('div', 'xl-selbox');
      this.handle = el('div', 'xl-handle');
      this.handle.title = 'Kéo để điền (Fill)';
      this.fillBox = el('div', 'xl-fillbox');
      this.refLayer = el('div', 'xl-reflayer');
      inner.appendChild(this.refLayer);
      inner.appendChild(this.selBox);
      inner.appendChild(this.fillBox);
      inner.appendChild(this.handle);
      this.hintBox = el('div', 'xl-hint');
      this.acBox = el('div', 'xl-ac');
      inner.appendChild(this.hintBox);
      inner.appendChild(this.acBox);
      this.toastBox = el('div', 'xl-toast');
      root.appendChild(this.toastBox);
      wrap.appendChild(inner);
      root.appendChild(wrap);
      this.host.innerHTML = '';
      this.host.appendChild(root);
      this.bind();
    }

    bind() {
      const cellOf = (e) => { const td = e.target.closest && e.target.closest('td'); return td && td.dataset.r !== undefined ? { r: +td.dataset.r, c: +td.dataset.c } : null; };
      this.table.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        const hd = e.target.closest('th');
        if (hd && !this.edit) {
          e.preventDefault();
          const ci = this.colHeads.indexOf(hd), ri = this.rowHeads.indexOf(hd);
          if (ci >= 0) this.select(0, ci, this.sheet.nRows - 1, ci, 0, ci);
          if (ri >= 0) this.select(ri, 0, ri, this.sheet.nCols - 1, ri, 0);
          this.wrap.focus({ preventScroll: true });
          return;
        }
        const p = cellOf(e);
        if (!p) return;
        e.preventDefault();
        if (this.edit && this.isPointable()) { this.pointAt(p.r, p.c); this.dragPoint = true; return; }
        if (this.edit) this.commit(null);
        if (e.shiftKey) this.select(this.sel.ar, this.sel.ac, p.r, p.c, this.sel.ar, this.sel.ac);
        else this.select(p.r, p.c, p.r, p.c, p.r, p.c);
        this.dragSel = true;
        this.wrap.focus({ preventScroll: true });
      });
      this.table.addEventListener('mouseover', (e) => {
        const p = cellOf(e);
        if (!p) return;
        if (this.dragPoint && this.edit && this.point) this.pointExtend(p.r, p.c);
        else if (this.dragSel) this.select(this.sel.ar, this.sel.ac, p.r, p.c, this.sel.ar, this.sel.ac);
        else if (this.dragFill) this.fillPreview(p.r, p.c);
      });
      this.table.addEventListener('dblclick', (e) => { const p = cellOf(e); if (p) { this.select(p.r, p.c, p.r, p.c, p.r, p.c); this.startEdit(null, false); } });
      this.handle.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); if (this.edit) this.commit(null); this.dragFill = true; this.fillTarget = null; });
      this.onUp = () => {
        if (this.dragFill) { this.dragFill = false; this.fillBox.style.display = 'none'; if (this.fillTarget) this.fillTo(this.fillTarget); }
        this.dragSel = false; this.dragPoint = false;
      };
      document.addEventListener('mouseup', this.onUp);
      this.wrap.addEventListener('keydown', (e) => this.onGridKey(e));
      this.wrap.addEventListener('copy', (e) => { if (!this.edit) { this.copy(e); } });
      this.wrap.addEventListener('paste', (e) => { if (!this.edit) { this.paste(e); } });
      this.fInput.addEventListener('focus', () => { if (!this.edit) this.startEdit(null, false, true); });
      this.fInput.addEventListener('input', () => { if (this.edit) { this.edit.input.value = this.fInput.value; this.point = null; this.onEditInput(); } });
      this.fInput.addEventListener('keydown', (e) => this.onEditKey(e, this.fInput));
      this.fInput.addEventListener('keyup', () => this.onEditInput());
      this.fInput.addEventListener('blur', () => setTimeout(() => {
        if (this.edit && document.activeElement !== this.edit.input && document.activeElement !== this.fInput && !this.dragPoint) this.commit(null);
      }, 0));
      this.wrap.addEventListener('scroll', () => this.positionEditor());
    }

    destroy() { document.removeEventListener('mouseup', this.onUp); }

    // ---------- Hiển thị ----------
    refresh() {
      const s = this.sheet;
      s.compute();
      for (let r = 0; r < s.nRows; r++) {
        for (let c = 0; c < s.nCols; c++) {
          const td = this.tds[r][c];
          const x = s.cell(r, c);
          const v = s.value(r, c);
          td.textContent = s.display(r, c);
          let cls = '';
          if (x && x.cls) cls += x.cls + ' ';
          if (x && x.task) cls += 'x-task ';
          if (typeof v === 'number') cls += 'x-num ';
          else if (XL.isErr(v)) cls += 'x-err ';
          else if (typeof v === 'boolean') cls += 'x-bool ';
          const m = this.marks.get(k(r, c));
          if (m) cls += m.ok ? 'x-ok ' : 'x-bad ';
          td.className = cls.trim();
          td.title = (x && x.parseError) ? 'Lỗi công thức: ' + x.parseError : (m && m.tip) || '';
        }
      }
      this.updateSelUI();
    }

    setMarks(map) { this.marks = map || new Map(); this.refresh(); }

    rectOf(r1, c1, r2, c2) {
      const a = this.tds[r1][c1], b = this.tds[r2][c2];
      return { left: a.offsetLeft, top: a.offsetTop, width: b.offsetLeft + b.offsetWidth - a.offsetLeft, height: b.offsetTop + b.offsetHeight - a.offsetTop };
    }
    place(box, r) { Object.assign(box.style, { left: r.left + 'px', top: r.top + 'px', width: r.width + 'px', height: r.height + 'px', display: 'block' }); }

    updateSelUI() {
      const s = this.sel;
      const r1 = Math.min(s.r1, s.r2), r2 = Math.max(s.r1, s.r2), c1 = Math.min(s.c1, s.c2), c2 = Math.max(s.c1, s.c2);
      const rect = this.rectOf(r1, c1, r2, c2);
      this.place(this.selBox, rect);
      this.selBox.classList.toggle('multi', r1 !== r2 || c1 !== c2);
      Object.assign(this.handle.style, { left: rect.left + rect.width - 4 + 'px', top: rect.top + rect.height - 4 + 'px', display: this.edit ? 'none' : 'block' });
      for (const td of this.tds.flat()) td.classList.remove('x-insel', 'x-active');
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) this.tds[r][c].classList.add('x-insel');
      this.tds[s.ar][s.ac].classList.add('x-active');
      this.colHeads.forEach((h, i) => h.classList.toggle('on', i >= c1 && i <= c2));
      this.rowHeads.forEach((h, i) => h.classList.toggle('on', i >= r1 && i <= r2));
      this.nameBox.textContent = r1 === r2 && c1 === c2 ? XL.addr(s.ar, s.ac) : XL.addr(r1, c1) + ':' + XL.addr(r2, c2);
      if (!this.edit) this.fInput.value = this.sheet.input(s.ar, s.ac);
    }

    norm() {
      const s = this.sel;
      return { r1: Math.min(s.r1, s.r2), r2: Math.max(s.r1, s.r2), c1: Math.min(s.c1, s.c2), c2: Math.max(s.c1, s.c2) };
    }

    select(r1, c1, r2, c2, ar, ac) {
      const cl = (v, n) => Math.max(0, Math.min(n - 1, v));
      this.sel = { r1: cl(r1, this.sheet.nRows), c1: cl(c1, this.sheet.nCols), r2: cl(r2, this.sheet.nRows), c2: cl(c2, this.sheet.nCols), ar: cl(ar, this.sheet.nRows), ac: cl(ac, this.sheet.nCols) };
      this.updateSelUI();
      this.scrollTo(this.sel.r2, this.sel.c2);
      if (this.opts.onSelect) this.opts.onSelect(this.sel);
    }

    scrollTo(r, c) {
      const td = this.tds[r][c];
      const w = this.wrap;
      if (td.offsetTop < w.scrollTop + 26) w.scrollTop = td.offsetTop - 26;
      else if (td.offsetTop + td.offsetHeight > w.scrollTop + w.clientHeight) w.scrollTop = td.offsetTop + td.offsetHeight - w.clientHeight;
      if (td.offsetLeft < w.scrollLeft + 38) w.scrollLeft = td.offsetLeft - 38;
      else if (td.offsetLeft + td.offsetWidth > w.scrollLeft + w.clientWidth) w.scrollLeft = td.offsetLeft + td.offsetWidth - w.clientWidth;
    }

    toast(msg) {
      this.toastBox.textContent = msg;
      this.toastBox.classList.add('show');
      clearTimeout(this.toastT);
      this.toastT = setTimeout(() => this.toastBox.classList.remove('show'), 2600);
    }

    // ---------- Phím trên lưới ----------
    onGridKey(e) {
      if (this.edit) return;
      const s = this.sel;
      const ctrl = e.ctrlKey || e.metaKey;
      const move = (dr, dc) => {
        e.preventDefault();
        if (e.shiftKey) this.select(s.ar, s.ac, s.r2 + dr, s.c2 + dc, s.ar, s.ac);
        else { const r = s.ar + dr, c = s.ac + dc; this.select(r, c, r, c, r, c); }
      };
      if (ctrl) {
        const key = e.key.toLowerCase();
        if (key === 'd') { e.preventDefault(); this.fillDown(); return; }
        if (key === 'r') { e.preventDefault(); this.fillRight(); return; }
        if (key === 'z') { e.preventDefault(); this.undo(); return; }
        if (key === 'y') { e.preventDefault(); this.redoAct(); return; }
        if (key === 'a') { e.preventDefault(); this.select(0, 0, this.sheet.nRows - 1, this.sheet.nCols - 1, 0, 0); return; }
        if (e.key === 'Home') { e.preventDefault(); this.select(0, 0, 0, 0, 0, 0); return; }
        return;
      }
      if (e.altKey && (e.key === '=' || e.code === 'Equal')) { e.preventDefault(); this.autoSum(); return; }
      switch (e.key) {
        case 'ArrowUp': return move(-1, 0);
        case 'ArrowDown': return move(1, 0);
        case 'ArrowLeft': return move(0, -1);
        case 'ArrowRight': return move(0, 1);
        case 'Enter': return move(e.shiftKey ? -1 : 1, 0);
        case 'Tab': e.preventDefault(); { const c = s.ac + (e.shiftKey ? -1 : 1); this.select(s.ar, c, s.ar, c, s.ar, c); } return;
        case 'F2': e.preventDefault(); this.startEdit(null, false); return;
        case 'Delete': e.preventDefault(); this.clearSel(); return;
        case 'Backspace': e.preventDefault(); if (this.lockedMsg()) return; this.startEdit('', true); return;
        case 'Escape': return;
      }
      if (e.key.length === 1 && !e.altKey) {
        if (this.lockedMsg()) { e.preventDefault(); return; }
        // Không chặn phím: ký tự sẽ được gõ vào ô sửa vừa nhận focus (hỗ trợ cả bộ gõ tiếng Việt).
        this.startEdit('', true);
      }
    }

    lockedMsg() {
      const x = this.sheet.cell(this.sel.ar, this.sel.ac);
      if (x && x.locked) { this.toast('Ô ' + XL.addr(this.sel.ar, this.sel.ac) + ' là dữ liệu đề bài (đã khóa). Hãy nhập ở ô màu vàng.'); return true; }
      return false;
    }

    // ---------- Sửa ô ----------
    startEdit(text, enterMode, inFbar) {
      const { ar, ac } = this.sel;
      if (this.lockedMsg()) { if (inFbar) this.wrap.focus(); return; }
      const input = el('input', 'xl-editor');
      input.spellcheck = false;
      input.value = text === null || text === undefined ? this.sheet.input(ar, ac) : text;
      this.inner.appendChild(input);
      this.edit = { r: ar, c: ac, input, enterMode };
      this.point = null;
      this.positionEditor();
      this.handle.style.display = 'none';
      input.addEventListener('input', () => { this.point = null; this.fInput.value = input.value; this.onEditInput(); });
      input.addEventListener('keydown', (e) => this.onEditKey(e, input));
      input.addEventListener('keyup', () => this.onEditInput(true));
      input.addEventListener('click', () => this.onEditInput(true));
      input.addEventListener('blur', () => setTimeout(() => {
        if (this.edit && this.edit.input === input && document.activeElement !== this.fInput && document.activeElement !== input && !this.dragPoint) this.commit(null);
      }, 0));
      this.fInput.value = input.value;
      if (inFbar) { this.edit.active = this.fInput; }
      else { this.edit.active = input; input.focus(); const n = input.value.length; input.setSelectionRange(n, n); }
      this.onEditInput();
    }

    activeInput() { return this.edit ? (document.activeElement === this.fInput ? this.fInput : this.edit.input) : null; }

    positionEditor() {
      if (!this.edit) return;
      const td = this.tds[this.edit.r][this.edit.c];
      const inp = this.edit.input;
      Object.assign(inp.style, { left: td.offsetLeft + 'px', top: td.offsetTop + 'px', height: td.offsetHeight + 'px', minWidth: td.offsetWidth + 'px' });
      inp.style.width = Math.max(td.offsetWidth, Math.min(560, inp.value.length * 8 + 24)) + 'px';
    }

    onEditInput() {
      if (!this.edit) return;
      const inp = this.activeInput();
      const text = inp.value;
      this.positionEditor();
      this.drawRefs(text);
      this.showHint(text, inp.selectionStart || 0);
    }

    drawRefs(text) {
      this.refLayer.innerHTML = '';
      const refs = XL.refsIn(text).slice(0, 12);
      refs.forEach((g, i) => {
        if (g.r2 >= this.sheet.nRows || g.c2 >= this.sheet.nCols) return;
        const b = el('div', 'xl-ref c' + (i % 6));
        this.place(b, this.rectOf(g.r1, g.c1, g.r2, g.c2));
        this.refLayer.appendChild(b);
      });
    }

    showHint(text, caret) {
      const hb = this.hintBox, ac = this.acBox;
      hb.style.display = 'none'; ac.style.display = 'none';
      this.acList = null;
      if (!this.edit || !text.startsWith('=')) return;
      const td = this.tds[this.edit.r][this.edit.c];
      const below = td.offsetTop + td.offsetHeight + 2;
      const m = /(^|[=(,;+\-*/^&<> ])([A-Za-z]{1,12})$/.exec(text.slice(0, caret));
      if (m) {
        const pre = m[2].toUpperCase();
        const list = FUNC_NAMES.filter((n) => n.startsWith(pre)).slice(0, 7);
        if (list.length && !(list.length === 1 && list[0] === pre && text[caret] === '(')) {
          this.acList = { items: list, idx: 0, start: caret - m[2].length, end: caret };
          this.renderAc();
          Object.assign(ac.style, { left: td.offsetLeft + 'px', top: below + 'px', display: 'block' });
          return;
        }
      }
      const call = XL.callAt(text, caret);
      if (call && FUNC_HELP[call.name]) {
        const args = FUNC_HELP[call.name];
        const lastIsMore = args[args.length - 1] === '…';
        hb.innerHTML = '';
        hb.appendChild(el('b', 'fn', call.name));
        hb.appendChild(document.createTextNode('('));
        args.forEach((a, i) => {
          const active = i === call.arg || (lastIsMore && i >= args.length - 2 && call.arg >= args.length - 2 && i === Math.min(call.arg, args.length - 2));
          const span = el('span', active ? 'on' : '', a);
          hb.appendChild(span);
          if (i < args.length - 1) hb.appendChild(document.createTextNode(', '));
        });
        hb.appendChild(document.createTextNode(')'));
        Object.assign(hb.style, { left: td.offsetLeft + 'px', top: below + 'px', display: 'block' });
      }
    }

    renderAc() {
      const ac = this.acBox;
      ac.innerHTML = '';
      this.acList.items.forEach((n, i) => {
        const it = el('div', 'xl-ac-item' + (i === this.acList.idx ? ' on' : ''), n);
        it.addEventListener('mousedown', (e) => { e.preventDefault(); this.acList.idx = i; this.acceptAc(); });
        ac.appendChild(it);
      });
    }

    acceptAc() {
      const inp = this.activeInput();
      const a = this.acList;
      const name = a.items[a.idx];
      inp.value = inp.value.slice(0, a.start) + name + '(' + inp.value.slice(a.end);
      const pos = a.start + name.length + 1;
      inp.setSelectionRange(pos, pos);
      this.syncInputs(inp);
      this.onEditInput();
    }

    syncInputs(from) {
      if (!this.edit) return;
      if (from === this.fInput) this.edit.input.value = from.value; else this.fInput.value = from.value;
    }

    onEditKey(e, inp) {
      if (!this.edit) return;
      if (this.acList && this.acBox.style.display === 'block') {
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          const n = this.acList.items.length;
          this.acList.idx = (this.acList.idx + (e.key === 'ArrowDown' ? 1 : n - 1)) % n;
          this.renderAc();
          return;
        }
        if (e.key === 'Tab') { e.preventDefault(); this.acceptAc(); return; }
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.ctrlKey) this.commit('fillsel');
        else this.commit(e.shiftKey ? 'up' : 'down');
        return;
      }
      if (e.key === 'Tab') { e.preventDefault(); this.commit(e.shiftKey ? 'left' : 'right'); return; }
      if (e.key === 'Escape') { e.preventDefault(); this.cancel(); return; }
      if (e.key === 'F4') {
        e.preventDefault();
        const r = XL.toggleAbs(inp.value, inp.selectionStart);
        if (r) { inp.value = r.text; inp.setSelectionRange(r.caret, r.caret); this.syncInputs(inp); this.point = null; this.onEditInput(); }
        return;
      }
      if (this.edit.enterMode && !inp.value.startsWith('=') && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
        e.preventDefault();
        this.commit({ ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' }[e.key]);
      }
    }

    isPointable() {
      const inp = this.activeInput();
      if (!inp || !inp.value.startsWith('=')) return false;
      const caret = inp.selectionStart;
      if (this.point && caret === this.point.end) return true;
      const before = inp.value.slice(0, caret).replace(/\s+$/, '');
      return /[=(,;+\-*/^&<>:]$/.test(before);
    }

    pointAt(r, c) {
      const inp = this.activeInput();
      const ref = XL.addr(r, c);
      let start = inp.selectionStart, end = inp.selectionEnd;
      if (this.point && start === this.point.end) { start = this.point.start; end = this.point.end; }
      inp.value = inp.value.slice(0, start) + ref + inp.value.slice(end);
      this.point = { start, end: start + ref.length, ar: r, ac: c };
      inp.setSelectionRange(this.point.end, this.point.end);
      this.syncInputs(inp);
      inp.focus();
      this.onEditInput();
    }

    pointExtend(r, c) {
      const inp = this.activeInput();
      const p = this.point;
      const ref = p.ar === r && p.ac === c ? XL.addr(r, c) : XL.addr(Math.min(p.ar, r), Math.min(p.ac, c)) + ':' + XL.addr(Math.max(p.ar, r), Math.max(p.ac, c));
      inp.value = inp.value.slice(0, p.start) + ref + inp.value.slice(p.end);
      p.end = p.start + ref.length;
      inp.setSelectionRange(p.end, p.end);
      this.syncInputs(inp);
      this.onEditInput();
    }

    endEdit() {
      if (!this.edit) return;
      const inp = this.edit.input;
      this.edit = null;
      this.point = null;
      inp.remove();
      this.refLayer.innerHTML = '';
      this.hintBox.style.display = 'none';
      this.acBox.style.display = 'none';
      this.wrap.focus({ preventScroll: true });
    }

    cancel() { this.endEdit(); this.updateSelUI(); }

    commit(dir) {
      if (!this.edit) return;
      const { r, c } = this.edit;
      let text = this.activeInput().value;
      if (text.startsWith('=')) text = this.autoClose(text);
      this.endEdit();
      this.pushUndo();
      if (dir === 'fillsel') {
        const n = this.norm();
        for (let rr = n.r1; rr <= n.r2; rr++) for (let cc = n.c1; cc <= n.c2; cc++) this.sheet.setInput(rr, cc, XL.shiftFormula(text, rr - r, cc - c));
      } else this.sheet.setInput(r, c, text);
      this.changed();
      const mv = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[dir];
      if (mv) { const nr = r + mv[0], nc = c + mv[1]; this.select(nr, nc, nr, nc, nr, nc); }
      else this.updateSelUI();
    }

    // Tự đóng ngoặc còn thiếu như Excel.
    autoClose(text) {
      let depth = 0, inStr = false;
      for (const ch of text) { if (ch === '"') inStr = !inStr; else if (!inStr) { if (ch === '(') depth++; else if (ch === ')') depth--; } }
      return depth > 0 && !inStr ? text + ')'.repeat(depth) : text;
    }

    changed() {
      this.refresh();
      if (this.opts.onChange) this.opts.onChange();
    }

    // ---------- Hoàn tác ----------
    snapshot() {
      const out = [];
      for (const [key, x] of this.sheet.cells) if (!x.locked) out.push([key, { formula: x.formula, lit: x.lit, fmt: x.fmt }]);
      return out;
    }
    restore(snap) {
      const m = new Map(snap);
      for (const [key, x] of this.sheet.cells) {
        if (x.locked) continue;
        const d = m.get(key) || { lit: null };
        delete x.tree; delete x.parseError;
        if (d.formula) x.formula = d.formula; else delete x.formula;
        x.lit = d.lit === undefined ? null : d.lit;
        if (d.fmt) x.fmt = d.fmt; else delete x.fmt;
      }
      this.sheet.dirty = true;
    }
    pushUndo() { this.undoStack.push(this.snapshot()); if (this.undoStack.length > 100) this.undoStack.shift(); this.redoStack = []; }
    undo() { if (this.edit) return; const s = this.undoStack.pop(); if (!s) return; this.redoStack.push(this.snapshot()); this.restore(s); this.changed(); }
    redoAct() { if (this.edit) return; const s = this.redoStack.pop(); if (!s) return; this.undoStack.push(this.snapshot()); this.restore(s); this.changed(); }

    // ---------- Thao tác vùng ----------
    eachSel(fn) { const n = this.norm(); for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) fn(r, c); }
    writable(r, c) { const x = this.sheet.cell(r, c); return !(x && x.locked); }

    clearSel() {
      if (this.edit) return;
      this.pushUndo();
      let locked = 0;
      this.eachSel((r, c) => { if (this.writable(r, c)) this.sheet.setInput(r, c, ''); else locked++; });
      if (locked && this.norm().r1 === this.norm().r2 && this.norm().c1 === this.norm().c2) this.lockedMsg();
      this.changed();
    }

    setFormat(fmt) {
      if (this.edit) this.commit(null);
      this.pushUndo();
      this.eachSel((r, c) => { if (this.writable(r, c)) { const x = this.sheet.ensure(r, c); x.fmt = fmt; } });
      this.sheet.dirty = true;
      this.changed();
      this.wrap.focus({ preventScroll: true });
    }

    copyCell(sr, sc, dr, dc) {
      if (!this.writable(dr, dc)) return;
      const x = this.sheet.cell(sr, sc);
      const text = this.sheet.input(sr, sc);
      this.sheet.setInput(dr, dc, XL.shiftFormula(text, dr - sr, dc - sc));
      if (x && x.fmt && !(x.locked)) this.sheet.ensure(dr, dc).fmt = x.fmt;
    }

    fillDown() {
      if (this.edit) this.commit(null);
      const n = this.norm();
      const src = n.r1 === n.r2 ? n.r1 - 1 : n.r1;
      if (src < 0) return;
      this.pushUndo();
      for (let c = n.c1; c <= n.c2; c++) for (let r = src + 1; r <= n.r2; r++) this.copyCell(src, c, r, c);
      this.changed();
    }

    fillRight() {
      if (this.edit) this.commit(null);
      const n = this.norm();
      const src = n.c1 === n.c2 ? n.c1 - 1 : n.c1;
      if (src < 0) return;
      this.pushUndo();
      for (let r = n.r1; r <= n.r2; r++) for (let c = src + 1; c <= n.c2; c++) this.copyCell(r, src, r, c);
      this.changed();
    }

    fillPreview(r, c) {
      const n = this.norm();
      let t;
      const dDown = r > n.r2 ? r - n.r2 : r < n.r1 ? n.r1 - r : 0;
      const dRight = c > n.c2 ? c - n.c2 : c < n.c1 ? n.c1 - c : 0;
      if (dDown === 0 && dRight === 0) { this.fillTarget = null; this.fillBox.style.display = 'none'; return; }
      if (dDown >= dRight) t = { r1: Math.min(n.r1, r), r2: Math.max(n.r2, r), c1: n.c1, c2: n.c2, dir: r > n.r2 ? 'down' : 'up' };
      else t = { r1: n.r1, r2: n.r2, c1: Math.min(n.c1, c), c2: Math.max(n.c2, c), dir: c > n.c2 ? 'right' : 'left' };
      this.fillTarget = t;
      this.place(this.fillBox, this.rectOf(t.r1, t.c1, t.r2, t.c2));
    }

    // Kéo nút điền: sao chép công thức (dịch tham chiếu) hoặc tạo dãy số nếu nguồn là dãy cách đều.
    fillTo(t) {
      const n = this.norm();
      this.pushUndo();
      const vertical = t.dir === 'down' || t.dir === 'up';
      const h = n.r2 - n.r1 + 1, w = n.c2 - n.c1 + 1;
      const series = (vals) => {
        if (vals.length < 2 || !vals.every((v) => typeof v === 'number')) return null;
        const step = vals[1] - vals[0];
        for (let i = 2; i < vals.length; i++) if (Math.abs(vals[i] - vals[i - 1] - step) > 1e-9) return null;
        return step;
      };
      if (vertical) {
        for (let c = n.c1; c <= n.c2; c++) {
          const srcVals = [];
          for (let r = n.r1; r <= n.r2; r++) { const x = this.sheet.cell(r, c); srcVals.push(x && !x.formula ? x.lit : undefined); }
          const step = series(srcVals);
          const rows = t.dir === 'down' ? range(n.r2 + 1, t.r2) : range(t.r1, n.r1 - 1).reverse();
          rows.forEach((r, i) => {
            if (!this.writable(r, c)) return;
            if (step !== null) {
              const v = t.dir === 'down' ? srcVals[h - 1] + step * (i + 1) : srcVals[0] - step * (i + 1);
              this.sheet.setInput(r, c, String(v));
            } else {
              const off = t.dir === 'down' ? i % h : h - 1 - (i % h);
              this.copyCell(n.r1 + off, c, r, c);
            }
          });
        }
      } else {
        for (let r = n.r1; r <= n.r2; r++) {
          const srcVals = [];
          for (let c = n.c1; c <= n.c2; c++) { const x = this.sheet.cell(r, c); srcVals.push(x && !x.formula ? x.lit : undefined); }
          const step = series(srcVals);
          const cols = t.dir === 'right' ? range(n.c2 + 1, t.c2) : range(t.c1, n.c1 - 1).reverse();
          cols.forEach((c, i) => {
            if (!this.writable(r, c)) return;
            if (step !== null) {
              const v = t.dir === 'right' ? srcVals[w - 1] + step * (i + 1) : srcVals[0] - step * (i + 1);
              this.sheet.setInput(r, c, String(v));
            } else {
              const off = t.dir === 'right' ? i % w : w - 1 - (i % w);
              this.copyCell(r, n.c1 + off, r, c);
            }
          });
        }
      }
      this.sel = { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2, ar: this.sel.ar, ac: this.sel.ac };
      this.changed();
    }

    autoSum() {
      if (this.edit) return;
      const { ar, ac } = this.sel;
      if (this.lockedMsg()) return;
      const isNum = (r, c) => typeof this.sheet.value(r, c) === 'number' && !((this.sheet.cell(r, c) || {}).cls || '').includes('x-head');
      let r = ar - 1;
      while (r >= 0 && isNum(r, ac)) r--;
      let text;
      if (r < ar - 1) text = '=SUM(' + XL.addr(r + 1, ac) + ':' + XL.addr(ar - 1, ac) + ')';
      else {
        let c = ac - 1;
        while (c >= 0 && isNum(ar, c)) c--;
        text = c < ac - 1 ? '=SUM(' + XL.addr(ar, c + 1) + ':' + XL.addr(ar, ac - 1) + ')' : '=SUM()';
      }
      this.startEdit(text, false);
      if (text === '=SUM()') this.edit.input.setSelectionRange(5, 5);
    }

    copy(e) {
      const n = this.norm();
      const rows = [], inputs = [];
      for (let r = n.r1; r <= n.r2; r++) {
        const a = [], b = [];
        for (let c = n.c1; c <= n.c2; c++) { a.push(this.sheet.display(r, c)); b.push({ text: this.sheet.input(r, c), fmt: (this.sheet.cell(r, c) || {}).fmt }); }
        rows.push(a.join('\t')); inputs.push(b);
      }
      const text = rows.join('\n');
      this.clip = { r1: n.r1, c1: n.c1, inputs, text };
      e.clipboardData.setData('text/plain', text);
      e.preventDefault();
      this.toast('Đã sao chép ' + this.nameBox.textContent + '. Chọn ô đích rồi Ctrl+V.');
    }

    paste(e) {
      const text = (e.clipboardData.getData('text/plain') || '').replace(/\r/g, '');
      e.preventDefault();
      const n = this.norm();
      this.pushUndo();
      if (this.clip && text.replace(/\n$/, '') === this.clip.text) {
        const ins = this.clip.inputs, h = ins.length, w = ins[0].length;
        const H = Math.max(h, (n.r2 - n.r1 + 1) % h === 0 ? n.r2 - n.r1 + 1 : h);
        const W = Math.max(w, (n.c2 - n.c1 + 1) % w === 0 ? n.c2 - n.c1 + 1 : w);
        for (let i = 0; i < H; i++) for (let j = 0; j < W; j++) {
          const r = n.r1 + i, c = n.c1 + j;
          if (r >= this.sheet.nRows || c >= this.sheet.nCols || !this.writable(r, c)) continue;
          const src = ins[i % h][j % w];
          const sr = this.clip.r1 + (i % h), sc = this.clip.c1 + (j % w);
          this.sheet.setInput(r, c, XL.shiftFormula(src.text, r - sr, c - sc));
          if (src.fmt) this.sheet.ensure(r, c).fmt = src.fmt;
        }
      } else {
        text.replace(/\n$/, '').split('\n').forEach((line, i) => line.split('\t').forEach((v, j) => {
          const r = n.r1 + i, c = n.c1 + j;
          if (r < this.sheet.nRows && c < this.sheet.nCols && this.writable(r, c)) this.sheet.setInput(r, c, v);
        }));
      }
      this.changed();
    }
  }

  function range(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }

  global.SheetView = SheetView;
  global.FUNC_HELP = FUNC_HELP;
})(window);
