/* Trình soạn thảo kiểu Word + bộ đọc định dạng để chấm bài.
   Văn bản trong mỗi đoạn được lưu thành các "run" phẳng: <span class="r" data-s='{"ff":"Tahoma","fs":26,"b":1}'>. */
(function (global) {
  'use strict';

  const FONTS = ['Times New Roman', 'Arial', 'Tahoma', 'Calibri', 'Verdana', 'Cambria', 'Georgia', 'Courier New', 'Segoe UI'];
  const SIZES = [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 48, 72];
  const COLORS = ['#000000', '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0', '#ffffff', '#7f7f7f', '#f2dcdb', '#dbe5f1', '#ebf1de', '#fde9d9', '#e5e0ec'];
  const SHADES = ['#ffffff', '#f2f2f2', '#d9d9d9', '#fff2cc', '#fde9d9', '#e2efda', '#ddebf7', '#dbe5f1', '#e5e0ec', '#fbe4d5', '#c6e0b4', '#bdd7ee', '#ffd966', '#f4b084', '#9bc2e6', '#a9d08e'];
  const BULLETS = [['disc', '•'], ['hand', '☞'], ['diamond', '❖'], ['check', '✓'], ['arrow', '➢'], ['star', '★'], ['plane', '✈']];
  const SYMBOLS = ['☞', '✈', '★', '❖', '✓', '✔', '☎', '♥', '⇒', '➢', '©', '®', '™', '€', '♣', '♠', '◆', '●', '☺', '✿'];
  const BLOCK_SEL = 'p, li';

  const el = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt !== undefined) e.textContent = txt; return e; };
  const isAtom = (n) => n.nodeType === 1 && (n.tagName === 'IMG' || n.classList.contains('dropcap'));

  // ---------- Đọc định dạng ký tự ----------
  function firstFamily(s) { return (s || '').split(',')[0].replace(/["']/g, '').trim(); }
  function toPt(v) {
    if (!v) return undefined;
    const n = parseFloat(v);
    if (/pt$/.test(v)) return n;
    if (/px$/.test(v)) return Math.round(n * 0.75 * 2) / 2;
    return undefined;
  }
  function readInline(e, st) {
    const set = (key, val) => { if (st[key] === undefined && val !== undefined) st[key] = val; };
    if (e.dataset && e.dataset.s) {
      let d = {};
      try { d = JSON.parse(e.dataset.s); } catch (err) { d = {}; }
      for (const key in d) set(key, d[key]);
      ['ff', 'fs', 'b', 'i', 'ul', 'st', 'va', 'caps', 'color', 'hl', 'sh', 'bx', 'bxbg'].forEach((key) => set(key, 0));
      return;
    }
    const s = e.style, tag = e.tagName;
    if (s.fontFamily) set('ff', firstFamily(s.fontFamily));
    if (s.fontSize) set('fs', toPt(s.fontSize));
    if (s.fontWeight) set('b', s.fontWeight === 'bold' || parseInt(s.fontWeight, 10) >= 600 ? 1 : 0);
    if (tag === 'B' || tag === 'STRONG') set('b', 1);
    if (s.fontStyle) set('i', s.fontStyle === 'italic' ? 1 : 0);
    if (tag === 'I' || tag === 'EM') set('i', 1);
    const deco = s.textDecorationLine || s.textDecoration || '';
    if (deco.includes('underline')) set('ul', s.textDecorationStyle === 'double' ? 'double' : 'single');
    if (deco.includes('line-through')) set('st', s.textDecorationStyle === 'double' ? 'double' : 'single');
    if (tag === 'U') set('ul', 'single');
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') set('st', 'single');
    if (tag === 'SUP' || s.verticalAlign === 'super') set('va', 'super');
    if (tag === 'SUB' || s.verticalAlign === 'sub') set('va', 'sub');
    if ((s.fontVariant || '').includes('small-caps') || s.fontVariantCaps === 'small-caps') set('caps', 'small');
    if (s.textTransform === 'uppercase') set('caps', 'all');
    if (s.color) set('color', s.color);
    if (s.backgroundColor) set('hl', s.backgroundColor);
    if (s.textShadow && s.textShadow !== 'none') set('sh', 1);
    if (tag === 'FONT') { if (e.getAttribute('face')) set('ff', firstFamily(e.getAttribute('face'))); if (e.getAttribute('color')) set('color', e.getAttribute('color')); }
  }
  function styleOf(node, block) {
    const st = {};
    let e = node.nodeType === 1 ? node : node.parentElement;
    while (e && e !== block && e.parentElement) {
      if (e.classList && e.classList.contains('st')) {
        const d = e.dataset.st; if (st.st === undefined) st.st = d || 'single';
      } else readInline(e, st);
      e = e.parentElement;
    }
    const clean = {};
    for (const key in st) if (st[key] !== 0 && st[key] !== undefined && st[key] !== null && st[key] !== '') clean[key] = st[key];
    return clean;
  }
  const styleKey = (st) => JSON.stringify(Object.keys(st).sort().reduce((o, key) => { o[key] = st[key]; return o; }, {}));

  // ---------- Run: dàn phẳng / dựng lại đoạn ----------
  function flatten(block) {
    const items = [];
    const walk = (n) => {
      for (const ch of Array.from(n.childNodes)) {
        if (ch.nodeType === 3) { if (ch.nodeValue) items.push({ t: 'text', text: ch.nodeValue, st: styleOf(ch, block) }); }
        else if (ch.nodeType === 1) {
          if (ch.tagName === 'BR') items.push({ t: 'br' });
          else if (isAtom(ch)) items.push({ t: 'atom', node: ch });
          else walk(ch);
        }
      }
    };
    walk(block);
    const out = [];
    for (const it of items) {
      const last = out[out.length - 1];
      if (it.t === 'text' && last && last.t === 'text' && styleKey(last.st) === styleKey(it.st)) last.text += it.text;
      else out.push(it);
    }
    return out;
  }

  function runCss(st, piece) {
    const css = [];
    if (st.ff) css.push(`font-family:"${st.ff}"`);
    if (st.fs) css.push(`font-size:${st.va ? Math.round(st.fs * 0.65 * 2) / 2 : st.fs}pt`);
    else if (st.va) css.push('font-size:65%');
    if (st.b) css.push('font-weight:700');
    if (st.i) css.push('font-style:italic');
    if (st.ul && !(st.ul === 'words' && piece === 'space')) {
      css.push('text-decoration-line:underline');
      if (st.ul === 'double') css.push('text-decoration-style:double');
      if (st.ul === 'dotted') css.push('text-decoration-style:dotted');
    }
    if (st.va) css.push('vertical-align:' + st.va);
    if (st.caps === 'small') css.push('font-variant:small-caps');
    if (st.caps === 'all') css.push('text-transform:uppercase');
    if (st.color) css.push('color:' + st.color);
    if (st.hl) css.push('background-color:' + st.hl);
    else if (st.bxbg) css.push('background-color:' + st.bxbg);
    if (st.sh) css.push('text-shadow:1px 1px 2px rgba(0,0,0,.45)');
    if (st.bx) css.push('border:1px solid currentColor;padding:0 3px;box-decoration-break:clone;-webkit-box-decoration-break:clone');
    return css.join(';');
  }

  function renderRun(text, st) {
    const json = JSON.stringify(st);
    const pieces = st.ul === 'words' ? text.split(/(\s+)/).filter((x) => x) : [text];
    return pieces.map((p) => {
      const span = el('span', 'r');
      span.dataset.s = json;
      span.style.cssText = runCss(st, /^\s+$/.test(p) ? 'space' : 'word');
      if (st.st) {
        const inner = el('span', 'st', p);
        inner.dataset.st = st.st;
        span.appendChild(inner);
      } else span.textContent = p;
      return span;
    });
  }

  function rebuild(block, items) {
    while (block.firstChild) block.removeChild(block.firstChild);
    for (const it of items) {
      if (it.t === 'text') { if (it.text) renderRun(it.text, it.st).forEach((n) => block.appendChild(n)); }
      else if (it.t === 'br') block.appendChild(el('br'));
      else block.appendChild(it.node);
    }
    if (!block.firstChild) block.appendChild(el('br'));
  }

  const itemLen = (it) => (it.t === 'text' ? it.text.length : 1);

  // Vị trí (node, offset) → số ký tự tính từ đầu đoạn.
  function offsetIn(block, container, off) {
    const r = document.createRange();
    r.setStart(block, 0);
    try { r.setEnd(container, off); } catch (e) { return 0; }
    const frag = r.cloneContents();
    let n = 0;
    const walk = (x) => {
      for (const ch of Array.from(x.childNodes)) {
        if (ch.nodeType === 3) n += ch.nodeValue.length;
        else if (ch.nodeType === 1) { if (ch.tagName === 'BR' || isAtom(ch)) n += 1; else walk(ch); }
      }
    };
    walk(frag);
    return n;
  }
  function boundaryAt(block, n) {
    let res = null;
    const walk = (x) => {
      for (const ch of Array.from(x.childNodes)) {
        if (res) return;
        if (ch.nodeType === 3) {
          const L = ch.nodeValue.length;
          if (n <= L) { res = [ch, n]; return; }
          n -= L;
        } else if (ch.nodeType === 1) {
          if (ch.tagName === 'BR' || isAtom(ch)) {
            if (n === 0) { res = [ch.parentNode, Array.prototype.indexOf.call(ch.parentNode.childNodes, ch)]; return; }
            n -= 1;
          } else walk(ch);
        }
      }
    };
    walk(block);
    return res || [block, block.childNodes.length];
  }

  // Lọc HTML dán vào: chỉ giữ đoạn, xuống dòng, danh sách và run định dạng của trình soạn thảo.
  function cleanHTML(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const out = document.createElement('div');
    const copy = (src, dst) => {
      for (const n of Array.from(src.childNodes)) {
        if (n.nodeType === 3) { dst.appendChild(document.createTextNode(n.nodeValue)); continue; }
        if (n.nodeType !== 1) continue;
        const tag = n.tagName;
        let e = null;
        if (tag === 'BR') { dst.appendChild(el('br')); continue; }
        if (['P', 'UL', 'OL', 'LI'].includes(tag)) { e = el(tag.toLowerCase()); if (n.className && /^bl-/.test(n.className)) e.className = n.className; if (n.getAttribute('style')) e.setAttribute('style', n.getAttribute('style')); }
        else if (tag === 'SPAN' && n.dataset.s) { e = el('span', 'r'); e.dataset.s = n.dataset.s; let st = {}; try { st = JSON.parse(n.dataset.s); } catch (err) { st = {}; } e.style.cssText = runCss(st, 'word'); }
        else if (tag === 'SPAN' && n.classList.contains('st')) { e = el('span', 'st'); e.dataset.st = n.dataset.st || 'single'; }
        if (e) { copy(n, e); dst.appendChild(e); } else copy(n, dst);
      }
    };
    copy(doc.body, out);
    return out.innerHTML;
  }

  // ---------- Bộ đọc tài liệu để chấm ----------
  class WordDoc {
    constructor(page, defaults) {
      this.page = page;
      const cs = getComputedStyle(page);
      this.def = { ff: firstFamily(cs.fontFamily).toLowerCase(), fs: toPt(cs.fontSize) || 12 };
      Object.assign(this.def, defaults || {});
      this.chars = [];
      this.blocks = Array.from(page.querySelectorAll(BLOCK_SEL)).filter((b) => !b.querySelector(BLOCK_SEL));
      for (const b of this.blocks) {
        const walk = (n) => {
          for (const ch of Array.from(n.childNodes)) {
            if (ch.nodeType === 3) {
              const st = this.full(styleOf(ch, b));
              for (const c of ch.nodeValue) this.chars.push({ ch: c, st, block: b });
            } else if (ch.nodeType === 1) {
              if (ch.tagName === 'BR') this.chars.push({ ch: '\n', st: {}, block: b });
              else if (ch.classList.contains('dropcap')) this.chars.push({ ch: ch.textContent, st: { dropcap: 1 }, block: b, dropcap: ch });
              else if (ch.tagName === 'IMG') { /* bỏ qua */ }
              else walk(ch);
            }
          }
        };
        walk(b);
        this.chars.push({ ch: '\n', st: {}, block: b });
      }
      // Chuỗi chuẩn hóa: gộp khoảng trắng, không phân biệt hoa thường.
      this.norm = '';
      this.map = [];
      let prevSpace = true;
      this.chars.forEach((c, i) => {
        const sp = /\s/.test(c.ch);
        if (sp) { if (!prevSpace) { this.norm += ' '; this.map.push(i); } prevSpace = true; }
        else { this.norm += c.ch.normalize('NFC').toLowerCase(); this.map.push(i); prevSpace = false; }
      });
    }
    full(st) {
      return {
        ff: (st.ff || this.def.ff).toLowerCase(), fs: st.fs || this.def.fs, b: !!st.b, i: !!st.i, ul: st.ul || '', st: st.st || '',
        va: st.va || '', caps: st.caps || '', color: st.color || '', hl: st.hl || st.bxbg || '', sh: !!st.sh, bx: !!st.bx,
      };
    }
    static n(s) { return s.normalize('NFC').toLowerCase().replace(/\s+/g, ' ').trim(); }
    find(needle, nth) {
      const q = WordDoc.n(needle);
      let pos = -1, from = 0;
      for (let k = 0; k <= (nth || 0); k++) { pos = this.norm.indexOf(q, from); if (pos < 0) return null; from = pos + 1; }
      return { start: this.map[pos], end: this.map[pos + q.length - 1] + 1 };
    }
    count(needle) { const q = WordDoc.n(needle); let n = 0, p = this.norm.indexOf(q); while (p >= 0) { n++; p = this.norm.indexOf(q, p + 1); } return n; }
    indexOf(needle, nth) { const f = this.find(needle, nth); return f ? f.start : -1; }
    charsOf(needle, nth) {
      const f = this.find(needle, nth);
      if (!f) return null;
      return this.chars.slice(f.start, f.end).filter((c) => !/\s/.test(c.ch) && !c.dropcap);
    }
    // Mọi ký tự (không tính khoảng trắng) của đoạn văn bản thỏa điều kiện.
    all(needle, pred, nth) { const cs = this.charsOf(needle, nth); return !!(cs && cs.length && cs.every((c) => pred(c.st))); }
    blockOf(needle, nth) { const f = this.find(needle, nth); return f ? this.chars[f.start].block : null; }
    has(needle) { return !!this.find(needle); }
    dropcapOf(needle) { const b = this.blockOf(needle); return b ? b.querySelector('.dropcap') : null; }
    colsOf(needle) { const b = this.blockOf(needle); const c = b && b.closest('.w-cols'); return c ? { n: +c.dataset.cols, line: c.dataset.line === '1' } : { n: 1, line: false }; }
    tables() { return Array.from(this.page.querySelectorAll('table.w-table')); }
    blocksOf(needle) { const out = []; for (let i = 0; ; i++) { const b = this.blockOf(needle, i); if (!b) return out; out.push(b); } }
    tableWith(needle) { for (const b of this.blocksOf(needle)) { const t = b.closest('table.w-table'); if (t) return t; } return null; }
    cellWith(needle) { for (const b of this.blocksOf(needle)) { const td = b.closest('table.w-table td'); if (td) return td; } return null; }
    images() { return Array.from(this.page.querySelectorAll('img.w-img')); }
    blockStyle(needle) { const b = this.blockOf(needle); return b ? b.style : null; }
  }

  // ---------- Trình soạn thảo ----------
  class WordEditor {
    constructor(host, opts) {
      this.host = host;
      this.opts = opts || {};
      this.undoStack = [];
      this.redoStack = [];
      this.lastSnap = 0;
      this.build();
      this.setHTML(this.opts.html || '<p><br></p>');
    }

    build() {
      const root = el('div', 'wd');
      this.root = root;
      const tabs = el('div', 'wd-tabs');
      const ribbon = el('div', 'wd-ribbon');
      this.tabsEl = tabs;
      this.panels = {};
      const tabNames = [['home', 'Home'], ['insert', 'Insert'], ['layout', 'Layout'], ['table', 'Table Tools'], ['picture', 'Picture Tools']];
      tabNames.forEach(([id, label]) => {
        const t = el('button', 'wd-tab', label);
        t.type = 'button';
        t.dataset.tab = id;
        t.addEventListener('mousedown', (e) => e.preventDefault());
        t.addEventListener('click', () => this.showTab(id));
        if (id === 'table' || id === 'picture') t.classList.add('ctx');
        tabs.appendChild(t);
        const p = el('div', 'wd-panel');
        p.dataset.tab = id;
        ribbon.appendChild(p);
        this.panels[id] = p;
      });
      const undo = el('div', 'wd-qat');
      undo.appendChild(this.mkBtn('↶', 'Hoàn tác (Ctrl+Z)', () => this.undo()));
      undo.appendChild(this.mkBtn('↷', 'Làm lại (Ctrl+Y)', () => this.redo()));
      tabs.prepend(undo);
      root.appendChild(tabs);
      root.appendChild(ribbon);
      this.buildHome();
      this.buildInsert();
      this.buildLayout();
      this.buildTable();
      this.buildPicture();

      const area = el('div', 'wd-area');
      const page = el('div', 'wd-page');
      page.contentEditable = 'true';
      page.spellcheck = false;
      this.page = page;
      area.appendChild(page);
      root.appendChild(area);
      this.toastBox = el('div', 'xl-toast');
      root.appendChild(this.toastBox);
      this.host.innerHTML = '';
      this.host.appendChild(root);
      this.showTab('home');
      this.bind();
    }

    mkBtn(label, title, fn, cls) {
      const b = el('button', 'wd-btn ' + (cls || ''));
      b.type = 'button';
      b.innerHTML = label;
      b.title = title;
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', (e) => { e.stopPropagation(); fn(e, b); });
      return b;
    }
    group(panel, name) {
      const g = el('div', 'wd-group');
      const body = el('div', 'wd-gbody');
      g.appendChild(body);
      g.appendChild(el('div', 'wd-gname', name));
      panel.appendChild(g);
      return body;
    }
    // Nút có menu thả xuống.
    menu(label, title, build) {
      const wrap = el('div', 'wd-dd');
      const b = this.mkBtn(label + ' <span class="car">▾</span>', title, () => {
        const open = wrap.classList.contains('open');
        this.closeMenus();
        if (!open) wrap.classList.add('open');
      });
      wrap.appendChild(b);
      const m = el('div', 'wd-menu');
      m.addEventListener('mousedown', (e) => { if (!e.target.closest('input, label')) e.preventDefault(); });
      build(m, () => this.closeMenus());
      wrap.appendChild(m);
      return wrap;
    }
    closeMenus() { this.root.querySelectorAll('.wd-dd.open').forEach((d) => d.classList.remove('open')); }
    mItem(m, html, fn, close) { const it = el('div', 'wd-mi'); it.innerHTML = html; it.addEventListener('click', () => { fn(); close(); }); m.appendChild(it); return it; }
    swatches(m, colors, fn, close, noneLabel) {
      if (noneLabel) this.mItem(m, noneLabel, () => fn(null), close);
      const g = el('div', 'wd-sw');
      colors.forEach((c) => { const s = el('button', 'sw'); s.type = 'button'; s.style.background = c; s.title = c; s.addEventListener('click', () => { fn(c); close(); }); g.appendChild(s); });
      m.appendChild(g);
    }

    buildHome() {
      const p = this.panels.home;
      const f = this.group(p, 'Font');
      const row1 = el('div', 'wd-row'), row2 = el('div', 'wd-row');
      f.appendChild(row1); f.appendChild(row2);
      const fsel = el('select', 'wd-font');
      FONTS.forEach((n) => { const o = el('option', '', n); o.value = n; o.style.fontFamily = n; fsel.appendChild(o); });
      fsel.title = 'Font chữ';
      fsel.addEventListener('mousedown', () => this.saveSel());
      fsel.addEventListener('change', () => { this.restoreSel(); this.applyInline(() => ({ ff: fsel.value })); });
      this.fontSel = fsel;
      const ssel = el('select', 'wd-size');
      SIZES.forEach((n) => { const o = el('option', '', String(n)); o.value = n; ssel.appendChild(o); });
      ssel.title = 'Cỡ chữ';
      ssel.addEventListener('mousedown', () => this.saveSel());
      ssel.addEventListener('change', () => { this.restoreSel(); this.applyInline(() => ({ fs: +ssel.value })); });
      this.sizeSel = ssel;
      row1.appendChild(fsel); row1.appendChild(ssel);
      row1.appendChild(this.mkBtn('A<sup>+</sup>', 'Tăng cỡ chữ (Ctrl+])', () => this.bumpSize(1)));
      row1.appendChild(this.mkBtn('A<sup>−</sup>', 'Giảm cỡ chữ (Ctrl+[)', () => this.bumpSize(-1)));
      row1.appendChild(this.mkBtn('<span style="text-decoration:line-through">A</span>✕', 'Xóa định dạng', () => this.clearFormat()));
      this.bBtn = this.mkBtn('<b>B</b>', 'In đậm (Ctrl+B)', () => this.toggle('b'));
      this.iBtn = this.mkBtn('<i>I</i>', 'In nghiêng (Ctrl+I)', () => this.toggle('i'));
      row2.appendChild(this.bBtn); row2.appendChild(this.iBtn);
      this.uBtn = this.mkBtn('<u>U</u>', 'Gạch dưới (Ctrl+U)', () => this.toggleVal('ul', 'single'));
      row2.appendChild(this.uBtn);
      row2.appendChild(this.menu('', 'Kiểu gạch dưới', (m, close) => {
        this.mItem(m, '<u>Gạch đơn</u> (Single)', () => this.setVal('ul', 'single'), close);
        this.mItem(m, '<u style="text-decoration-style:double">Gạch đôi</u> (Double)', () => this.setVal('ul', 'double'), close);
        this.mItem(m, '<u>Gạch</u> <u>từng</u> <u>từ</u> (Words only)', () => this.setVal('ul', 'words'), close);
        this.mItem(m, '<u style="text-decoration-style:dotted">Gạch chấm</u> (Dotted)', () => this.setVal('ul', 'dotted'), close);
        this.mItem(m, 'Bỏ gạch dưới', () => this.setVal('ul', null), close);
      }));
      row2.appendChild(this.mkBtn('<s>abc</s>', 'Gạch ngang (Strikethrough)', () => this.toggleVal('st', 'single')));
      row2.appendChild(this.mkBtn('<s style="text-decoration-style:double">abc</s>', 'Gạch ngang đôi (Double strikethrough)', () => this.toggleVal('st', 'double')));
      row2.appendChild(this.mkBtn('x<sub>2</sub>', 'Chỉ số dưới (Ctrl+=)', () => this.toggleVal('va', 'sub')));
      row2.appendChild(this.mkBtn('x<sup>2</sup>', 'Chỉ số trên (Ctrl+Shift+=)', () => this.toggleVal('va', 'super')));
      row2.appendChild(this.menu('Aa', 'Kiểu chữ hoa', (m, close) => {
        this.mItem(m, '<span style="font-variant:small-caps">Small Caps</span> – chữ hoa nhỏ (Ctrl+Shift+K)', () => this.toggleVal('caps', 'small'), close);
        this.mItem(m, 'ALL CAPS – chữ hoa lớn (Ctrl+Shift+A)', () => this.toggleVal('caps', 'all'), close);
        this.mItem(m, 'Bỏ', () => this.setVal('caps', null), close);
      }));
      row2.appendChild(this.menu('<span style="border-bottom:3px solid #c00000">A</span>', 'Màu chữ', (m, close) => this.swatches(m, COLORS, (c) => this.setVal('color', c), close, 'Tự động (Automatic)')));
      row2.appendChild(this.menu('<span style="background:#ffff00;padding:0 2px;color:#000">ab</span>', 'Tô sáng (Highlight)', (m, close) => this.swatches(m, ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#ff0000', '#d9d9d9'], (c) => this.setVal('hl', c), close, 'Không tô')));
      row2.appendChild(this.mkBtn('<span style="text-shadow:1px 1px 2px #555">A</span>', 'Hiệu ứng bóng chữ (Text Effects → Shadow)', () => this.toggleVal('sh', 1)));

      const para = this.group(p, 'Paragraph');
      const r1 = el('div', 'wd-row'), r2 = el('div', 'wd-row');
      para.appendChild(r1); para.appendChild(r2);
      r1.appendChild(this.menu('☰•', 'Bullets – đánh dấu đầu dòng', (m, close) => {
        BULLETS.forEach(([k, s]) => this.mItem(m, `<span class="bl-prev">${s}</span> Ký hiệu ${s}`, () => this.bullets(k), close));
        this.mItem(m, 'Bỏ bullets', () => this.bullets(null), close);
      }));
      r1.appendChild(this.mkBtn('1.☰', 'Numbering – đánh số', () => this.numbering()));
      r1.appendChild(this.mkBtn('⇥¶', 'Thụt lề dòng đầu 1 cm (First line indent)', () => this.toggleIndent()));
      r1.appendChild(this.menu('↕', 'Giãn dòng (Line spacing)', (m, close) => {
        [['1.0', 1], ['1.15', 1.15], ['1.5', 1.5], ['2.0', 2]].forEach(([t, v]) => this.mItem(m, t, () => this.setBlock({ lineHeight: String(v) }), close));
        this.mItem(m, 'Spacing After 6pt', () => this.setBlock({ marginBottom: '6pt' }), close);
        this.mItem(m, 'Spacing After 0pt', () => this.setBlock({ marginBottom: '0pt' }), close);
      }));
      r1.appendChild(this.menu('▢', 'Borders and Shading', (m, close) => {
        this.mItem(m, '▭ Viền chữ (Text border)', () => this.toggleVal('bx', 1), close);
        this.mItem(m, '▤ Viền đoạn (Paragraph border)', () => this.toggleBlockBorder(), close);
        m.appendChild(el('div', 'wd-mh', 'Tô nền chữ (Text shading)'));
        this.swatches(m, SHADES, (c) => this.setVal('bxbg', c), close);
        m.appendChild(el('div', 'wd-mh', 'Tô nền đoạn (Paragraph shading)'));
        this.swatches(m, SHADES, (c) => this.setBlock({ backgroundColor: c || '' }), close, 'Không tô nền đoạn');
      }));
      this.alignBtns = {};
      [['left', 'Căn trái (Ctrl+L)'], ['center', 'Căn giữa (Ctrl+E)'], ['right', 'Căn phải (Ctrl+R)'], ['justify', 'Căn đều hai bên (Ctrl+J)']].forEach(([a, t]) => {
        const b = this.mkBtn('<span class="al al-' + a + '"><i></i><i></i><i></i><i></i></span>', t, () => this.setBlock({ textAlign: a }));
        this.alignBtns[a] = b;
        r2.appendChild(b);
      });
    }

    buildInsert() {
      const p = this.panels.insert;
      const t = this.group(p, 'Tables');
      t.appendChild(this.menu('▦ Table', 'Chèn bảng', (m, close) => {
        const label = el('div', 'wd-mh', 'Chọn số dòng × cột');
        const grid = el('div', 'wd-tgrid');
        for (let r = 1; r <= 8; r++) for (let c = 1; c <= 8; c++) {
          const cell = el('div', 'tg');
          cell.dataset.r = r; cell.dataset.c = c;
          cell.addEventListener('mouseenter', () => {
            label.textContent = r + ' dòng × ' + c + ' cột';
            grid.querySelectorAll('.tg').forEach((x) => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c));
          });
          cell.addEventListener('click', () => { this.insertTable(r, c); close(); });
          grid.appendChild(cell);
        }
        m.appendChild(label); m.appendChild(grid);
      }));
      const il = this.group(p, 'Illustrations');
      il.appendChild(this.menu('🖼 Picture', 'Chèn ảnh', (m, close) => {
        (global.WORD_IMAGES || []).forEach((im) => this.mItem(m, `<img src="${im.src}" alt="" class="wd-thumb"> ${im.name}`, () => this.insertImage(im.src), close));
        const up = el('label', 'wd-mi', '📁 Ảnh từ máy tính…');
        const inp = el('input'); inp.type = 'file'; inp.accept = 'image/*'; inp.style.display = 'none';
        inp.addEventListener('change', () => {
          const file = inp.files[0];
          if (!file) return;
          const fr = new FileReader();
          fr.onload = () => { this.insertImage(fr.result); close(); };
          fr.readAsDataURL(file);
        });
        up.appendChild(inp);
        m.appendChild(up);
      }));
      const tx = this.group(p, 'Text');
      tx.appendChild(this.menu('<span class="wa-prev">A</span> WordArt', 'Biến đoạn đang chọn thành WordArt', (m, close) => {
        [1, 2, 3, 4].forEach((n) => this.mItem(m, `<span class="wordart wa-${n} wa-mini">WordArt</span>`, () => this.wordArt(n), close));
        this.mItem(m, 'Bỏ WordArt', () => this.wordArt(0), close);
      }));
      tx.appendChild(this.menu('<b style="font-size:18px">A</b>≡ Drop Cap', 'Chữ cái lớn đầu đoạn', (m, close) => {
        this.mItem(m, 'None – bỏ', () => this.dropCap(0), close);
        [2, 3, 4].forEach((n) => this.mItem(m, `Dropped – ${n} dòng`, () => this.dropCap(n), close));
      }));
      const sy = this.group(p, 'Symbols');
      sy.appendChild(this.menu('Ω Symbol', 'Chèn ký hiệu', (m, close) => {
        const g = el('div', 'wd-sym');
        SYMBOLS.forEach((s) => { const b = el('button', 'sy', s); b.type = 'button'; b.addEventListener('click', () => { this.insertText(s); close(); }); g.appendChild(b); });
        m.appendChild(g);
      }));
    }

    buildLayout() {
      const p = this.panels.layout;
      const g = this.group(p, 'Page Setup');
      g.appendChild(this.menu('▥ Columns', 'Chia cột', (m, close) => {
        this.mItem(m, '▯ One – 1 cột', () => this.columns(1, false), close);
        this.mItem(m, '▯▯ Two – 2 cột', () => this.columns(2, false), close);
        this.mItem(m, '▯▯▯ Three – 3 cột', () => this.columns(3, false), close);
        this.mItem(m, '▯|▯ Two + Line between', () => this.columns(2, true), close);
        this.mItem(m, '▯|▯|▯ Three + Line between', () => this.columns(3, true), close);
      }));
      g.appendChild(el('div', 'wd-note', 'Bôi đen các đoạn cần chia cột rồi chọn số cột.'));
    }

    buildTable() {
      const p = this.panels.table;
      const g = this.group(p, 'Table');
      g.appendChild(this.menu('🪣 Shading', 'Tô nền ô', (m, close) => this.swatches(m, SHADES.concat(['#4472c4', '#70ad47', '#ed7d31', '#ffc000']), (c) => this.cellShade(c), close, 'Không tô')));
      g.appendChild(this.mkBtn('⊞ Gộp ô', 'Merge Cells – gộp các ô đã bôi đen trên cùng một dòng', () => this.mergeCells()));
      g.appendChild(this.mkBtn('⤓ Thêm dòng', 'Insert Below', () => this.addRow()));
      g.appendChild(this.mkBtn('⇥ Thêm cột', 'Insert Right', () => this.addCol()));
      g.appendChild(this.mkBtn('✕ Xóa dòng', 'Delete Row', () => this.delRow()));
      g.appendChild(this.mkBtn('✕ Xóa bảng', 'Delete Table', () => this.delTable()));
    }

    buildPicture() {
      const p = this.panels.picture;
      const g = this.group(p, 'Wrap Text');
      [['inline', 'In Line with Text'], ['left', 'Square – ảnh bên trái'], ['right', 'Square – ảnh bên phải'], ['tb', 'Top and Bottom']].forEach(([w, t]) => {
        g.appendChild(this.mkBtn(t, 'Kiểu bao quanh ảnh', () => this.wrapImage(w)));
      });
      const s = this.group(p, 'Size');
      [25, 35, 50, 100].forEach((n) => s.appendChild(this.mkBtn(n + '%', 'Độ rộng ảnh', () => this.sizeImage(n))));
      s.appendChild(this.mkBtn('✕ Xóa ảnh', 'Delete', () => { if (this.selImg) { this.snap(); this.selImg.remove(); this.selImg = null; this.changed(); this.updateCtx(); } }));
    }

    showTab(id) {
      this.tabsEl.querySelectorAll('.wd-tab').forEach((t) => t.classList.toggle('on', t.dataset.tab === id));
      Object.entries(this.panels).forEach(([k, p]) => { p.style.display = k === id ? 'flex' : 'none'; });
      this.curTab = id;
    }

    bind() {
      document.execCommand('defaultParagraphSeparator', false, 'p');
      this.page.addEventListener('keydown', (e) => this.onKey(e));
      this.page.addEventListener('beforeinput', () => { if (Date.now() - this.lastSnap > 1200) this.snap(); this.lastSnap = Date.now(); });
      this.page.addEventListener('input', () => this.changed());
      this.page.addEventListener('paste', (e) => {
        e.preventDefault();
        this.snap();
        const html = e.clipboardData.getData('text/html') || '';
        // Nội dung chép từ chính trình soạn thảo: giữ định dạng run (giống Ctrl+V trong Word).
        if (html.includes('data-s=')) { this.pasteRuns(cleanHTML(html)); this.changed(); return; }
        const text = (e.clipboardData.getData('text/plain') || '').replace(/\r/g, '');
        text.split('\n').forEach((line, i) => { if (i) document.execCommand('insertParagraph'); if (line) document.execCommand('insertText', false, line); });
      });
      this.page.addEventListener('click', (e) => {
        const img = e.target.closest('img.w-img');
        this.page.querySelectorAll('img.w-img.sel').forEach((x) => x.classList.remove('sel'));
        this.selImg = img || null;
        if (img) img.classList.add('sel');
        this.updateCtx();
      });
      this.page.addEventListener('dragstart', (e) => { if (e.target.tagName === 'IMG') e.preventDefault(); });
      this.onSelChange = () => { const s = document.getSelection(); if (s.anchorNode && this.page.contains(s.anchorNode)) { this.updateState(); this.updateCtx(); } };
      document.addEventListener('selectionchange', this.onSelChange);
      this.onDocDown = (e) => { if (!e.target.closest || !e.target.closest('.wd-dd')) this.closeMenus(); };
      document.addEventListener('mousedown', this.onDocDown);
    }
    destroy() { document.removeEventListener('selectionchange', this.onSelChange); document.removeEventListener('mousedown', this.onDocDown); }

    // Dán nội dung có định dạng: tách đoạn tại con trỏ rồi chèn các run, không qua insertHTML của trình duyệt.
    pasteRuns(html) {
      let r = this.range();
      if (!r) return;
      if (!r.collapsed) { document.execCommand('delete'); r = this.range(); if (!r) return; }
      const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const block = n && n.closest(BLOCK_SEL);
      if (!block) return;
      const tmp = document.createElement('div');
      tmp.innerHTML = html;
      // Gom nội dung dán thành danh sách đoạn (mỗi đoạn là danh sách run).
      const paras = [];
      let loose = document.createElement('p');
      const flushLoose = () => { if (loose.childNodes.length) { paras.push(flatten(loose)); loose = document.createElement('p'); } };
      const visit = (node) => {
        for (const ch of Array.from(node.childNodes)) {
          if (ch.nodeType === 1 && ['P', 'LI'].includes(ch.tagName)) { flushLoose(); paras.push(flatten(ch)); }
          else if (ch.nodeType === 1 && ['UL', 'OL'].includes(ch.tagName)) visit(ch);
          else loose.appendChild(ch.cloneNode(true));
        }
      };
      visit(tmp);
      flushLoose();
      if (!paras.length) return;
      const off = offsetIn(block, r.startContainer, r.startOffset);
      const items = flatten(block);
      const left = [], right = [];
      let pos = 0;
      for (const it of items) {
        const L = itemLen(it);
        if (pos + L <= off) left.push(it);
        else if (pos >= off) right.push(it);
        else { left.push({ t: 'text', text: it.text.slice(0, off - pos), st: it.st }); right.push({ t: 'text', text: it.text.slice(off - pos), st: it.st }); }
        pos += L;
      }
      const lenOf = (arr) => arr.reduce((s, it) => s + itemLen(it), 0);
      let caretBlock = block, caretOff;
      if (paras.length === 1) {
        rebuild(block, left.concat(paras[0], right));
        caretOff = lenOf(left) + lenOf(paras[0]);
      } else {
        rebuild(block, left.concat(paras[0]));
        let after = block;
        paras.slice(1).forEach((p, i) => {
          const nb = document.createElement(block.tagName.toLowerCase());
          const last = i === paras.length - 2;
          rebuild(nb, last ? p.concat(right) : p);
          after.after(nb);
          after = nb;
          if (last) { caretBlock = nb; caretOff = lenOf(p); }
        });
      }
      rebuild(caretBlock, flatten(caretBlock));
      const b = boundaryAt(caretBlock, caretOff);
      const nr = document.createRange();
      nr.setStart(b[0], b[1]); nr.collapse(true);
      const sel = document.getSelection(); sel.removeAllRanges(); sel.addRange(nr);
    }

    setHTML(html) {
      this.page.innerHTML = html;
      this.page.querySelectorAll(BLOCK_SEL).forEach((b) => { if (!b.querySelector(BLOCK_SEL)) rebuild(b, flatten(b)); });
    }
    getHTML() { return this.page.innerHTML; }

    toast(msg) {
      this.toastBox.textContent = msg;
      this.toastBox.classList.add('show');
      clearTimeout(this.toastT);
      this.toastT = setTimeout(() => this.toastBox.classList.remove('show'), 2600);
    }

    changed() {
      clearTimeout(this.chT);
      this.chT = setTimeout(() => { if (this.opts.onChange) this.opts.onChange(); }, 250);
    }

    // ---------- Hoàn tác ----------
    snap() { this.undoStack.push(this.page.innerHTML); if (this.undoStack.length > 80) this.undoStack.shift(); this.redoStack = []; }
    undo() { const h = this.undoStack.pop(); if (h === undefined) return; this.redoStack.push(this.page.innerHTML); this.page.innerHTML = h; this.lastSnap = 0; this.selImg = null; this.changed(); }
    redo() { const h = this.redoStack.pop(); if (h === undefined) return; this.undoStack.push(this.page.innerHTML); this.page.innerHTML = h; this.selImg = null; this.changed(); }

    // ---------- Vùng chọn ----------
    range() {
      const s = document.getSelection();
      if (!s.rangeCount) return null;
      const r = s.getRangeAt(0);
      return this.page.contains(r.commonAncestorContainer) ? r : null;
    }
    saveSel() { const r = this.range(); this.saved = r ? r.cloneRange() : this.saved; }
    restoreSel() { if (this.saved) { const s = document.getSelection(); s.removeAllRanges(); s.addRange(this.saved); } }
    leafBlocks() { return Array.from(this.page.querySelectorAll(BLOCK_SEL)).filter((b) => !b.querySelector(BLOCK_SEL)); }
    blocksIn(r) {
      const bl = this.leafBlocks();
      const out = bl.filter((b) => r.intersectsNode(b));
      if (!out.length) {
        const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
        const b = n && n.closest(BLOCK_SEL);
        if (b) out.push(b);
      }
      return out;
    }
    // Không bôi đen → áp cho từ tại con trỏ (giống Word).
    expandToWord(r) {
      const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const block = n && n.closest(BLOCK_SEL);
      if (!block) return null;
      const off = offsetIn(block, r.startContainer, r.startOffset);
      const text = flatten(block).map((it) => (it.t === 'text' ? it.text : '\u0001')).join('');
      const isW = (c) => c && /[\p{L}\p{N}_]/u.test(c);
      let a = off, b = off;
      while (a > 0 && isW(text[a - 1])) a--;
      while (b < text.length && isW(text[b])) b++;
      if (a === b) return null;
      const nr = document.createRange();
      const s = boundaryAt(block, a), e = boundaryAt(block, b);
      nr.setStart(s[0], s[1]); nr.setEnd(e[0], e[1]);
      return nr;
    }

    // Áp định dạng ký tự: patchFn(danh sách style đang có) → {khóa: giá trị | null}.
    applyInline(patchFn, opts) {
      let r = this.range() || (this.saved && this.page.contains(this.saved.commonAncestorContainer) ? this.saved : null);
      if (!r) { this.toast('Hãy đặt con trỏ / bôi đen văn bản trong trang trước.'); return; }
      if (r.collapsed) {
        const w = this.expandToWord(r);
        if (!w) { this.toast('Hãy bôi đen đoạn văn bản cần định dạng.'); return; }
        r = w;
      }
      const blocks = this.blocksIn(r);
      if (!blocks.length) return;
      if (!(opts && opts.noSnap)) this.snap();
      const first = blocks[0], last = blocks[blocks.length - 1];
      const startOff = first.contains(r.startContainer) || first === r.startContainer ? offsetIn(first, r.startContainer, r.startOffset) : 0;
      const endOff = last.contains(r.endContainer) || last === r.endContainer ? offsetIn(last, r.endContainer, r.endOffset) : Infinity;
      const segs = blocks.map((b) => {
        const items = flatten(b);
        const total = items.reduce((s, it) => s + itemLen(it), 0);
        return { b, items, s: b === first ? startOff : 0, e: b === last ? Math.min(endOff, total) : total };
      });
      // Chia run tại biên vùng chọn.
      const selected = [];
      for (const sg of segs) {
        const out = [];
        let pos = 0;
        for (const it of sg.items) {
          const L = itemLen(it);
          if (it.t !== 'text') { out.push(it); pos += L; continue; }
          const a = Math.max(sg.s, pos), z = Math.min(sg.e, pos + L);
          if (z <= a) { out.push(it); pos += L; continue; }
          const pre = it.text.slice(0, a - pos), mid = it.text.slice(a - pos, z - pos), post = it.text.slice(z - pos);
          if (pre) out.push({ t: 'text', text: pre, st: Object.assign({}, it.st) });
          const m = { t: 'text', text: mid, st: Object.assign({}, it.st), sel: true };
          out.push(m);
          selected.push(m);
          if (post) out.push({ t: 'text', text: post, st: Object.assign({}, it.st) });
          pos += L;
        }
        sg.out = out;
      }
      if (!selected.length) return;
      const patch = patchFn(selected.map((m) => m.st));
      for (const m of selected) {
        for (const key in patch) { if (patch[key] === null || patch[key] === false || patch[key] === undefined) delete m.st[key]; else m.st[key] = patch[key]; }
      }
      for (const sg of segs) {
        // Gộp lại các run giống nhau.
        const merged = [];
        for (const it of sg.out) {
          const lst = merged[merged.length - 1];
          if (it.t === 'text' && lst && lst.t === 'text' && styleKey(lst.st) === styleKey(it.st)) lst.text += it.text;
          else merged.push(it);
        }
        rebuild(sg.b, merged);
      }
      const s = boundaryAt(first, startOff), e = boundaryAt(last, Math.min(segs[segs.length - 1].e, endOff));
      const nr = document.createRange();
      try { nr.setStart(s[0], s[1]); nr.setEnd(e[0], e[1]); const sel = document.getSelection(); sel.removeAllRanges(); sel.addRange(nr); } catch (err) { /* bỏ qua */ }
      this.page.focus();
      this.updateState();
      this.changed();
    }

    toggle(key) { this.applyInline((sts) => ({ [key]: sts.every((s) => s[key]) ? null : 1 })); }
    toggleVal(key, val) { this.applyInline((sts) => ({ [key]: sts.every((s) => s[key] === val) ? null : val })); }
    setVal(key, val) { this.applyInline(() => ({ [key]: val })); }
    bumpSize(d) {
      this.applyInline((sts) => {
        const cur = sts[0].fs || this.defaultSize();
        const i = SIZES.findIndex((x) => x >= cur);
        const ni = d > 0 ? (SIZES[i] > cur ? i : i + 1) : i - 1;
        return { fs: SIZES[Math.max(0, Math.min(SIZES.length - 1, ni))] };
      });
    }
    clearFormat() { this.applyInline(() => ({ ff: null, fs: null, b: null, i: null, ul: null, st: null, va: null, caps: null, color: null, hl: null, sh: null, bx: null, bxbg: null })); }
    defaultSize() { return toPt(getComputedStyle(this.page).fontSize) || 12; }

    // ---------- Đoạn ----------
    setBlock(css) {
      const r = this.range() || this.saved;
      if (!r) { this.toast('Đặt con trỏ vào đoạn văn trước.'); return; }
      this.snap();
      for (const b of this.blocksIn(r)) {
        Object.assign(b.style, css);
        if (css.textAlign && b.parentElement && b.parentElement.tagName === 'TD' && b.parentElement.children.length === 1) b.parentElement.style.textAlign = css.textAlign;
      }
      this.updateState();
      this.changed();
    }
    toggleIndent() {
      const r = this.range(); if (!r) return;
      const on = this.blocksIn(r).every((b) => b.style.textIndent);
      this.setBlock({ textIndent: on ? '' : '1cm' });
    }
    toggleBlockBorder() {
      const r = this.range(); if (!r) return;
      const on = this.blocksIn(r).every((b) => b.style.borderStyle === 'solid');
      this.setBlock(on ? { border: '', padding: '' } : { border: '1px solid #000', padding: '4px 6px' });
    }
    bullets(kind) {
      const r = this.range(); if (!r) { this.toast('Đặt con trỏ vào đoạn văn trước.'); return; }
      this.snap();
      const inList = this.blocksIn(r).some((b) => b.tagName === 'LI');
      if (!kind) { if (inList) document.execCommand('insertUnorderedList'); this.changed(); return; }
      if (!inList || this.blocksIn(r)[0].parentElement.tagName === 'OL') document.execCommand('insertUnorderedList');
      const r2 = this.range();
      if (r2) this.blocksIn(r2).forEach((b) => { const ul = b.closest('ul'); if (ul) ul.className = 'bl-' + kind; });
      this.changed();
    }
    numbering() { if (!this.range()) return; this.snap(); document.execCommand('insertOrderedList'); this.changed(); }

    topBlockOf(n) {
      let x = n.nodeType === 1 ? n : n.parentElement;
      while (x && x.parentElement !== this.page && x !== this.page) x = x.parentElement;
      return x === this.page ? null : x;
    }

    columns(n, line) {
      const r = this.range() || this.saved;
      if (!r) { this.toast('Bôi đen các đoạn cần chia cột.'); return; }
      const startTop = this.topBlockOf(r.startContainer), endTop = this.topBlockOf(r.endContainer);
      if (!startTop) return;
      this.snap();
      // Đang trong khung cột → đổi số cột hoặc bỏ.
      if (startTop.classList.contains('w-cols') && startTop === endTop) {
        if (n === 1) { while (startTop.firstChild) this.page.insertBefore(startTop.firstChild, startTop); startTop.remove(); }
        else this.setCols(startTop, n, line);
        this.changed(); return;
      }
      if (n === 1) { this.changed(); return; }
      const tops = [];
      let x = startTop;
      while (x) { tops.push(x); if (x === endTop) break; x = x.nextElementSibling; }
      const box = el('div', 'w-cols');
      this.page.insertBefore(box, startTop);
      for (const t of tops) {
        if (t.classList.contains('w-cols')) { while (t.firstChild) box.appendChild(t.firstChild); t.remove(); }
        else box.appendChild(t);
      }
      this.setCols(box, n, line);
      if (!box.nextElementSibling) { const p = el('p'); p.appendChild(el('br')); this.page.appendChild(p); }
      this.changed();
    }
    setCols(box, n, line) {
      box.dataset.cols = n;
      box.dataset.line = line ? '1' : '0';
      box.style.columnCount = n;
      box.style.columnRule = line ? '1px solid #555' : '';
    }

    dropCap(lines) {
      const r = this.range() || this.saved;
      if (!r) { this.toast('Đặt con trỏ vào đoạn cần tạo Drop Cap.'); return; }
      const block = this.blocksIn(r)[0];
      if (!block) return;
      this.snap();
      const old = block.querySelector('.dropcap');
      if (old) {
        const items = flatten(block);
        const i = items.findIndex((it) => it.node === old);
        let st = {};
        try { st = JSON.parse(old.dataset.s || '{}'); } catch (e) { st = {}; }
        items[i] = { t: 'text', text: old.textContent, st };
        rebuild(block, items);
        rebuild(block, flatten(block)); // gộp chữ cái vừa trả về với run phía sau
      }
      if (lines) {
        const items = flatten(block);
        const i = items.findIndex((it) => it.t === 'text' && it.text.trim());
        if (i < 0) { this.toast('Đoạn này chưa có chữ.'); return; }
        const it = items[i];
        const lead = it.text.match(/^\s*/)[0];
        const letter = it.text[lead.length];
        const dc = el('span', 'dropcap', letter);
        dc.contentEditable = 'false';
        dc.dataset.lines = lines;
        dc.dataset.s = JSON.stringify(it.st);
        if (it.st.ff) dc.style.fontFamily = `"${it.st.ff}"`;
        if (it.st.color) dc.style.color = it.st.color;
        const rest = { t: 'text', text: it.text.slice(lead.length + 1), st: it.st };
        items.splice(i, 1, { t: 'atom', node: dc }, rest);
        rebuild(block, items);
      }
      this.changed();
    }

    wordArt(n) {
      const r = this.range() || this.saved;
      if (!r) { this.toast('Đặt con trỏ vào đoạn tiêu đề trước.'); return; }
      this.snap();
      for (const b of this.blocksIn(r)) {
        b.classList.remove('wordart', 'wa-1', 'wa-2', 'wa-3', 'wa-4');
        if (n) b.classList.add('wordart', 'wa-' + n);
      }
      this.changed();
    }

    insertText(s) {
      if (!this.range()) { this.restoreSel(); if (!this.range()) { this.toast('Đặt con trỏ vào trang trước.'); return; } }
      this.snap();
      this.page.focus();
      document.execCommand('insertText', false, s);
      this.changed();
    }

    insertAfterCaretBlock(node) {
      const r = this.range() || this.saved;
      const top = r ? this.topBlockOf(r.startContainer) : null;
      if (top && top.classList.contains('w-cols')) {
        const b = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement).closest(BLOCK_SEL);
        if (b && b.parentElement === top) { top.insertBefore(node, b.nextSibling); return; }
      }
      if (top) this.page.insertBefore(node, top.nextSibling); else this.page.appendChild(node);
    }

    insertTable(rows, cols) {
      this.snap();
      const t = el('table', 'w-table');
      const tb = el('tbody');
      for (let i = 0; i < rows; i++) {
        const tr = el('tr');
        for (let j = 0; j < cols; j++) { const td = el('td'); const p = el('p'); p.appendChild(el('br')); td.appendChild(p); tr.appendChild(td); }
        tb.appendChild(tr);
      }
      t.appendChild(tb);
      this.insertAfterCaretBlock(t);
      if (!t.nextElementSibling) { const p = el('p'); p.appendChild(el('br')); t.after(p); }
      const p0 = t.querySelector('p');
      const sel = document.getSelection(); const nr = document.createRange(); nr.setStart(p0, 0); nr.collapse(true); sel.removeAllRanges(); sel.addRange(nr);
      this.page.focus();
      this.updateCtx();
      this.changed();
    }

    selCells() {
      const r = this.range(); if (!r) return [];
      const cells = Array.from(this.page.querySelectorAll('table.w-table td')).filter((td) => r.intersectsNode(td));
      if (cells.length) return cells;
      const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const td = n && n.closest('td');
      return td ? [td] : [];
    }
    cellShade(c) { const cells = this.selCells(); if (!cells.length) { this.toast('Đặt con trỏ trong bảng.'); return; } this.snap(); cells.forEach((td) => { td.style.backgroundColor = c || ''; }); this.changed(); }
    mergeCells() {
      const cells = this.selCells();
      if (cells.length < 2) { this.toast('Bôi đen từ 2 ô trở lên trên cùng một dòng rồi bấm Gộp ô.'); return; }
      const tr = cells[0].parentElement;
      if (!cells.every((c) => c.parentElement === tr)) { this.toast('Chỉ gộp được các ô trên cùng một dòng.'); return; }
      this.snap();
      const first = cells[0];
      let span = first.colSpan;
      for (const c of cells.slice(1)) {
        span += c.colSpan;
        if (c.textContent.trim()) Array.from(c.childNodes).forEach((n) => first.appendChild(n));
        c.remove();
      }
      first.colSpan = span;
      this.changed();
    }
    tableCols(t) { return Math.max(...Array.from(t.rows).map((row) => Array.from(row.cells).reduce((s, c) => s + c.colSpan, 0))); }
    addRow() {
      const cells = this.selCells(); if (!cells.length) { this.toast('Đặt con trỏ trong bảng.'); return; }
      this.snap();
      const tr = cells[cells.length - 1].parentElement;
      const t = tr.closest('table');
      const nr = el('tr');
      for (let i = 0; i < this.tableCols(t); i++) { const td = el('td'); const p = el('p'); p.appendChild(el('br')); td.appendChild(p); nr.appendChild(td); }
      tr.after(nr);
      this.changed();
    }
    addCol() {
      const cells = this.selCells(); if (!cells.length) { this.toast('Đặt con trỏ trong bảng.'); return; }
      this.snap();
      const idx = cells[0].cellIndex;
      const t = cells[0].closest('table');
      Array.from(t.rows).forEach((row) => {
        const td = el('td'); const p = el('p'); p.appendChild(el('br')); td.appendChild(p);
        const ref = row.cells[Math.min(idx, row.cells.length - 1)];
        ref.after(td);
      });
      this.changed();
    }
    delRow() { const cells = this.selCells(); if (!cells.length) return; this.snap(); const tr = cells[0].parentElement; const t = tr.closest('table'); tr.remove(); if (!t.rows.length) t.remove(); this.changed(); this.updateCtx(); }
    delTable() { const cells = this.selCells(); if (!cells.length) return; this.snap(); cells[0].closest('table').remove(); this.changed(); this.updateCtx(); }

    insertImage(src) {
      this.snap();
      const img = el('img', 'w-img');
      img.src = src;
      img.alt = 'Hình ảnh';
      img.dataset.wrap = 'inline';
      img.style.width = '35%';
      const r = this.range() || this.saved;
      const block = r ? this.blocksIn(r)[0] : null;
      if (block) block.insertBefore(img, block.firstChild);
      else { const p = el('p'); p.appendChild(img); this.page.appendChild(p); }
      this.page.querySelectorAll('img.w-img.sel').forEach((x) => x.classList.remove('sel'));
      img.classList.add('sel');
      this.selImg = img;
      this.updateCtx();
      this.showTab('picture');
      this.changed();
    }
    wrapImage(w) { if (!this.selImg) { this.toast('Bấm chọn một ảnh trước.'); return; } this.snap(); this.selImg.dataset.wrap = w; this.changed(); }
    sizeImage(n) { if (!this.selImg) { this.toast('Bấm chọn một ảnh trước.'); return; } this.snap(); this.selImg.style.width = n + '%'; this.changed(); }

    // ---------- Trạng thái ribbon ----------
    updateState() {
      const r = this.range();
      if (!r) return;
      const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const block = n && n.closest(BLOCK_SEL);
      if (!block) return;
      const probe = r.startContainer.nodeType === 3 ? r.startContainer : (r.startContainer.childNodes[r.startOffset] || r.startContainer);
      const st = styleOf(probe, block);
      const cs = getComputedStyle(n);
      this.fontSel.value = st.ff || firstFamily(cs.fontFamily);
      this.sizeSel.value = String(st.fs || toPt(cs.fontSize));
      this.bBtn.classList.toggle('on', !!st.b);
      this.iBtn.classList.toggle('on', !!st.i);
      this.uBtn.classList.toggle('on', !!st.ul);
      const al = block.style.textAlign || 'left';
      Object.entries(this.alignBtns).forEach(([a, b]) => b.classList.toggle('on', a === al));
    }
    updateCtx() {
      const r = this.range();
      const n = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement);
      const inTable = !!(n && n.closest('table.w-table'));
      this.tabsEl.querySelector('[data-tab=table]').classList.toggle('show', inTable);
      this.tabsEl.querySelector('[data-tab=picture]').classList.toggle('show', !!this.selImg);
      if (!inTable && this.curTab === 'table') this.showTab('home');
      if (!this.selImg && this.curTab === 'picture') this.showTab('home');
    }

    onKey(e) {
      const ctrl = e.ctrlKey || e.metaKey;
      if (!ctrl) {
        if ((e.key === 'Delete' || e.key === 'Backspace') && this.selImg) { e.preventDefault(); this.snap(); this.selImg.remove(); this.selImg = null; this.updateCtx(); this.changed(); }
        return;
      }
      const key = e.key.toLowerCase();
      const map = {
        b: () => this.toggle('b'), i: () => this.toggle('i'), u: () => this.toggleVal('ul', 'single'),
        e: () => this.setBlock({ textAlign: 'center' }), l: () => this.setBlock({ textAlign: 'left' }), r: () => this.setBlock({ textAlign: 'right' }), j: () => this.setBlock({ textAlign: 'justify' }),
        z: () => this.undo(), y: () => this.redo(), ']': () => this.bumpSize(1), '[': () => this.bumpSize(-1),
        1: () => this.setBlock({ lineHeight: '1' }), 2: () => this.setBlock({ lineHeight: '2' }), 5: () => this.setBlock({ lineHeight: '1.5' }),
      };
      if (e.shiftKey && (key === '+' || key === '=' || e.code === 'Equal')) { e.preventDefault(); this.toggleVal('va', 'super'); return; }
      if (e.shiftKey && key === 'k') { e.preventDefault(); this.toggleVal('caps', 'small'); return; }
      if (e.shiftKey && key === 'a') { e.preventDefault(); this.toggleVal('caps', 'all'); return; }
      if (!e.shiftKey && (key === '=' || e.code === 'Equal')) { e.preventDefault(); this.toggleVal('va', 'sub'); return; }
      if (map[key] && !e.altKey && !e.shiftKey) { e.preventDefault(); map[key](); }
    }
  }

  global.WordEditor = WordEditor;
  global.WordDoc = WordDoc;
})(window);
