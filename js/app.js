/* Ứng dụng: điều hướng, danh sách bài, trang luyện Excel/Word, thi thử, phím tắt, tra cứu. */
(function () {
  'use strict';
  const XL = window.XL, SM = window.SheetModel, XD = window.EXCEL_DATA, WD = window.WORD_DATA, DR = window.DRILLS;
  const main = document.getElementById('main');

  // ---------- Tiện ích ----------
  const store = {
    get(k, d) { try { const v = localStorage.getItem('ontap:' + k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('ontap:' + k, JSON.stringify(v)); } catch (e) { /* bộ nhớ trình duyệt bị chặn: vẫn chạy bình thường */ } },
    del(k) { try { localStorage.removeItem('ontap:' + k); } catch (e) { /* bỏ qua */ } },
  };
  function h(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k];
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (k === 'html') e.innerHTML = v;
      else e.setAttribute(k, v === true ? '' : v);
    }
    for (const kid of kids.flat()) if (kid !== null && kid !== undefined && kid !== false) e.appendChild(typeof kid === 'string' || typeof kid === 'number' ? document.createTextNode(String(kid)) : kid);
    return e;
  }
  const progressOf = (id) => store.get('prog:' + id, { done: [] });
  function setProgress(id, done) { store.set('prog:' + id, { done }); }
  const fmtTime = (s) => String(Math.floor(s / 60)).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0');
  const STATUS = {
    ok: ['✓', 'Đúng'], partial: ['◐', 'Đúng một phần'], wrong: ['✗', 'Chưa đúng'], empty: ['○', 'Chưa làm'],
    const: ['⚠', 'Đúng số nhưng gõ tay – phải dùng công thức'], fn: ['⚠', 'Đúng kết quả nhưng chưa dùng hàm yêu cầu'], fmt: ['⚠', 'Đúng kết quả, chưa đúng định dạng'],
  };

  let cleanup = [];
  function onCleanup(fn) { cleanup.push(fn); }

  // ---------- Điều hướng ----------
  const routes = [
    [/^$/, pageHome],
    [/^excel$/, pageExcelList],
    [/^excel\/(.+)$/, (m) => pageExcel(m[1])],
    [/^word$/, pageWordList],
    [/^word\/(.+)$/, (m) => pageWord(m[1])],
    [/^thi$/, pageExamList],
    [/^thi\/(.+)$/, (m) => pageExam(m[1])],
    [/^phim-tat$/, pageKeys],
    [/^trac-nghiem$/, pageQuiz],
    [/^tra-cuu$/, pageRef],
  ];
  function route() {
    cleanup.forEach((f) => { try { f(); } catch (e) { /* bỏ qua */ } });
    cleanup = [];
    const path = decodeURIComponent(location.hash.replace(/^#\/?/, ''));
    main.innerHTML = '';
    main.className = '';
    const top = path.split('/')[0];
    document.querySelectorAll('.nav a').forEach((a) => a.classList.toggle('on', a.getAttribute('href').replace('#/', '') === top));
    for (const [re, fn] of routes) { const m = re.exec(path); if (m) { fn(m); window.scrollTo(0, 0); return; } }
    pageHome();
  }
  window.addEventListener('hashchange', route);

  // ---------- Trang chủ ----------
  function countDone(list) {
    let done = 0, total = 0;
    for (const ex of list) { total += ex.tasks.length; done += Math.min(progressOf(ex.id).done.length, ex.tasks.length); }
    return { done, total };
  }
  function bar(done, total, cls) {
    const pct = total ? Math.round((done / total) * 100) : 0;
    return h('div', { class: 'bar ' + (cls || '') }, h('div', { class: 'bar-fill', style: 'width:' + pct + '%' }));
  }
  function pageHome() {
    const xs = countDone(XD.all), ws = countDone(WD.all);
    const exams = EXAMS.map((e) => store.get('exam:' + e.id, null)).filter(Boolean);
    const best = exams.length ? Math.max(...exams.map((x) => x.best)) : null;
    main.appendChild(h('section', { class: 'hero' },
      h('div', { class: 'hero-text' },
        h('p', { class: 'eyebrow' }, 'Ứng dụng CNTT – Ôn thi Word & Excel'),
        h('h1', null, 'Luyện trực tiếp, chấm ngay từng câu'),
        h('p', { class: 'lead' }, 'Bảng tính chạy công thức thật, trình soạn thảo có ribbon như Word. Bài lấy từ đề thi thử, bài tập trên lớp và BT ôn tập cuối kỳ. Làm đến đâu, hệ thống báo đúng/sai đến đó.'),
        h('div', { class: 'hero-cta' },
          h('a', { class: 'btn btn-xl', href: '#/excel' }, 'Luyện Excel'),
          h('a', { class: 'btn btn-wd', href: '#/word' }, 'Luyện Word'),
          h('a', { class: 'btn btn-ghost', href: '#/thi' }, 'Thi thử 60 phút'))),
      h('div', { class: 'stats' },
        stat('Excel', xs.done + '/' + xs.total, 'câu đã làm đúng', bar(xs.done, xs.total, 'xl')),
        stat('Word', ws.done + '/' + ws.total, 'yêu cầu đã đạt', bar(ws.done, ws.total, 'wd')),
        stat('Thi thử', best === null ? '–' : best.toFixed(2) + '/10', best === null ? 'chưa thi đề nào' : 'điểm cao nhất', null))));
    main.appendChild(h('section', { class: 'grid3' },
      card('Excel: hàm & công thức', '6 đề/BT ôn thi + 7 bài theo chủ đề + 20 bài luyện hàm nhanh (LEFT, MID, IF, VLOOKUP, HLOOKUP, SUMIF, RANK…). Gõ = rồi bấm ô để chèn tham chiếu, F4 để khóa $.', '#/excel', 'xl'),
      card('Word: định dạng & trình bày', 'Font, gạch đôi, chỉ số trên/dưới, Small Caps, Drop Cap, chia cột, WordArt, bảng gộp ô, ảnh Square… chấm tự động từng ký tự.', '#/word', 'wd'),
      card('Phím tắt & trắc nghiệm', 'Bấm tổ hợp phím thật để nhớ phím tắt; trắc nghiệm "lệnh nằm ở tab nào" và các bẫy hay gặp trong đề.', '#/phim-tat', 'ac')));
    main.appendChild(h('section', { class: 'howto' },
      h('h2', null, 'Cách luyện hiệu quả'),
      h('ol', null,
        h('li', null, h('b', null, 'Luyện hàm nhanh trước'), ' – mỗi bài 1 hàm, 1–2 phút.'),
        h('li', null, h('b', null, 'Làm đề thi thử ở chế độ luyện'), ' – có gợi ý, đánh dấu ô đúng/sai, xem đáp án khi bí.'),
        h('li', null, h('b', null, 'Thi thử 60 phút'), ' – không gợi ý, nộp bài để xem điểm trên thang 10.'),
        h('li', null, h('b', null, 'Thao tác chưa giả lập được'), ' (sắp xếp, lọc nâng cao, vẽ biểu đồ, SmartArt) – mỗi đề ghi rõ để bạn luyện thêm trên Office thật.'))));
  }
  function stat(label, value, sub, extra) {
    return h('div', { class: 'stat' }, h('div', { class: 'stat-label' }, label), h('div', { class: 'stat-value' }, value), h('div', { class: 'stat-sub' }, sub), extra);
  }
  function card(title, text, href, cls) {
    return h('a', { class: 'card card-' + cls, href }, h('h3', null, title), h('p', null, text), h('span', { class: 'card-go' }, 'Bắt đầu →'));
  }

  // ---------- Danh sách bài ----------
  function exItem(ex, base) {
    const p = Math.min(progressOf(ex.id).done.length, ex.tasks.length), n = ex.tasks.length;
    const done = p >= n;
    return h('a', { class: 'ex-item' + (done ? ' done' : ''), href: '#/' + base + '/' + ex.id },
      h('div', { class: 'ex-item-top' }, h('span', { class: 'ex-src' }, ex.source), done ? h('span', { class: 'pill ok' }, '✓ Xong') : null),
      h('div', { class: 'ex-item-title' }, ex.title),
      h('div', { class: 'ex-item-foot' }, h('span', null, p + '/' + n + ' câu'), bar(p, n, base === 'excel' ? 'xl' : 'wd')));
  }
  function pageExcelList() {
    main.appendChild(h('header', { class: 'page-head xl' }, h('h1', null, 'Luyện Excel'), h('p', null, 'Ô vàng là ô cần điền. Bảng tính tính toán thật nên kết quả sai hiện ngay. Viết công thức ở ô đầu rồi kéo nút điền hoặc Ctrl+D cho cả cột.')));
    main.appendChild(h('h2', { class: 'sec' }, 'Đề thi thử & BT ôn tập (Câu 2)'));
    main.appendChild(h('div', { class: 'ex-grid' }, XD.EXAMS.map((e) => exItem(e, 'excel'))));
    main.appendChild(h('h2', { class: 'sec' }, 'Bài tập theo chủ đề'));
    main.appendChild(h('div', { class: 'ex-grid' }, XD.LESSONS.map((e) => exItem(e, 'excel'))));
    main.appendChild(h('h2', { class: 'sec' }, 'Luyện hàm nhanh'));
    main.appendChild(h('div', { class: 'chips' }, XD.DRILLS.map((e) => {
      const done = progressOf(e.id).done.length >= e.tasks.length;
      return h('a', { class: 'chip' + (done ? ' done' : ''), href: '#/excel/' + e.id }, h('b', null, e.fn), ' ', e.title, done ? ' ✓' : '');
    })));
  }
  function pageWordList() {
    main.appendChild(h('header', { class: 'page-head wd' }, h('h1', null, 'Luyện Word'), h('p', null, 'Văn bản đã gõ sẵn; bạn chỉ trình bày. Danh sách yêu cầu bên trái tự chuyển ✓ khi đạt.')));
    main.appendChild(h('h2', { class: 'sec' }, 'Bài tập định dạng'));
    main.appendChild(h('div', { class: 'ex-grid' }, WD.LESSONS.map((e) => exItem(e, 'word'))));
    main.appendChild(h('h2', { class: 'sec' }, 'Đề thi thử & BT ôn tập (Câu 1)'));
    main.appendChild(h('div', { class: 'ex-grid' }, WD.EXAMS.map((e) => exItem(e, 'word'))));
  }

  // ---------- Khung luyện Excel ----------
  function mountExcel(host, ex, opts) {
    opts = opts || {};
    const key = opts.key || 'xl:' + ex.id;
    const sheet = SM.buildSheet(ex);
    const saved = store.get(key, null);
    if (saved) sheet.loadUserData(saved);
    const solved = SM.solvedSheet(ex);
    let reveal = !opts.exam;
    let results = {};

    const side = h('aside', { class: 'ex-side' });
    const sheetHost = h('div', { class: 'ex-sheet' });
    const layout = h('div', { class: 'ex-layout' }, side, h('section', { class: 'ex-main' }, sheetHost,
      h('div', { class: 'tips' }, h('b', null, 'Mẹo: '), 'gõ ', h('kbd', null, '='), ' rồi bấm vào ô để chèn tham chiếu • ', h('kbd', null, 'F4'), ' khóa $ • kéo ô vuông nhỏ ở góc vùng chọn để điền • ', h('kbd', null, 'Ctrl'), '+', h('kbd', null, 'D'), ' điền xuống • ', h('kbd', null, 'Ctrl'), '+', h('kbd', null, 'Enter'), ' nhập cho cả vùng chọn • dấu phẩy hay chấm phẩy giữa các đối số đều được.')));
    host.appendChild(layout);

    const head = h('div', { class: 'side-head' },
      h('div', { class: 'ex-src' }, ex.source),
      h('h1', null, ex.title),
      opts.exam ? null : h('div', { class: 'side-prog' }, h('span', { class: 'prog-text' }), bar(0, 1, 'xl')));
    side.appendChild(head);
    const list = h('div', { class: 'tasks' });
    side.appendChild(list);
    const cards = {};
    let view = null;
    for (const t of ex.tasks) {
      const hint = h('div', { class: 'task-hint', hidden: true }, '💡 ', t.hint || '');
      const ans = h('div', { class: 'task-ans', hidden: true }, t.answer ? h('code', null, t.answer) : t.values ? 'Giá trị: ' + t.values.join(', ') : 'Dùng nút định dạng trên thanh công cụ.', t.answer ? h('span', { class: 'muted' }, ' (viết ở ' + t.range.split(':')[0] + ' rồi điền cho ' + t.range + ')') : null);
      const msg = h('div', { class: 'task-msg' });
      const icon = h('span', { class: 'task-ic' }, '○');
      const c = h('div', { class: 'task', tabindex: 0 },
        h('div', { class: 'task-top' }, icon, h('span', { class: 'task-label' }, t.label), h('span', { class: 'task-title' }, t.title), h('span', { class: 'task-range' }, t.range), opts.exam && t.pts ? h('span', { class: 'task-pts' }, t.pts + 'đ') : null),
        h('div', { class: 'task-desc' }, t.desc),
        msg,
        opts.exam ? null : h('div', { class: 'task-btns' },
          h('button', { class: 'lnk', type: 'button', onclick: (e) => { e.stopPropagation(); hint.hidden = !hint.hidden; } }, 'Gợi ý'),
          h('button', { class: 'lnk', type: 'button', onclick: (e) => { e.stopPropagation(); ans.hidden = !ans.hidden; } }, 'Xem đáp án')),
        hint, ans);
      c.addEventListener('click', (e) => {
        if (e.target.closest('code, .task-hint, .task-ans')) return;
        const rg = XL.parseRange(t.range);
        view.select(rg.r1, rg.c1, t.type === 'format' ? rg.r2 : rg.r1, t.type === 'format' ? rg.c2 : rg.c1, rg.r1, rg.c1);
        view.wrap.focus({ preventScroll: true });
      });
      cards[t.id] = { c, icon, msg };
      list.appendChild(c);
    }
    if (ex.extra && ex.extra.length) {
      side.appendChild(h('div', { class: 'extra' }, h('div', { class: 'extra-h' }, 'Luyện thêm trên Excel thật'), h('ul', null, ex.extra.map((x) => h('li', null, x)))));
    }
    const doneBanner = h('div', { class: 'done-banner', hidden: true }, '🎉 Hoàn thành bài này!', opts.next ? h('a', { class: 'btn btn-xl btn-sm', href: opts.next }, 'Bài tiếp theo →') : null);
    side.appendChild(doneBanner);
    if (!opts.exam) {
      side.appendChild(h('div', { class: 'side-foot' },
        h('button', { class: 'btn btn-ghost btn-sm', type: 'button', onclick: () => { if (confirm('Xóa toàn bộ bài làm của bài này?')) { store.del(key); view.destroy(); host.removeChild(layout); mountExcel(host, ex, opts); } } }, 'Làm lại từ đầu')));
    }

    view = new window.SheetView(sheetHost, sheet, { widths: ex.widths, onChange: () => { store.set(key, sheet.userData()); update(); } });
    onCleanup(() => view.destroy());

    function update() {
      results = SM.grade(sheet, ex, solved);
      const marks = new Map();
      let okN = 0;
      for (const t of ex.tasks) {
        const r = results[t.id];
        if (r.status === 'ok') okN++;
        const cd = cards[t.id];
        if (reveal) {
          cd.icon.textContent = STATUS[r.status][0];
          cd.c.className = 'task st-' + r.status;
          cd.msg.textContent = r.status === 'empty' || r.status === 'ok' ? '' : r.status === 'fn' ? 'Cần dùng hàm ' + r.missingFn + '.' : r.status === 'partial' ? r.okCount + '/' + r.total + ' ô đúng.' : STATUS[r.status][1] + '.';
          for (const cell of r.cells) {
            if (!cell.filled) continue;
            const exp = solved.display(cell.r, cell.c);
            marks.set(cell.r + ',' + cell.c, { ok: cell.ok && !cell.msg, tip: cell.ok ? (cell.msg || 'Đúng') : 'Chưa đúng. Kết quả mong đợi: ' + (exp === '' ? '(trống)' : exp) });
          }
        } else {
          const filled = r.filled > 0;
          cd.icon.textContent = filled ? '●' : '○';
          cd.c.className = 'task' + (filled ? ' st-filled' : '');
          cd.msg.textContent = '';
        }
      }
      view.setMarks(marks);
      if (!opts.exam) {
        head.querySelector('.prog-text').textContent = okN + '/' + ex.tasks.length + ' câu đúng';
        head.querySelector('.bar-fill').style.width = Math.round((okN / ex.tasks.length) * 100) + '%';
        const prev = progressOf(ex.id).done;
        const done = ex.tasks.filter((t) => results[t.id].status === 'ok').map((t) => t.id);
        setProgress(ex.id, Array.from(new Set(prev.concat(done))));
        doneBanner.hidden = okN !== ex.tasks.length;
      }
    }
    update();
    return {
      score() { results = SM.grade(sheet, ex, solved); const s = SM.score(ex, results); return { got: s.got, max: s.max, results }; },
      reveal() { reveal = true; update(); },
      view,
    };
  }

  // ---------- Khung luyện Word ----------
  function mountWord(host, ex, opts) {
    opts = opts || {};
    const key = opts.key || 'wd:' + ex.id;
    let reveal = !opts.exam;
    const side = h('aside', { class: 'ex-side' });
    const edHost = h('div', { class: 'ex-word' });
    host.appendChild(h('div', { class: 'ex-layout' }, side, h('section', { class: 'ex-main' }, edHost)));
    const head = h('div', { class: 'side-head' },
      h('div', { class: 'ex-src' }, ex.source),
      h('h1', null, ex.title),
      ex.intro ? h('p', { class: 'intro' }, ex.intro) : null,
      opts.exam ? null : h('div', { class: 'side-prog' }, h('span', { class: 'prog-text' }), bar(0, 1, 'wd')));
    side.appendChild(head);
    const list = h('div', { class: 'tasks' });
    side.appendChild(list);
    const cards = {};
    for (const t of ex.tasks) {
      const hint = h('div', { class: 'task-hint', hidden: true }, '💡 ', t.hint || '');
      const icon = h('span', { class: 'task-ic' }, '○');
      const c = h('div', { class: 'task' },
        h('div', { class: 'task-top' }, icon, h('span', { class: 'task-title' }, t.title), opts.exam && t.pts ? h('span', { class: 'task-pts' }, t.pts + 'đ') : null),
        opts.exam ? null : h('div', { class: 'task-btns' }, h('button', { class: 'lnk', type: 'button', onclick: () => { hint.hidden = !hint.hidden; } }, 'Cách làm')),
        hint);
      cards[t.id] = { c, icon };
      list.appendChild(c);
    }
    const doneBanner = h('div', { class: 'done-banner', hidden: true }, '🎉 Đạt tất cả yêu cầu!', opts.next ? h('a', { class: 'btn btn-wd btn-sm', href: opts.next }, 'Bài tiếp theo →') : null);
    side.appendChild(doneBanner);
    if (!opts.exam) {
      side.appendChild(h('div', { class: 'side-foot' },
        h('button', { class: 'btn btn-ghost btn-sm', type: 'button', onclick: () => { if (confirm('Khôi phục văn bản gốc? Bài làm hiện tại sẽ mất (vẫn có thể Ctrl+Z).')) { store.del(key); editor.snap(); editor.setHTML(ex.html); update(); } } }, 'Làm lại từ đầu')));
    }
    const saved = store.get(key, null);
    const editor = new window.WordEditor(edHost, { html: saved || ex.html, onChange: () => { store.set(key, editor.getHTML()); update(); } });
    onCleanup(() => editor.destroy());

    function evaluate() {
      const doc = new window.WordDoc(editor.page);
      const res = {};
      for (const t of ex.tasks) { try { res[t.id] = !!t.check(doc); } catch (e) { res[t.id] = false; } }
      return res;
    }
    function update() {
      const res = evaluate();
      let okN = 0;
      for (const t of ex.tasks) {
        const ok = res[t.id];
        if (ok) okN++;
        const cd = cards[t.id];
        if (reveal) { cd.icon.textContent = ok ? '✓' : '○'; cd.c.className = 'task ' + (ok ? 'st-ok' : 'st-empty'); }
      }
      if (!opts.exam) {
        head.querySelector('.prog-text').textContent = okN + '/' + ex.tasks.length + ' yêu cầu đạt';
        head.querySelector('.bar-fill').style.width = Math.round((okN / ex.tasks.length) * 100) + '%';
        const prev = progressOf(ex.id).done;
        setProgress(ex.id, Array.from(new Set(prev.concat(ex.tasks.filter((t) => res[t.id]).map((t) => t.id)))));
        doneBanner.hidden = okN !== ex.tasks.length;
      }
      return res;
    }
    update();
    return {
      score() {
        const res = evaluate();
        let got = 0, max = 0;
        for (const t of ex.tasks) { max += t.pts || 1; if (res[t.id]) got += t.pts || 1; }
        return { got, max, res };
      },
      reveal() { reveal = true; update(); },
      editor,
    };
  }

  function nextOf(list, id, base) {
    const i = list.findIndex((e) => e.id === id);
    return i >= 0 && i < list.length - 1 ? '#/' + base + '/' + list[i + 1].id : null;
  }
  function crumbs(base, label, title) {
    return h('nav', { class: 'crumbs' }, h('a', { href: '#/' + base }, label), ' / ', h('span', null, title));
  }
  function pageExcel(id) {
    const ex = XD.all.find((e) => e.id === id);
    if (!ex) { location.hash = '#/excel'; return; }
    main.className = 'wide';
    main.appendChild(crumbs('excel', 'Excel', ex.title));
    const list = ex.group === 'drill' ? XD.DRILLS : ex.group === 'exam' ? XD.EXAMS : XD.LESSONS;
    mountExcel(main, ex, { next: nextOf(list, id, 'excel') });
  }
  function pageWord(id) {
    const ex = WD.all.find((e) => e.id === id);
    if (!ex) { location.hash = '#/word'; return; }
    main.className = 'wide';
    main.appendChild(crumbs('word', 'Word', ex.title));
    mountWord(main, ex, { next: nextOf(WD.all, id, 'word') });
  }

  // ---------- Thi thử ----------
  const EXAMS = [
    { id: 'de-2612', title: 'Đề 26.12 – Ca 1 – Đề chẵn', word: 'w-tech', excel: 'x-2612' },
    { id: 'de-d1c1chan', title: 'Đợt 1 – Ca 1 – Đề CHẴN', word: 'w-baden', excel: 'x-d1c1chan' },
    { id: 'de-d1c1le', title: 'Đợt 1 – Ca 1 – Đề LẺ', word: 'w-thu', excel: 'x-d1c1le' },
    { id: 'de-d1c2le', title: 'Đợt 1 – Ca 2 – Đề LẺ', word: 'w-cloud', excel: 'x-d1c2le' },
    { id: 'de-bt', title: 'BT ôn tập cuối kỳ', word: 'w-phanthiet', excel: 'x-bt1' },
  ];
  function pageExamList() {
    main.appendChild(h('header', { class: 'page-head' }, h('h1', null, 'Thi thử 60 phút'), h('p', null, 'Giống phòng thi: Câu 1 Word (5đ) + Câu 2 Excel (5đ), không gợi ý, không báo đúng/sai cho đến khi nộp bài. Bài làm được lưu tự động nếu lỡ tải lại trang.')));
    main.appendChild(h('div', { class: 'ex-grid' }, EXAMS.map((e) => {
      const r = store.get('exam:' + e.id, null);
      return h('a', { class: 'ex-item exam', href: '#/thi/' + e.id },
        h('div', { class: 'ex-item-top' }, h('span', { class: 'ex-src' }, 'Word + Excel • 60 phút'), r ? h('span', { class: 'pill ' + (r.best >= 5 ? 'ok' : 'warn') }, 'Cao nhất ' + r.best.toFixed(2)) : null),
        h('div', { class: 'ex-item-title' }, e.title),
        h('div', { class: 'ex-item-foot' }, h('span', null, (WD.all.find((w) => w.id === e.word) || {}).title + ' • ' + (XD.all.find((x) => x.id === e.excel) || {}).title)));
    })));
  }
  function pageExam(id) {
    const ex = EXAMS.find((e) => e.id === id);
    if (!ex) { location.hash = '#/thi'; return; }
    const wEx = WD.all.find((w) => w.id === ex.word), xEx = XD.all.find((x) => x.id === ex.excel);
    main.className = 'wide';
    const runKey = 'examrun:' + id;
    let run = store.get(runKey, null);
    if (!run || run.submitted) {
      // Màn hình bắt đầu.
      main.appendChild(crumbs('thi', 'Thi thử', ex.title));
      const last = run && run.submitted ? run : null;
      main.appendChild(h('div', { class: 'exam-start' },
        h('h1', null, ex.title),
        h('ul', null, h('li', null, 'Câu 1 (5đ): Word – ', wEx.title), h('li', null, 'Câu 2 (5đ): Excel – ', xEx.title), h('li', null, 'Thời gian 60 phút. Hết giờ hệ thống tự nộp.')),
        last ? h('p', { class: 'muted' }, 'Lần trước: ' + last.total.toFixed(2) + '/10') : null,
        h('div', { class: 'hero-cta' },
          h('button', { class: 'btn btn-xl', type: 'button', onclick: () => { store.del('exam-wd:' + id); store.del('exam-xl:' + id); store.set(runKey, { start: Date.now() }); route(); } }, 'Bắt đầu làm bài'),
          last ? h('button', { class: 'btn btn-ghost', type: 'button', onclick: () => { store.set(runKey, Object.assign({}, last, { review: true, submitted: false })); route(); } }, 'Xem lại bài đã nộp') : null)));
      return;
    }
    const review = !!run.review;
    const LIMIT = 60 * 60;
    const timer = h('span', { class: 'timer' }, review ? 'Xem lại' : '60:00');
    const tabW = h('button', { class: 'etab on', type: 'button' }, 'Câu 1 – Word (5đ)');
    const tabX = h('button', { class: 'etab', type: 'button' }, 'Câu 2 – Excel (5đ)');
    const submit = h('button', { class: 'btn btn-xl btn-sm', type: 'button' }, review ? 'Đóng' : 'Nộp bài');
    main.appendChild(h('div', { class: 'exam-bar' }, h('b', null, ex.title), tabW, tabX, h('span', { class: 'grow' }), timer, submit));
    const paneW = h('div', { class: 'epane' }), paneX = h('div', { class: 'epane', hidden: true });
    main.appendChild(paneW); main.appendChild(paneX);
    const wc = mountWord(paneW, wEx, { exam: true, key: 'exam-wd:' + id });
    const xc = mountExcel(paneX, xEx, { exam: true, key: 'exam-xl:' + id });
    const show = (w) => { paneW.hidden = !w; paneX.hidden = w; tabW.classList.toggle('on', w); tabX.classList.toggle('on', !w); if (!w) xc.view.refresh(); };
    tabW.addEventListener('click', () => show(true));
    tabX.addEventListener('click', () => show(false));
    let tick = null;
    if (review) { wc.reveal(); xc.reveal(); }
    else {
      const upd = () => {
        const left = Math.max(0, LIMIT - Math.floor((Date.now() - run.start) / 1000));
        timer.textContent = fmtTime(left);
        timer.classList.toggle('warn', left < 300);
        if (left === 0) finish(true);
      };
      upd();
      tick = setInterval(upd, 1000);
      onCleanup(() => clearInterval(tick));
    }
    submit.addEventListener('click', () => {
      if (run.review) { store.set(runKey, Object.assign({}, run, { review: false, submitted: true })); location.hash = '#/thi'; return; }
      if (confirm('Nộp bài và xem điểm?')) finish(false);
    });
    function finish(auto) {
      clearInterval(tick);
      const ws = wc.score(), xs = xc.score();
      const wPts = (ws.got / ws.max) * 5, xPts = (xs.got / xs.max) * 5;
      const total = Math.round((wPts + xPts) * 100) / 100;
      const prev = store.get('exam:' + id, null);
      store.set('exam:' + id, { best: Math.max(total, prev ? prev.best : 0), last: total });
      run = { start: run.start, submitted: false, review: true, total };
      store.set(runKey, run);
      wc.reveal(); xc.reveal();
      timer.textContent = 'Đã nộp';
      submit.textContent = 'Đóng';
      const rows = (items) => items.map(([t, ok, pts]) => h('tr', { class: ok ? 'ok' : 'bad' }, h('td', null, ok ? '✓' : '✗'), h('td', null, t), h('td', null, pts)));
      const modal = h('div', { class: 'modal' }, h('div', { class: 'modal-box', role: 'dialog', 'aria-label': 'Kết quả thi thử' },
        h('h2', null, (auto ? 'Hết giờ! ' : '') + 'Điểm: ' + total.toFixed(2) + ' / 10'),
        h('p', { class: 'muted' }, 'Word ' + wPts.toFixed(2) + '/5 • Excel ' + xPts.toFixed(2) + '/5. Phần sắp xếp/lọc/biểu đồ không tính ở đây.'),
        h('div', { class: 'res-cols' },
          h('table', { class: 'res' }, h('caption', null, 'Câu 1 – Word'), h('tbody', null, rows(wEx.tasks.map((t) => [t.title, ws.res[t.id], (t.pts || 1) + 'đ'])))),
          h('table', { class: 'res' }, h('caption', null, 'Câu 2 – Excel'), h('tbody', null, rows(xEx.tasks.map((t) => [t.label + ' ' + t.title, xs.results[t.id].status === 'ok', (t.pts || 1) + 'đ']))))),
        h('div', { class: 'hero-cta' },
          h('button', { class: 'btn btn-xl', type: 'button', onclick: () => modal.remove() }, 'Xem lại bài làm'),
          h('a', { class: 'btn btn-ghost', href: '#/excel/' + xEx.id }, 'Luyện Excel đề này'),
          h('a', { class: 'btn btn-ghost', href: '#/word/' + wEx.id }, 'Luyện Word đề này'))));
      document.body.appendChild(modal);
      onCleanup(() => modal.remove());
    }
  }

  // ---------- Phím tắt ----------
  function pageKeys() {
    const pool = DR.SHORTCUTS.filter((s) => !/^Ctrl\+\d$/.test(s.keys));
    let app = store.get('keys-app', 'Word');
    let queue = [], cur = null, right = 0, total = 0, streak = 0, misses = 0, lock = false;
    main.appendChild(h('header', { class: 'page-head' }, h('h1', null, 'Luyện phím tắt'), h('p', null, 'Đọc thao tác rồi bấm đúng tổ hợp phím. Sai 2 lần sẽ hiện đáp án. Các phím Ctrl+số bị trình duyệt giữ lại nên nằm ở phần trắc nghiệm.')));
    const seg = h('div', { class: 'seg' }, ['Word', 'Excel', 'Cả hai'].map((a) => h('button', { type: 'button', class: a === app ? 'on' : '', onclick: (e) => { app = a; store.set('keys-app', a); seg.querySelectorAll('button').forEach((b) => b.classList.toggle('on', b === e.currentTarget)); e.currentTarget.blur(); restart(); } }, a)));
    const act = h('div', { class: 'k-act' });
    const appTag = h('div', { class: 'k-app' });
    const pressed = h('div', { class: 'k-pressed' }, 'Bấm tổ hợp phím…');
    const fb = h('div', { class: 'k-fb' });
    const score = h('div', { class: 'k-score' });
    const skip = h('button', { class: 'btn btn-ghost btn-sm', type: 'button', onclick: (e) => { e.currentTarget.blur(); if (lock) return; lock = true; fb.textContent = 'Đáp án: ' + cur.keys; fb.className = 'k-fb bad'; total++; streak = 0; setTimeout(next, 1400); } }, 'Bỏ qua');
    const box = h('div', { class: 'k-box' }, appTag, act, pressed, fb, h('div', { class: 'k-foot' }, score, skip));
    main.appendChild(h('div', { class: 'k-wrap' }, seg, box));
    function restart() { queue = []; right = 0; total = 0; streak = 0; next(); }
    function next() {
      lock = false; misses = 0;
      if (!queue.length) queue = DR.shuffle(pool.filter((s) => app === 'Cả hai' || s.app === app));
      cur = queue.pop();
      appTag.textContent = cur.app;
      appTag.className = 'k-app ' + (cur.app === 'Word' ? 'wd' : 'xl');
      act.textContent = cur.act;
      pressed.textContent = 'Bấm tổ hợp phím…';
      fb.textContent = ''; fb.className = 'k-fb';
      score.textContent = 'Đúng ' + right + '/' + total + (streak > 2 ? ' • chuỗi ' + streak + ' 🔥' : '');
    }
    const onKey = (e) => {
      if (['Control', 'Shift', 'Alt', 'Meta'].includes(e.key)) return;
      const combo = DR.comboOf(e);
      if (!combo) return;
      e.preventDefault();
      if (lock) return;
      pressed.innerHTML = '';
      combo.split('+').forEach((k, i) => { if (i) pressed.appendChild(document.createTextNode(' + ')); pressed.appendChild(h('kbd', null, k)); });
      if (combo === cur.keys) {
        lock = true; right++; total++; streak++;
        fb.textContent = '✓ Chính xác!'; fb.className = 'k-fb ok';
        setTimeout(next, 650);
      } else {
        misses++;
        fb.className = 'k-fb bad';
        if (misses >= 2) { lock = true; total++; streak = 0; fb.textContent = '✗ Đáp án đúng: ' + cur.keys; setTimeout(next, 1600); }
        else fb.textContent = '✗ Chưa đúng, thử lại.';
      }
    };
    window.addEventListener('keydown', onKey);
    onCleanup(() => window.removeEventListener('keydown', onKey));
    restart();
    main.appendChild(h('p', { class: 'center' }, h('a', { class: 'btn btn-ghost', href: '#/trac-nghiem' }, 'Sang phần trắc nghiệm →')));
  }

  // ---------- Trắc nghiệm ----------
  function pageQuiz() {
    main.appendChild(h('header', { class: 'page-head' }, h('h1', null, 'Trắc nghiệm nhanh'), h('p', null, 'Mỗi lượt 10 câu: lệnh nằm ở tab nào, phím tắt và các bẫy công thức hay gặp trong đề.')));
    const allKeys = Array.from(new Set(DR.SHORTCUTS.map((s) => s.keys)));
    const fromKeys = DR.SHORTCUTS.map((s) => ({ app: s.app, q: s.app + ': phím tắt "' + s.act + '" là…', a: s.keys, o: DR.shuffle(allKeys.filter((k) => k !== s.keys)).slice(0, 3) }));
    const bank = DR.QUIZ.concat(fromKeys.filter((q) => /^Ctrl\+\d$/.test(q.a)), DR.shuffle(fromKeys).slice(0, 10));
    const box = h('div', { class: 'quiz' });
    main.appendChild(box);
    let qs, i, right;
    function start() { qs = DR.shuffle(bank).slice(0, 10); i = 0; right = 0; show(); }
    function show() {
      box.innerHTML = '';
      if (i >= qs.length) {
        const best = Math.max(store.get('quiz-best', 0), right);
        store.set('quiz-best', best);
        box.appendChild(h('div', { class: 'quiz-end' }, h('h2', null, 'Kết quả: ' + right + '/10'), h('p', { class: 'muted' }, 'Cao nhất: ' + best + '/10'), h('button', { class: 'btn btn-xl', type: 'button', onclick: start }, 'Làm lượt mới')));
        return;
      }
      const q = qs[i];
      const expl = h('div', { class: 'quiz-expl' });
      const nextBtn = h('button', { class: 'btn btn-xl btn-sm', type: 'button', hidden: true, onclick: () => { i++; show(); } }, i === qs.length - 1 ? 'Xem kết quả' : 'Câu tiếp →');
      const btns = DR.shuffle([q.a, ...q.o]).map((o) => h('button', { class: 'opt', type: 'button', 'data-v': o, onclick: (e) => {
        btns.forEach((b) => { b.disabled = true; if (b.dataset.v === q.a) b.classList.add('ok'); });
        if (o === q.a) { right++; expl.textContent = '✓ Chính xác.'; } else { e.currentTarget.classList.add('bad'); expl.textContent = '✗ Đáp án: ' + q.a; }
        nextBtn.hidden = false; nextBtn.focus();
      } }, o));
      box.appendChild(h('div', { class: 'quiz-card' },
        h('div', { class: 'quiz-meta' }, h('span', { class: 'k-app ' + (q.app === 'Word' ? 'wd' : 'xl') }, q.app), h('span', null, 'Câu ' + (i + 1) + '/10 • đúng ' + right)),
        h('h2', null, q.q), h('div', { class: 'opts' }, btns), expl, nextBtn));
    }
    start();
  }

  // ---------- Tra cứu ----------
  function pageRef() {
    const fx = [
      ['Chuỗi', [
        ['LEFT(chuỗi, n)', '=LEFT("HN1U",2) → "HN"', 'Lấy n ký tự bên trái.'],
        ['RIGHT(chuỗi, n)', '=RIGHT("NB3PQ",2) → "PQ"', 'Lấy n ký tự bên phải.'],
        ['MID(chuỗi, vị trí, n)', '=MID("MY-2B",4,1) → "2"', 'Dấu "-" cũng là 1 ký tự.'],
        ['VALUE(chuỗi)', '=VALUE(RIGHT("1IN210",3)) → 210', 'Đổi chuỗi số thành số để tính/so sánh với số.'],
        ['& (nối chuỗi)', '=A4&" "&B4', 'Thêm dấu cách " " giữa hai phần.'],
      ]],
      ['Logic', [
        ['IF(đk, đúng, sai)', '=IF(F4>50,10%*E4*F4,0)', 'Chữ phải đặt trong "". Để trống: "".'],
        ['IF lồng', '=IF(B2>=8,"Giỏi",IF(B2>=6.5,"Khá","TB"))', 'Xét từ mức cao xuống thấp.'],
        ['AND(đk1, đk2)', '=IF(AND(C4="Máy in",E4>DATE(2024,12,20)),…)', 'Đúng khi tất cả đều đúng.'],
        ['OR(đk1, đk2)', '=IF(OR(AND(…),AND(…)),…)', 'Đúng khi có ít nhất một điều kiện đúng.'],
      ]],
      ['Dò tìm', [
        ['VLOOKUP(giá trị, $bảng$, cột, 0)', '=VLOOKUP(LEFT(B4,2),$A$18:$D$21,2,0)', 'Bảng DỌC: mã ở cột đầu. Luôn để 0 và khóa $ bảng.'],
        ['HLOOKUP(giá trị, $bảng$, hàng, 0)', '=HLOOKUP(C4,$E$17:$G$19,2,0)', 'Bảng NGANG: mã ở dòng đầu.'],
        ['Cột/hàng đổi theo điều kiện', '=VLOOKUP(…, IF(D4="Loại 1",3,4), 0)', 'Mẫu hay gặp: giá loại 1/loại 2, 3 ngày/5 ngày.'],
      ]],
      ['Thống kê', [
        ['SUMIF(vùng đk, đk, vùng tổng)', '=SUMIF($C$4:$C$13,F18,$H$4:$H$13)', 'Tổng có 1 điều kiện.'],
        ['SUMIFS(vùng tổng, vùng đk1, đk1, …)', '=SUMIFS($H$4:$H$12,$D$4:$D$12,A24,$F$4:$F$12,">50")', 'Nhiều điều kiện; vùng tổng đứng ĐẦU.'],
        ['COUNTIF / COUNTIFS', '=COUNTIF(B4:B12,"MITSUBISHI*")', '* thay nhiều ký tự, ? thay một ký tự.'],
        ['AVERAGEIF', '=AVERAGEIF($C$4:$C$13,"Nam",D4:D13)', 'Trung bình có điều kiện.'],
        ['RANK(số, $vùng$)', '=RANK(G4,$G$4:$G$13)', 'Mặc định giảm dần (cao nhất hạng 1).'],
        ['MAX / MIN / AVERAGE / SUM', '=MAX(D4:D13)', 'Alt+= để AutoSum.'],
      ]],
      ['Số & ngày', [
        ['ROUND(số, n)', '=ROUND((D4*2+E4*2+F4)/5,2)', 'Làm tròn n chữ số thập phân.'],
        ['INT / MOD', '=INT(F4/7), =MOD(F4,7)', 'Phần nguyên / phần dư (số tuần, ngày lẻ).'],
        ['DATE(năm, tháng, ngày)', 'E4>DATE(2024,12,20)', 'So sánh ngày luôn dùng DATE.'],
        ['DAY / MONTH / YEAR / TODAY', '=YEAR(TODAY())-D4', 'Tính tuổi theo ngày hệ thống.'],
      ]],
    ];
    const refs = [
      ['$A$1', 'Khóa cả cột và dòng – dùng cho bảng tra, ô tỷ giá.'],
      ['$A1', 'Khóa cột – kéo sang phải vẫn giữ cột A (bảng thống kê 2 chiều).'],
      ['A$1', 'Khóa dòng – kéo xuống vẫn giữ dòng 1 (tiêu đề cột của bảng thống kê).'],
      ['F4', 'Bấm khi con trỏ đang ở tham chiếu để xoay vòng A1 → $A$1 → A$1 → $A1.'],
    ];
    const errs = [['#N/A', 'VLOOKUP/HLOOKUP không tìm thấy (sai mã, thừa dấu cách, so chuỗi "1" với số 1).'], ['#VALUE!', 'Sai kiểu dữ liệu (cộng chữ với số, cột trả về < 1).'], ['#REF!', 'Cột/hàng trả về vượt ngoài bảng dò.'], ['#NAME?', 'Sai tên hàm hoặc quên "" quanh chữ.'], ['#DIV/0!', 'Chia cho 0 hoặc ô trống.']];
    const word = [
      ['Font: kiểu, cỡ, đậm, nghiêng, gạch chân', 'Home → Font (Ctrl+D mở hộp thoại)'],
      ['Gạch đôi, chỉ số trên/dưới, Small Caps', 'Hộp thoại Font → Effects'],
      ['Căn lề, giãn dòng, thụt đầu dòng', 'Home → Paragraph'],
      ['Bullets / Numbering, ký hiệu riêng', 'Home → Bullets → Define New Bullet → Symbol'],
      ['Viền & tô nền chữ / đoạn', 'Home → Borders → Borders and Shading (Apply to: Text / Paragraph)'],
      ['Drop Cap', 'Insert → Drop Cap → Drop Cap Options (Lines to drop)'],
      ['WordArt', 'Insert → WordArt; chỉnh Text Effects → Transform'],
      ['Chia cột', 'Page Layout → Columns → More Columns (Line between)'],
      ['Ảnh & kiểu bao chữ', 'Insert → Picture → Wrap Text → Square'],
      ['Bảng, gộp ô, tô nền ô', 'Insert → Table; Layout → Merge Cells; Design → Shading'],
      ['Tab có dấu chấm dẫn', 'Home → Paragraph → Tabs → Leader'],
      ['Watermark, màu trang, viền trang', 'Page Layout (Design) → Watermark / Page Color / Page Borders'],
      ['Trộn thư', 'Mailings → Start Mail Merge → Select Recipients → Insert Merge Field'],
    ];
    main.appendChild(h('header', { class: 'page-head' }, h('h1', null, 'Tra cứu nhanh'), h('p', null, 'Tóm tắt từ bài giảng Excel/Word – mở khi bí trong lúc luyện.')));
    main.appendChild(h('h2', { class: 'sec' }, 'Hàm Excel hay ra thi'));
    for (const [g, rows] of fx) {
      main.appendChild(h('div', { class: 'ref-group' }, h('h3', null, g), h('div', { class: 'tbl-wrap' }, h('table', { class: 'ref' },
        h('thead', null, h('tr', null, h('th', null, 'Cú pháp'), h('th', null, 'Ví dụ trong đề'), h('th', null, 'Ghi nhớ'))),
        h('tbody', null, rows.map((r) => h('tr', null, h('td', null, h('code', null, r[0])), h('td', null, h('code', null, r[1])), h('td', null, r[2]))))))));
    }
    main.appendChild(h('div', { class: 'grid2' },
      h('div', { class: 'ref-group' }, h('h3', null, 'Tham chiếu $'), h('table', { class: 'ref' }, h('tbody', null, refs.map((r) => h('tr', null, h('td', null, h('code', null, r[0])), h('td', null, r[1])))))),
      h('div', { class: 'ref-group' }, h('h3', null, 'Lỗi thường gặp'), h('table', { class: 'ref' }, h('tbody', null, errs.map((r) => h('tr', null, h('td', null, h('code', null, r[0])), h('td', null, r[1]))))))));
    main.appendChild(h('h2', { class: 'sec' }, 'Thao tác Word'));
    main.appendChild(h('div', { class: 'tbl-wrap' }, h('table', { class: 'ref' }, h('thead', null, h('tr', null, h('th', null, 'Việc cần làm'), h('th', null, 'Đường dẫn lệnh'))), h('tbody', null, word.map((r) => h('tr', null, h('td', null, r[0]), h('td', null, r[1])))))));
  }

  route();
})();
