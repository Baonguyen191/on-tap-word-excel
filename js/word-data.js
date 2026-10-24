/* Bài tập Word: nội dung lấy từ "Bài tập MS Word", "BT ôn tập cuối kỳ" và các đề thi thử. */
(function (global) {
  'use strict';

  // ---------- Ảnh minh họa (SVG nội tuyến, không cần mạng) ----------
  const svg = (s) => 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(s);
  global.WORD_IMAGES = [
    { name: 'Núi non', src: svg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><defs><linearGradient id="s" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7ec8f0"/><stop offset="1" stop-color="#e8f6ff"/></linearGradient></defs><rect width="320" height="200" fill="url(#s)"/><circle cx="250" cy="50" r="22" fill="#ffd34d"/><path d="M0 170 L90 60 L150 130 L210 40 L320 170 Z" fill="#3d6b4f"/><path d="M190 70 L210 40 L232 72 Z" fill="#fff"/><rect y="165" width="320" height="35" fill="#6aa84f"/><path d="M150 200 L170 165 L180 165 L200 200Z" fill="#9b8b6e"/></svg>') },
    { name: 'Công nghệ', src: svg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><rect width="320" height="200" fill="#0b1e3d"/><g stroke="#2ec4ff" stroke-width="2" fill="none" opacity=".8"><path d="M20 40 H120 V90 H200"/><path d="M40 160 H140 V120 H300"/><path d="M220 20 V70 H300"/><circle cx="200" cy="90" r="6"/><circle cx="300" cy="120" r="6"/><circle cx="20" cy="40" r="6"/></g><text x="160" y="112" font-family="Arial" font-size="26" font-weight="700" fill="#fff" text-anchor="middle">TECHNOLOGY</text><text x="160" y="140" font-family="Arial" font-size="16" fill="#2ec4ff" text-anchor="middle">TRENDS</text></svg>') },
    { name: 'Mùa thu', src: svg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><rect width="320" height="200" fill="#fbe7c6"/><rect y="150" width="320" height="50" fill="#c9a26b"/><rect x="60" y="70" width="14" height="90" fill="#6b4226"/><circle cx="67" cy="62" r="46" fill="#e07a1f"/><circle cx="40" cy="80" r="26" fill="#f2a541"/><rect x="220" y="80" width="12" height="80" fill="#6b4226"/><circle cx="226" cy="70" r="40" fill="#d1495b"/><g fill="#e07a1f"><circle cx="140" cy="170" r="5"/><circle cx="170" cy="182" r="4"/><circle cx="120" cy="186" r="4"/></g></svg>') },
    { name: 'Đám mây', src: svg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><rect width="320" height="200" fill="#e9f2ff"/><path d="M90 140 a40 40 0 0 1 10-78 a55 55 0 0 1 104 8 a36 36 0 0 1 26 70 Z" fill="#4a90e2"/><g fill="#fff"><rect x="120" y="100" width="70" height="8" rx="4"/><rect x="120" y="116" width="50" height="8" rx="4"/></g><g stroke="#4a90e2" stroke-width="3"><line x1="110" y1="140" x2="110" y2="175"/><line x1="160" y1="140" x2="160" y2="175"/><line x1="210" y1="140" x2="210" y2="175"/></g><g fill="#2c5aa0"><rect x="96" y="175" width="28" height="18" rx="3"/><rect x="146" y="175" width="28" height="18" rx="3"/><rect x="196" y="175" width="28" height="18" rx="3"/></g></svg>') },
    { name: 'Biển', src: svg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><rect width="320" height="120" fill="#8fd3fe"/><circle cx="60" cy="45" r="20" fill="#ffd34d"/><rect y="110" width="320" height="50" fill="#1f8ac0"/><path d="M0 120 q20 -8 40 0 t40 0 t40 0 t40 0 t40 0 t40 0 t40 0 t40 0" stroke="#fff" fill="none" stroke-width="3"/><rect y="155" width="320" height="45" fill="#f3d9a4"/><rect x="250" y="95" width="6" height="65" fill="#7a4e2d"/><path d="M253 95 q-30 5 -40 20 M253 95 q30 5 40 20 M253 95 q-10 -20 -30 -22 M253 95 q12 -20 32 -20" stroke="#2f7d32" stroke-width="6" fill="none"/></svg>') },
  ];

  // ---------- Bộ điều kiện dùng chung ----------
  const P = {
    font: (f) => (s) => s.ff === f.toLowerCase(),
    size: (n) => (s) => s.fs === n,
    sizeMin: (n) => (s) => s.fs >= n,
    bold: (s) => s.b,
    italic: (s) => s.i,
    ul: (t) => (s) => (t ? s.ul === t : !!s.ul),
    strike: (t) => (s) => s.st === t,
    caps: (t) => (s) => s.caps === t,
    shadow: (s) => s.sh,
    box: (s) => s.bx,
    shaded: (s) => !!s.hl,
    colored: (s) => !!s.color && !/^(#000000|#000|black|rgb\(0, 0, 0\))$/i.test(s.color),
    and: (...ps) => (s) => ps.every((p) => p(s)),
  };
  const align = (d, needle) => { const b = d.blockOf(needle); if (!b) return ''; const a = getComputedStyle(b).textAlign; return a === 'start' ? 'left' : a; };
  const isWordArt = (d, needle) => { const b = d.blockOf(needle); return !!(b && b.classList.contains('wordart')); };
  const dropLines = (d, needle) => { const dc = d.dropcapOf(needle); return dc ? +dc.dataset.lines : 0; };
  const squareImg = (d) => d.images().some((im) => im.dataset.wrap === 'left' || im.dataset.wrap === 'right');
  const ordered = (d, a, b) => { const x = d.indexOf(a), y = d.indexOf(b); return x >= 0 && y >= 0 && x < y; };
  // Ký tự sau tiền tố (bỏ tiền tố dài n ký tự không tính khoảng trắng) thỏa điều kiện.
  const tailAll = (d, needle, skip, pred) => { const c = d.charsOf(needle); return !!c && c.length > skip && c.slice(skip).every((x) => pred(x.st)); };
  const html = (paras) => paras.map((p) => (typeof p === 'string' ? `<p>${p}</p>` : p.html)).join('');
  const note = (t) => ({ html: `<p class="w-note">${t}</p>` });

  const POEM = [
    'Một mùa thu trước, mỗi hoàng hôn<br>Nhặt cánh hoa rơi chẳng thấy buồn,<br>Nhuộm ánh nắng tà qua mái tóc,<br>Tôi chờ người đến với yêu đương.',
    'Người ấy thường hay vuốt tóc tôi,<br>Thở dài trong lúc thấy tôi vui,<br>Bảo rằng: “Hoa, dáng như tim vỡ,<br>Anh sợ tình ta cũng vỡ thôi!”',
    'Thuở ấy, nào tôi đã hiểu gì<br>Cánh hoa tan tác của sinh ly,<br>Cho nên cười đáp: “Màu hoa trắng<br>Là chút lòng trong chẳng biến suy”',
    'Người ấy thường hay ngắm lạnh lùng<br>Dải đường xa vút bóng chiều phong,<br>Và phương trời thẳm mờ sương, cát,<br>Tay vít dây hoa trắng chạnh lòng.',
  ];
  const S1 = 'Một mùa thu trước, mỗi hoàng hôn Nhặt cánh hoa rơi chẳng thấy buồn, Nhuộm ánh nắng tà qua mái tóc, Tôi chờ người đến với yêu đương.';
  const S2 = 'Người ấy thường hay vuốt tóc tôi, Thở dài trong lúc thấy tôi vui, Bảo rằng: “Hoa, dáng như tim vỡ, Anh sợ tình ta cũng vỡ thôi!”';
  const S3 = 'Thuở ấy, nào tôi đã hiểu gì Cánh hoa tan tác của sinh ly, Cho nên cười đáp: “Màu hoa trắng Là chút lòng trong chẳng biến suy”';
  const S4 = 'Người ấy thường hay ngắm lạnh lùng Dải đường xa vút bóng chiều phong, Và phương trời thẳm mờ sương, cát, Tay vít dây hoa trắng chạnh lòng.';

  const LESSONS = [
    {
      id: 'w-basic', group: 'lesson', title: 'Khởi động: định dạng cơ bản', source: 'Bài giảng MS Word – Phần 2',
      intro: 'Bôi đen văn bản rồi dùng nút trên ribbon hoặc phím tắt (Ctrl+B, Ctrl+I, Ctrl+U, Ctrl+E, Ctrl+J…). Danh sách bên trái tự đánh dấu ✓ khi bạn làm đúng.',
      html: html([
        'SOẠN THẢO VĂN BẢN CƠ BẢN',
        'Microsoft Word là phần mềm soạn thảo văn bản phổ biến nhất hiện nay. Word giúp chúng ta tạo, chỉnh sửa và trình bày văn bản một cách chuyên nghiệp.',
        'Các thao tác định dạng thường dùng gồm: chọn phông chữ, cỡ chữ, kiểu chữ đậm, nghiêng, gạch chân, màu chữ và căn lề đoạn văn.',
        'Tạo tài liệu mới: Ctrl + N', 'Lưu tài liệu: Ctrl + S', 'Mở tài liệu: Ctrl + O',
      ]),
      tasks: [
        { id: 'a', title: 'Tiêu đề căn giữa', hint: 'Đặt con trỏ vào tiêu đề → Ctrl+E (hoặc nút căn giữa).', check: (d) => align(d, 'SOẠN THẢO VĂN BẢN CƠ BẢN') === 'center' },
        { id: 'b', title: 'Tiêu đề in đậm, cỡ chữ từ 16 trở lên', hint: 'Bôi đen tiêu đề → Ctrl+B, chọn cỡ 16 ở ô cỡ chữ.', check: (d) => d.all('SOẠN THẢO VĂN BẢN CƠ BẢN', P.and(P.bold, P.sizeMin(16))) },
        { id: 'c', title: 'Tiêu đề có màu chữ (khác đen)', hint: 'Nút chữ A gạch đỏ (Font Color).', check: (d) => d.all('SOẠN THẢO VĂN BẢN CƠ BẢN', P.colored) },
        { id: 'd', title: 'Đoạn 1 căn đều hai bên', hint: 'Ctrl+J.', check: (d) => align(d, 'Microsoft Word là phần mềm') === 'justify' },
        { id: 'e', title: 'Đoạn 1 thụt lề dòng đầu', hint: 'Nút ⇥¶ (First line indent).', check: (d) => { const b = d.blockOf('Microsoft Word là phần mềm'); return !!(b && b.style.textIndent); } },
        { id: 'f', title: '"Microsoft Word" (đầu đoạn 1) in đậm và nghiêng', hint: 'Bôi đen 2 từ → Ctrl+B rồi Ctrl+I.', check: (d) => { const c = d.charsOf('Microsoft Word là'); return !!c && c.slice(0, 13).every((x) => x.st.b && x.st.i); } },
        { id: 'g', title: 'Đoạn 2: chữ "đậm" in đậm, "nghiêng" in nghiêng, "gạch chân" gạch dưới', hint: 'Định dạng đúng ý nghĩa của từng từ.',
          check: (d) => tailAll(d, 'chữ đậm', 3, P.bold) && d.charsOf('nghiêng, gạch chân').slice(0, 7).every((c) => c.st.i) && tailAll(d, 'nghiêng, gạch chân', 8, P.ul()) },
        { id: 'h', title: '3 dòng phím tắt thành danh sách Bullets', hint: 'Bôi đen 3 dòng → nút Bullets.', check: (d) => ['Ctrl + N', 'Ctrl + S', 'Ctrl + O'].every((t) => { const b = d.blockOf(t); return b && b.tagName === 'LI' && b.parentElement.tagName === 'UL'; }) },
        { id: 'i', title: 'Đoạn 2 giãn dòng 1.5', hint: 'Nút ↕ → 1.5 (hoặc Ctrl+5).', check: (d) => { const b = d.blockOf('Các thao tác định dạng'); return !!(b && b.style.lineHeight === '1.5'); } },
      ],
    },
    {
      id: 'w-poem', group: 'lesson', title: 'Bài 1.2 – Định dạng bài thơ "Hai sắc hoa Ti-gôn"', source: 'Bài tập thực hành MS Word – Bài 1.2',
      intro: 'Mỗi khổ thơ là một đoạn; các dòng trong khổ ngắt bằng Shift+Enter. Định dạng trước, rồi sao chép/di chuyển khổ thơ (Ctrl+C, Ctrl+X, Ctrl+V).',
      html: html(['HAI SẮC HOA TI-GÔN', ...POEM]),
      tasks: [
        { id: 'a', title: 'Tiêu đề: Tahoma, cỡ 26, đậm', hint: 'Bôi đen tiêu đề → chọn font Tahoma, cỡ 26, Ctrl+B.', check: (d) => d.all('HAI SẮC HOA TI-GÔN', P.and(P.font('Tahoma'), P.size(26), P.bold)) },
        { id: 'b', title: 'Khổ 1: Times New Roman, 18, đậm, gạch dưới một đường', hint: 'U (Single).', check: (d) => d.all(S1, P.and(P.font('Times New Roman'), P.size(18), P.bold, P.ul('single'))) },
        { id: 'c', title: 'Khổ 2: Arial, 18, nghiêng, gạch dưới hai đường', hint: 'Mũi tên ▾ cạnh nút U → Gạch đôi (Double).', check: (d) => d.all(S2, P.and(P.font('Arial'), P.size(18), P.italic, P.ul('double'))) },
        { id: 'd', title: 'Khổ 3: Calibri, 18, đậm nghiêng, gạch dưới từng từ', hint: 'U ▾ → Gạch từng từ (Words only).', check: (d) => d.all(S3, P.and(P.font('Calibri'), P.size(18), P.bold, P.italic, P.ul('words'))) },
        { id: 'e', title: 'Khổ 4: Verdana, 18, đậm nghiêng, Small Caps', hint: 'Nút Aa → Small Caps (Ctrl+Shift+K).', check: (d) => d.all(S4, P.and(P.font('Verdana'), P.size(18), P.bold, P.italic, P.caps('small'))) },
        { id: 'f', title: 'Sao chép khổ thơ đầu tiên vào cuối bài thơ', hint: 'Bôi đen khổ 1 → Ctrl+C → đặt con trỏ cuối bài → Enter → Ctrl+V.', check: (d) => d.count('Một mùa thu trước, mỗi hoàng hôn') >= 2 && d.indexOf('Một mùa thu trước, mỗi hoàng hôn', 1) > d.indexOf('Tay vít dây hoa trắng chạnh lòng') },
        { id: 'g', title: 'Di chuyển khổ thơ 4 lên trước khổ thơ 2', hint: 'Bôi đen khổ 4 → Ctrl+X → đặt con trỏ đầu khổ 2 → Ctrl+V.', check: (d) => ordered(d, 'Người ấy thường hay ngắm lạnh lùng', 'Người ấy thường hay vuốt tóc tôi') },
      ],
    },
    {
      id: 'w-font', group: 'lesson', title: 'Bài 1.3 – Hộp hội thoại Font', source: 'Bài tập thực hành MS Word – Bài 1.3',
      intro: 'Luyện các hiệu ứng của hộp thoại Font: gạch ngang, gạch đôi, chỉ số trên/dưới, chữ hoa nhỏ, chữ hoa lớn, bóng chữ.',
      html: html([
        'HỘP HỘI THOẠI FONT',
        'Nếu bạn muốn chọn những Font chữ khác nhau cho các ký tự, chỉ định in đậm Bold, in nghiêng Italic, gạch dưới Underline, chữ bóng Shadow, gạch một đường giữa các ký tự, chỉ số dưới, chỉ số trên thì ta có thể chọn lệnh trên tab Home. Tuy nhiên Word còn cung cấp rất nhiều kiểu định dạng khác cho các ký tự như: gạch hai đường giữa các ký tự, chữ hoa nhỏ… những định dạng này mặc nhiên không có nút lệnh trên tab Home. Do đó để sử dụng thì phải dùng hộp hội thoại Font.',
        'Strikethrough: gạch một đường giữa các ký tự.', 'Double strikethrough: gạch hai đường giữa các ký tự.', 'Superscript: chỉ số trên (a2+b2).', 'Subscript: chỉ số dưới (H2SO4).', 'Small Caps: chữ hoa nhỏ.', 'All Caps: chữ hoa lớn.',
      ]),
      tasks: [
        { id: 'a', title: 'Tiêu đề: đậm, căn giữa, cỡ từ 14', hint: 'Ctrl+B, Ctrl+E, chọn cỡ chữ.', check: (d) => d.all('HỘP HỘI THOẠI FONT', P.and(P.bold, P.sizeMin(14))) && align(d, 'HỘP HỘI THOẠI FONT') === 'center' },
        { id: 'b', title: 'Đoạn 1: chữ "Bold" in đậm, "Italic" in nghiêng, "Underline" gạch dưới', hint: 'Đặt con trỏ vào giữa từ rồi bấm Ctrl+B / Ctrl+I / Ctrl+U (không cần bôi đen).', check: (d) => tailAll(d, 'đậm Bold', 3, P.bold) && tailAll(d, 'nghiêng Italic', 7, P.italic) && tailAll(d, 'dưới Underline', 4, P.ul()) },
        { id: 'c', title: 'Chữ "Shadow" có hiệu ứng bóng', hint: 'Nút A có bóng (Text Effects → Shadow).', check: (d) => tailAll(d, 'bóng Shadow', 4, P.shadow) },
        { id: 'd', title: 'Trong đoạn 1: "gạch một đường giữa các ký tự" gạch ngang đơn', hint: 'Nút abc gạch ngang (Strikethrough).', check: (d) => d.all('gạch một đường giữa các ký tự', P.strike('single')) },
        { id: 'e', title: 'Trong đoạn 1: "gạch hai đường giữa các ký tự" gạch ngang đôi', hint: 'Nút abc gạch đôi (Double strikethrough).', check: (d) => d.all('gạch hai đường giữa các ký tự', P.strike('double')) },
        { id: 'f', title: 'a2+b2 → số 2 là chỉ số trên (a²+b²)', hint: 'Bôi đen từng số 2 → Ctrl+Shift+=.', check: (d) => { const c = d.charsOf('a2+b2'); return !!c && c[1].st.va === 'super' && c[4].st.va === 'super' && !c[0].st.va && !c[3].st.va; } },
        { id: 'g', title: 'H2SO4 → số 2 và 4 là chỉ số dưới (H₂SO₄)', hint: 'Bôi đen số → Ctrl+=.', check: (d) => { const c = d.charsOf('H2SO4'); return !!c && c[1].st.va === 'sub' && c[4].st.va === 'sub' && !c[2].st.va; } },
        { id: 'h', title: 'Dòng "Small Caps": cụm "chữ hoa nhỏ" dạng Small Caps', hint: 'Aa → Small Caps.', check: (d) => tailAll(d, 'Small Caps: chữ hoa nhỏ', 10, P.caps('small')) },
        { id: 'i', title: 'Dòng "All Caps": cụm "chữ hoa lớn" dạng All Caps', hint: 'Aa → ALL CAPS.', check: (d) => d.all('chữ hoa lớn', P.caps('all')) },
        { id: 'j', title: '6 nhãn đầu dòng (Strikethrough:, Double strikethrough:, …) in đậm', hint: 'Bôi đen từng nhãn → Ctrl+B.', check: (d) => ['Strikethrough: gạch một', 'Double strikethrough:', 'Superscript:', 'Subscript:', 'Small Caps:', 'All Caps:'].every((t) => { const c = d.charsOf(t); const n = t.split(':')[0].replace(/\s/g, '').length + 1; return !!c && c.slice(0, n).every((x) => x.st.b); }) },
      ],
    },
  ];

  const BADEN_1 = 'Núi Bà Đen nằm ở phía Đông Bắc thành phố Tây Ninh, tỉnh Tây Ninh, cách trung tâm thành phố 11km. Đây là ngọn núi cao nhất Nam Bộ (986m), nằm trong quần thể di tích lịch sử văn hóa thắng cảnh và du lịch núi Bà Đen đã được Thủ tướng Chính phủ công nhận là Khu du lịch quốc gia.';
  const BADEN_2 = 'Núi Bà Đen là một danh thắng nổi tiếng được coi là biểu tượng của tỉnh Tây Ninh. Từ núi Bà Đen nhìn xuống, sẽ thấy cả một vùng đồng bằng mênh mông bao phủ ngoại ô thành phố Tây Ninh. Nhìn từ xa, núi Bà Đen khi thì hiện rõ giữa nền trời xanh, lúc ẩn trong làn sương bảng lảng. Nơi đây còn thu hút du khách bởi một quần thể kiến trúc điện, chùa, miếu, tháp… đều mang đặc trưng của văn hóa Phật giáo và tín ngưỡng dân gian trải khắp từ chân núi, sườn núi, lên đến đỉnh. Trên núi Bà Đen có một số ngôi chùa như: chùa Phật, chùa Hang, chùa Hạ, chùa Trung, chùa Vân Sơn… Ngoài ra, tại đây còn có nhiều động nhỏ: động Thanh Long, động Ba Cô, động Ba Tuần…';

  const TECH = [
    ['Mô hình dữ liệu tin cậy', ': Việc gia tăng các hoạt động bất hợp pháp trên mạng cũng dẫn đến sự ra đời của hàng loạt mô hình dữ liệu tin cậy nhằm cung cấp các giải pháp bảo mật, bằng cách sử dụng bên thứ ba độc lập để đảm bảo dữ liệu được quản lý thích hợp và không bị độc quyền…'],
    ['Wifi 6 và 5g', ': Việc triển khai 5g trên toàn cầu vẫn tiếp tục mở rộng nhằm cải thiện băng thông rộng di động và cho phép người dùng tải xuống nhanh hơn.'],
    ['In 3d', ': là công nghệ bồi đắp vật liệu, đã trở thành giải pháp thay thế số một cho các quy trình sản xuất truyền thống.'],
    ['VR và AR', ': Tương tác với dữ liệu đã trở nên phổ biến, một phần nhờ vào thực tế mô phỏng. Với những cải tiến kỹ thuật số mới, chúng ta có thể tạo ra các mô phỏng và kịch bản thực tế mà không yêu cầu tài nguyên đắt tiền.'],
    ['Dịch vụ điện toán', ': Cho phép người dùng phân bổ không gian trên máy chủ thay vì ổ cứng cục bộ của máy tính…'],
  ];
  const TECH_K = 'Không giống như các phương pháp sản xuất cũ, in 3d tạo ra các vật thể ba chiều bằng cách đặt vật liệu từng lớp một để tạo ra sản phẩm đầu ra mong muốn.';
  // Tìm tiêu đề mục kèm dấu ":" để không nhầm với chữ trùng trong thân đoạn.
  const TECH_HEADS = TECH.map((t) => t[0]);
  const techHead = (d, h) => { const c = d.charsOf(h + ':'); return c ? c.slice(0, -1) : null; };

  const THU_1 = 'Mùa thu Hà Nội khiến lòng người xao xuyến bởi những con đường đầy lá vàng rơi, với tiết trời se lạnh và phảng phất chút mưa phùn. Thu về Hà Nội thay áo mới, lãng mạn hơn, mộng mơ hơn và cũng dịu dàng hơn.';
  const THU_2 = 'Mùa thu Hà Nội nhẹ nhàng, lãng mạn và trở thành nguồn cảm hứng sáng tác của biết bao nghệ sĩ. Thủ đô báo hiệu mùa thu bằng tiết trời mát mẻ, đường phố phủ đầy lá vàng và thoang thoảng đâu đây mùi thơm hoa sữa. Du lịch Hà Nội vào mùa thu, bạn mới cảm nhận được hết vẻ đẹp lãng mạn, mộng mơ của thành phố này…';
  const THU_PLACES = [['Chùa Trấn Quốc', 'Tây Hồ – Hà Nội'], ['Phố cổ Hà Nội', 'Hoàn Kiếm – Hà Nội'], ['Văn Miếu - Quốc Tử Giám', 'Đống Đa – Hà Nội'], ['Nhà thờ lớn Hà Nội', 'Hoàn Kiếm – Hà Nội'], ['Hoàng thành Thăng Long', 'Ba Đình – Hà Nội']];

  const CLOUD_1 = 'Các giải pháp công nghệ số hàng đầu hiện nay không thể không nhắc đến giải pháp điện toán đám mây. Đây đang là một trong những chủ đề mới mẻ được nhắc đến nhiều nhất tại các sự kiện công nghệ tại Việt Nam. Điện toán đám mây là các nguồn điện toán khổng lồ như phần mềm, dịch vụ… sẽ nằm tại các máy chủ ảo trên Internet thay vì trong máy tính gia đình và văn phòng để mọi người kết nối và sử dụng mỗi khi họ cần.';
  const CLOUD_2 = 'Với các dịch vụ sẵn có trên Internet, doanh nghiệp không phải mua và duy trì hàng trăm, thậm chí hàng nghìn máy tính cũng như phần mềm. Họ chỉ cần tập trung sản xuất bởi đã có người khác lo cơ sở hạ tầng và công nghệ thay họ. Bạn có thể truy cập đến bất kỳ tài nguyên nào tồn tại trong đám mây tại bất kỳ thời điểm nào và từ bất kỳ đâu thông qua hệ thống Internet.';
  const CLOUD_3 = 'Công nghệ số là một quá trình thay đổi từ phương thức thủ công truyền thống sang áp dụng công nghệ với các trụ cột là big data, internet kết nối vạn vật, điện toán đám mây,…';

  const PT_1 = 'Phan Thiết là thành phố biển nổi tiếng thuộc tỉnh Bình Thuận, cách thành phố Hồ Chí Minh khoảng 200Km về phía Đông và cách thành phố Nha Trang khoảng 240Km về phía Tây Nam. Nếu là “phượt thủ” kỳ cựu, bạn ắt hẳn biết rằng hành trình đi đến Phan Thiết sẽ đi qua một trong những cung đường ven biển đẹp nhất Việt Nam.';
  const PT_2 = 'Chẳng những hưởng trọn không khí biển trong lành và khung cảnh thiên nhiên ngọt ngào, bức tranh thôn quê hiện lên bên những khu vườn thanh long sai trĩu quả cũng là điểm nhấn thú vị.';
  const PT_LIST = ['Công viên nước Wonderland', 'Tháp Po Sah Inư', 'Bãi Rạng', 'Đồi cát bay'];

  const headerShaded = (t) => !!t && Array.from(t.rows).slice(0, 2).some((row) => Array.from(row.cells).every((c) => c.style.backgroundColor));

  const EXAMS = [
    {
      id: 'w-baden', group: 'exam', title: 'Du lịch núi Bà Đen', source: 'Đề thi thử Đợt 1 – Ca 1 – Đề CHẴN (Câu 1)',
      intro: 'Văn bản đã gõ sẵn. Trình bày giống mẫu đề: WordArt, Drop Cap, chia 2 cột, chèn ảnh, bảng lịch trình.',
      html: html(['DU LỊCH NÚI BÀ ĐEN', BADEN_1, BADEN_2,
        note('Dữ liệu bảng (tạo bảng 4 dòng × 2 cột, gộp 2 ô dòng đầu, gõ nội dung vào ô; xong có thể xóa đoạn ghi chú này):<br>LỊCH TRÌNH DU LỊCH NÚI TÂY NINH 2 NGÀY 1 ĐÊM | NGÀY – LỊCH TRÌNH | 1 – Khởi hành từ Sài Gòn đi Tây Ninh | 2 – Núi Bà Đen - Hồ Núi Đá - Cây thốt nốt trái tim - Sài Gòn')]),
      tasks: [
        { id: 'a', pts: 1, title: 'Tiêu đề "DU LỊCH NÚI BÀ ĐEN" dạng WordArt', hint: 'Đặt con trỏ vào tiêu đề → Insert → WordArt.', check: (d) => isWordArt(d, 'DU LỊCH NÚI BÀ ĐEN') },
        { id: 'b', pts: 0.25, title: 'Tiêu đề căn giữa', hint: 'Ctrl+E.', check: (d) => align(d, 'DU LỊCH NÚI BÀ ĐEN') === 'center' },
        { id: 'c', pts: 0.5, title: 'Drop Cap chữ "N" đầu đoạn 1, thả 3 dòng', hint: 'Con trỏ trong đoạn 1 → Insert → Drop Cap → 3 dòng.', check: (d) => dropLines(d, 'nằm ở phía Đông Bắc') === 3 },
        { id: 'd', pts: 0.25, title: 'Đoạn 1 in đậm', hint: 'Bôi đen đoạn 1 → Ctrl+B.', check: (d) => d.all('nằm ở phía Đông Bắc thành phố Tây Ninh', P.bold) },
        { id: 'e', pts: 1, title: 'Đoạn 1 và đoạn 2 chia 2 cột', hint: 'Bôi đen 2 đoạn → Layout → Columns → Two.', check: (d) => d.colsOf('nằm ở phía Đông Bắc').n === 2 && d.colsOf('là một danh thắng nổi tiếng').n === 2 },
        { id: 'f', pts: 0.25, title: 'Hai đoạn căn đều hai bên', hint: 'Ctrl+J.', check: (d) => align(d, 'nằm ở phía Đông Bắc') === 'justify' && align(d, 'là một danh thắng nổi tiếng') === 'justify' },
        { id: 'g', pts: 0.5, title: 'Chèn ảnh, kiểu bao chữ Square', hint: 'Insert → Picture → chọn ảnh → Picture Tools → Square.', check: squareImg },
        { id: 'h', pts: 0.5, title: 'Bảng có đủ 2 dòng lịch trình', hint: 'Insert → Table 4×2, gõ "Khởi hành từ Sài Gòn đi Tây Ninh" và "Núi Bà Đen - Hồ Núi Đá - Cây thốt nốt trái tim - Sài Gòn" vào ô.', check: (d) => !!d.tableWith('Khởi hành từ Sài Gòn đi Tây Ninh') && !!d.tableWith('Hồ Núi Đá - Cây thốt nốt trái tim') },
        { id: 'i', pts: 0.5, title: 'Dòng đầu bảng là ô gộp chứa "LỊCH TRÌNH DU LỊCH NÚI TÂY NINH 2 NGÀY 1 ĐÊM"', hint: 'Bôi đen 2 ô dòng 1 → Table Tools → Gộp ô.', check: (d) => { const td = d.cellWith('LỊCH TRÌNH DU LỊCH NÚI TÂY NINH'); return !!td && td.colSpan >= 2; } },
        { id: 'j', pts: 0.25, title: 'Tô nền dòng tiêu đề của bảng', hint: 'Table Tools → Shading.', check: (d) => headerShaded(d.tableWith('LỊCH TRÌNH DU LỊCH NÚI TÂY NINH') || d.tableWith('Khởi hành từ Sài Gòn')) },
      ],
    },
    {
      id: 'w-tech', group: 'exam', title: 'Xu hướng công nghệ 2023', source: 'Đề thi thử 26.12 – Ca 1 – Đề chẵn (Câu 1)',
      intro: 'Mỗi mục bắt đầu bằng ký hiệu ☞; tiêu đề mục có viền + tô nền, chữ đậm nghiêng gạch dưới. Phần SmartArt "Công việc IT" và hình mũi tên của đề gốc hãy luyện trên Word thật.',
      html: html(['XU HƯỚNG CÔNG NGHỆ 2023', ...TECH.slice(0, 3).map((t) => t[0] + t[1]), TECH_K, ...TECH.slice(3).map((t) => t[0] + t[1])]),
      tasks: [
        { id: 'a', pts: 1, title: 'Tiêu đề dạng WordArt', hint: 'Insert → WordArt.', check: (d) => isWordArt(d, 'XU HƯỚNG CÔNG NGHỆ 2023') },
        { id: 'b', pts: 1, title: '5 tiêu đề mục (Mô hình dữ liệu tin cậy, Wifi 6 và 5g, In 3d, VR và AR, Dịch vụ điện toán): đậm + nghiêng + gạch dưới', hint: 'Bôi đen từng tiêu đề mục → Ctrl+B, Ctrl+I, Ctrl+U.', check: (d) => TECH_HEADS.every((h) => { const c = techHead(d, h); return !!c && c.every((x) => x.st.b && x.st.i && x.st.ul); }) },
        { id: 'c', pts: 0.75, title: '5 tiêu đề mục có viền chữ và tô nền', hint: 'Home → ▢ Borders and Shading → Viền chữ, rồi chọn màu Tô nền chữ.', check: (d) => TECH_HEADS.every((h) => { const c = techHead(d, h); return !!c && c.every((x) => x.st.bx && x.st.hl); }) },
        { id: 'd', pts: 0.5, title: 'Chèn ký hiệu ☞ trước mỗi tiêu đề mục', hint: 'Đặt con trỏ đầu đoạn → Insert → Symbol → ☞.', check: (d) => TECH_HEADS.every((h) => { const i = d.find(h + ':'); if (!i) return false; const before = d.chars.slice(Math.max(0, i.start - 3), i.start).map((c) => c.ch).join(''); return before.includes('☞'); }) },
        { id: 'e', pts: 0.75, title: 'Các đoạn nội dung chia 2 cột có đường kẻ giữa', hint: 'Bôi đen từ mục đầu đến hết → Layout → Columns → Two + Line between.', check: (d) => { const c = d.colsOf('Mô hình dữ liệu tin cậy:'), c2 = d.colsOf('Dịch vụ điện toán:'); return c.n === 2 && c.line && c2.n === 2; } },
        { id: 'f', pts: 0.5, title: 'Drop Cap chữ "K" (đoạn "Không giống như…"), 2 dòng trở lên', hint: 'Insert → Drop Cap.', check: (d) => dropLines(d, 'hông giống như các phương pháp') >= 2 },
        { id: 'g', pts: 0.5, title: 'Chèn ảnh, kiểu bao chữ Square', hint: 'Insert → Picture → Square.', check: squareImg },
      ],
    },
    {
      id: 'w-thu', group: 'exam', title: 'Mùa thu Hà Nội', source: 'Đề thi thử Đợt 1 – Ca 1 – Đề LẺ (Câu 1)',
      intro: 'Trình bày 2 đoạn văn và bảng 5 địa danh theo mẫu đề.',
      html: html(['MÙA THU HÀ NỘI', THU_1, THU_2, note('Dữ liệu bảng (6 dòng × 3 cột: STT | Địa danh | Địa chỉ): ' + THU_PLACES.map((p, i) => (i + 1) + '. ' + p[0] + ' – ' + p[1]).join('; '))]),
      tasks: [
        { id: 'a', pts: 0.5, title: 'Tiêu đề dạng WordArt', hint: 'Insert → WordArt.', check: (d) => isWordArt(d, 'MÙA THU HÀ NỘI') },
        { id: 'b', pts: 0.5, title: 'Drop Cap chữ "M" đoạn 1, 3 dòng', hint: 'Insert → Drop Cap → 3 dòng.', check: (d) => dropLines(d, 'ùa thu Hà Nội khiến lòng người') === 3 },
        { id: 'c', pts: 1, title: 'Hai đoạn chia 2 cột', hint: 'Layout → Columns → Two.', check: (d) => d.colsOf('khiến lòng người xao xuyến').n === 2 && d.colsOf('nhẹ nhàng, lãng mạn').n === 2 },
        { id: 'd', pts: 0.5, title: 'Chèn ảnh kiểu Square', hint: 'Insert → Picture → Square.', check: squareImg },
        { id: 'e', pts: 0.5, title: 'Font Times New Roman cỡ 14 cho cả 2 đoạn', hint: 'Bôi đen 2 đoạn → chọn font và cỡ.', check: (d) => d.all('khiến lòng người xao xuyến bởi những con đường', P.and(P.font('Times New Roman'), P.size(14))) && d.all('nhẹ nhàng, lãng mạn và trở thành nguồn cảm hứng', P.and(P.font('Times New Roman'), P.size(14))) },
        { id: 'f', pts: 1, title: 'Bảng chứa đủ 5 địa danh', hint: 'Insert → Table 6×3, gõ STT, Địa danh, Địa chỉ.', check: (d) => THU_PLACES.every((p) => !!d.tableWith(p[0])) },
        { id: 'g', pts: 0.5, title: 'Dòng tiêu đề bảng: in đậm + tô nền', hint: 'Bôi đen dòng 1 → Ctrl+B → Table Tools → Shading.', check: (d) => { const t = d.tableWith('Chùa Trấn Quốc'); if (!t) return false; const row = t.rows[0]; const doc = new global.WordDoc(row, d.def); return Array.from(row.cells).every((c) => c.style.backgroundColor) && doc.chars.filter((c) => /\S/.test(c.ch)).every((c) => c.st.b); } },
        { id: 'h', pts: 0.5, title: 'Đoạn 2 in nghiêng, căn đều', hint: 'Ctrl+I, Ctrl+J.', check: (d) => d.all('nhẹ nhàng, lãng mạn và trở thành', P.italic) && align(d, 'nhẹ nhàng, lãng mạn') === 'justify' },
      ],
    },
    {
      id: 'w-cloud', group: 'exam', title: 'Điện toán đám mây', source: 'Đề thi thử Đợt 1 – Ca 2 – Đề LẺ (Câu 1)',
      intro: 'Trình bày bài viết về điện toán đám mây: WordArt, Drop Cap, chia cột có đường kẻ, ảnh, khung viền đoạn cuối.',
      html: html(['ĐIỆN TOÁN ĐÁM MÂY', CLOUD_1, CLOUD_2, CLOUD_3]),
      tasks: [
        { id: 'a', pts: 0.5, title: 'Tiêu đề dạng WordArt', hint: 'Insert → WordArt.', check: (d) => isWordArt(d, 'ĐIỆN TOÁN ĐÁM MÂY') },
        { id: 'b', pts: 0.5, title: 'Drop Cap chữ "C" đoạn 1, 3 dòng', hint: 'Insert → Drop Cap.', check: (d) => dropLines(d, 'ác giải pháp công nghệ số hàng đầu') === 3 },
        { id: 'c', pts: 1, title: 'Đoạn 1 và 2 chia 2 cột có đường kẻ giữa', hint: 'Layout → Columns → Two + Line between.', check: (d) => { const a = d.colsOf('giải pháp công nghệ số hàng đầu'), b = d.colsOf('Với các dịch vụ sẵn có'); return a.n === 2 && a.line && b.n === 2; } },
        { id: 'd', pts: 1, title: 'Chèn ảnh kiểu Square', hint: 'Insert → Picture → Square.', check: squareImg },
        { id: 'e', pts: 1, title: 'Đoạn 2 căn đều, giãn dòng 1.5', hint: 'Ctrl+J, Ctrl+5.', check: (d) => align(d, 'Với các dịch vụ sẵn có') === 'justify' && (d.blockStyle('Với các dịch vụ sẵn có') || {}).lineHeight === '1.5' },
        { id: 'f', pts: 1, title: 'Đoạn cuối "Công nghệ số là…" có viền đoạn và tô nền đoạn', hint: 'Borders and Shading → Viền đoạn + Tô nền đoạn.', check: (d) => { const s = d.blockStyle('Công nghệ số là một quá trình'); return !!s && s.borderStyle === 'solid' && !!s.backgroundColor; } },
      ],
    },
    {
      id: 'w-phanthiet', group: 'exam', title: 'Du lịch Phan Thiết', source: 'BT ôn tập cuối kỳ – Word',
      intro: 'Trình bày bài giới thiệu Phan Thiết và danh sách điểm đến dạng bullets ký hiệu.',
      html: html(['DU LỊCH PHAN THIẾT', PT_1, PT_2, 'Các điểm đến nổi bật:', ...PT_LIST]),
      tasks: [
        { id: 'a', pts: 1, title: 'Tiêu đề dạng WordArt', hint: 'Insert → WordArt.', check: (d) => isWordArt(d, 'DU LỊCH PHAN THIẾT') },
        { id: 'b', pts: 0.5, title: 'Drop Cap chữ "P" đoạn 1', hint: 'Insert → Drop Cap.', check: (d) => dropLines(d, 'han Thiết là thành phố biển') >= 2 },
        { id: 'c', pts: 1, title: 'Đoạn 1 và 2 chia 2 cột', hint: 'Layout → Columns → Two.', check: (d) => d.colsOf('là thành phố biển nổi tiếng').n === 2 && d.colsOf('Chẳng những hưởng trọn').n === 2 },
        { id: 'd', pts: 1, title: '4 điểm đến thành bullets ký hiệu (✈, ☞, ❖… không dùng chấm tròn)', hint: 'Bôi đen 4 dòng → Bullets ▾ → chọn ký hiệu.', check: (d) => PT_LIST.every((t) => { const b = d.blockOf(t); const ul = b && b.closest('ul'); return !!ul && /^bl-/.test(ul.className) && ul.className !== 'bl-disc'; }) },
        { id: 'e', pts: 0.5, title: '“phượt thủ” in đậm nghiêng', hint: 'Ctrl+B, Ctrl+I.', check: (d) => d.all('phượt thủ', P.and(P.bold, P.italic)) },
        { id: 'f', pts: 0.5, title: '"Các điểm đến nổi bật:" đậm, gạch dưới đôi', hint: 'U ▾ → Gạch đôi.', check: (d) => d.all('Các điểm đến nổi bật:', P.and(P.bold, P.ul('double'))) },
        { id: 'g', pts: 0.5, title: 'Chèn ảnh kiểu Square', hint: 'Insert → Picture → Square.', check: squareImg },
      ],
    },
  ];

  global.WORD_DATA = { LESSONS, EXAMS, all: [...LESSONS, ...EXAMS] };
})(window);
