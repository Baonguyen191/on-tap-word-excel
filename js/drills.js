/* Luyện phím tắt (bấm tổ hợp phím thật) + trắc nghiệm vị trí lệnh / kiến thức hay ra trong đề. */
(function (global) {
  'use strict';

  // Chỉ dùng tổ hợp mà trình duyệt cho phép chặn (không có Ctrl+N/T/W).
  const SHORTCUTS = [
    { app: 'Word', act: 'In đậm', keys: 'Ctrl+B' },
    { app: 'Word', act: 'In nghiêng', keys: 'Ctrl+I' },
    { app: 'Word', act: 'Gạch dưới', keys: 'Ctrl+U' },
    { app: 'Word', act: 'Căn giữa đoạn', keys: 'Ctrl+E' },
    { app: 'Word', act: 'Căn trái đoạn', keys: 'Ctrl+L' },
    { app: 'Word', act: 'Căn phải đoạn', keys: 'Ctrl+R' },
    { app: 'Word', act: 'Căn đều hai bên', keys: 'Ctrl+J' },
    { app: 'Word', act: 'Chỉ số dưới (H₂O)', keys: 'Ctrl+=' },
    { app: 'Word', act: 'Chỉ số trên (x²)', keys: 'Ctrl+Shift+=' },
    { app: 'Word', act: 'Tăng cỡ chữ 1pt', keys: 'Ctrl+]' },
    { app: 'Word', act: 'Giảm cỡ chữ 1pt', keys: 'Ctrl+[' },
    { app: 'Word', act: 'Giãn dòng 1.5', keys: 'Ctrl+5' },
    { app: 'Word', act: 'Giãn dòng đôi', keys: 'Ctrl+2' },
    { app: 'Word', act: 'Giãn dòng đơn', keys: 'Ctrl+1' },
    { app: 'Word', act: 'Mở hộp thoại Font', keys: 'Ctrl+D' },
    { app: 'Word', act: 'Chữ hoa nhỏ (Small Caps)', keys: 'Ctrl+Shift+K' },
    { app: 'Word', act: 'Đổi chữ hoa/thường', keys: 'Shift+F3' },
    { app: 'Word', act: 'Xuống dòng nhưng KHÔNG tạo đoạn mới', keys: 'Shift+Enter' },
    { app: 'Word', act: 'Ngắt trang (Page Break)', keys: 'Ctrl+Enter' },
    { app: 'Word', act: 'Tìm kiếm', keys: 'Ctrl+F' },
    { app: 'Word', act: 'Tìm và thay thế', keys: 'Ctrl+H' },
    { app: 'Word', act: 'Lưu tài liệu', keys: 'Ctrl+S' },
    { app: 'Word', act: 'Mở tài liệu', keys: 'Ctrl+O' },
    { app: 'Word', act: 'Chọn toàn bộ văn bản', keys: 'Ctrl+A' },
    { app: 'Word', act: 'Sao chép', keys: 'Ctrl+C' },
    { app: 'Word', act: 'Cắt', keys: 'Ctrl+X' },
    { app: 'Word', act: 'Dán', keys: 'Ctrl+V' },
    { app: 'Word', act: 'Hoàn tác', keys: 'Ctrl+Z' },
    { app: 'Word', act: 'Làm lại', keys: 'Ctrl+Y' },
    { app: 'Word', act: 'In tài liệu', keys: 'Ctrl+P' },
    { app: 'Excel', act: 'Điền công thức xuống (Fill Down)', keys: 'Ctrl+D' },
    { app: 'Excel', act: 'Điền công thức sang phải (Fill Right)', keys: 'Ctrl+R' },
    { app: 'Excel', act: 'Khóa tham chiếu $ (tuyệt đối)', keys: 'F4' },
    { app: 'Excel', act: 'Sửa nội dung ô hiện tại', keys: 'F2' },
    { app: 'Excel', act: 'AutoSum – tính tổng nhanh', keys: 'Alt+=' },
    { app: 'Excel', act: 'Nhập ngày hiện tại', keys: 'Ctrl+;' },
    { app: 'Excel', act: 'Mở Format Cells', keys: 'Ctrl+1' },
    { app: 'Excel', act: 'Bật/tắt AutoFilter', keys: 'Ctrl+Shift+L' },
    { app: 'Excel', act: 'Nhập cùng công thức cho cả vùng chọn', keys: 'Ctrl+Enter' },
    { app: 'Excel', act: 'Xuống dòng trong cùng 1 ô', keys: 'Alt+Enter' },
    { app: 'Excel', act: 'Về ô A1', keys: 'Ctrl+Home' },
    { app: 'Excel', act: 'Chọn cả cột', keys: 'Ctrl+Space' },
    { app: 'Excel', act: 'Chọn cả hàng', keys: 'Shift+Space' },
    { app: 'Excel', act: 'Mở hộp thoại chèn hàm (Insert Function)', keys: 'Shift+F3' },
    { app: 'Excel', act: 'Ẩn hàng đang chọn', keys: 'Ctrl+9' },
    { app: 'Excel', act: 'Định dạng phần trăm', keys: 'Ctrl+Shift+5' },
  ];

  const CODE_KEY = { Equal: '=', Minus: '-', BracketLeft: '[', BracketRight: ']', Semicolon: ';', Quote: "'", Comma: ',', Period: '.', Slash: '/', Backslash: '\\', Backquote: '`', Space: 'Space', Enter: 'Enter', NumpadEnter: 'Enter', Home: 'Home', End: 'End', Tab: 'Tab', Escape: 'Esc', Delete: 'Delete', Backspace: 'Backspace' };
  function comboOf(e) {
    let key = null;
    if (/^Key[A-Z]$/.test(e.code)) key = e.code.slice(3);
    else if (/^Digit\d$/.test(e.code)) key = e.code.slice(5);
    else if (/^F\d{1,2}$/.test(e.code)) key = e.code;
    else if (CODE_KEY[e.code]) key = CODE_KEY[e.code];
    if (!key) return null;
    const parts = [];
    if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
    if (e.shiftKey) parts.push('Shift');
    if (e.altKey) parts.push('Alt');
    parts.push(key);
    return parts.join('+');
  }

  const QUIZ = [
    { app: 'Word', q: 'Tạo chữ cái lớn đầu đoạn (Drop Cap) nằm ở đâu?', a: 'Insert → Text → Drop Cap', o: ['Home → Font → Drop Cap', 'Page Layout → Columns', 'View → Drop Cap'] },
    { app: 'Word', q: 'Chia đoạn văn thành 2 cột dùng lệnh nào?', a: 'Page Layout (Layout) → Columns', o: ['Insert → Table', 'Home → Paragraph → Columns', 'View → Split'] },
    { app: 'Word', q: 'Tạo dấu chìm (Watermark) cho trang theo bài giảng?', a: 'Page Layout → Watermark', o: ['Insert → Picture', 'Review → Watermark', 'Home → Styles'] },
    { app: 'Word', q: 'Đổi màu nền trang văn bản?', a: 'Page Layout → Page Color', o: ['Home → Shading', 'Insert → Shapes', 'View → Background'] },
    { app: 'Word', q: 'Tạo viền cho cả trang giấy?', a: 'Page Layout → Page Borders', o: ['Home → Borders → Outside Borders', 'Insert → Text Box', 'Review → Track Changes'] },
    { app: 'Word', q: 'Kiểm tra chính tả và ngữ pháp?', a: 'Review → Spelling & Grammar', o: ['Home → Editing', 'References → Citations', 'View → Proofing'] },
    { app: 'Word', q: 'Trộn thư (Mail Merge) bắt đầu ở tab nào?', a: 'Mailings → Start Mail Merge', o: ['Insert → Quick Parts', 'References → Mail', 'Review → Compare'] },
    { app: 'Word', q: 'Đặt điểm dừng Tab chính xác (vị trí, kiểu, Leader)?', a: 'Home → Paragraph (mũi tên góc) → Tabs…', o: ['Insert → Tab', 'View → Ruler → Tab', 'Page Layout → Margins'] },
    { app: 'Word', q: 'Chèn chữ nghệ thuật WordArt?', a: 'Insert → Text → WordArt', o: ['Home → Text Effects', 'Design → Themes', 'Insert → SmartArt'] },
    { app: 'Word', q: 'Chèn sơ đồ dạng khối như "Công việc IT tuyển dụng nhiều nhất"?', a: 'Insert → SmartArt', o: ['Insert → Chart', 'Insert → WordArt', 'Home → Bullets'] },
    { app: 'Word', q: 'Ảnh nằm giữa đoạn, chữ bao quanh hình chữ nhật – chọn kiểu Wrap Text nào?', a: 'Square', o: ['In Line with Text', 'Behind Text', 'Top and Bottom'] },
    { app: 'Word', q: 'Gạch hai đường giữa chữ (Double strikethrough) chỉ có trong…', a: 'Hộp thoại Font (Ctrl+D)', o: ['Nút abc trên tab Home', 'Tab Review', 'Hộp thoại Paragraph'] },
    { app: 'Word', q: 'Shift+Enter khác Enter ở điểm nào?', a: 'Xuống dòng nhưng vẫn cùng một đoạn', o: ['Tạo trang mới', 'Tạo đoạn mới có khoảng cách', 'Không khác gì'] },
    { app: 'Word', q: 'Gộp các ô đã chọn trong bảng?', a: 'Table Tools Layout → Merge Cells', o: ['Home → Merge & Center', 'Insert → Table → Merge', 'Design → Borders'] },
    { app: 'Word', q: 'Đánh số trang tự động?', a: 'Insert → Page Number', o: ['Home → Numbering', 'References → Footnote', 'Page Layout → Line Numbers'] },
    { app: 'Excel', q: 'Sắp xếp theo nhiều cột (Tên hàng tăng, Thành tiền giảm)?', a: 'Data → Sort (Add Level)', o: ['Home → Sort A→Z hai lần', 'Data → Filter', 'Formulas → Sort'] },
    { app: 'Excel', q: 'Lọc ra bảng mới đặt dưới bảng tính theo điều kiện "A hoặc B"?', a: 'Data → Advanced (Copy to another location)', o: ['Data → Filter', 'Home → Find & Select', 'Data → Remove Duplicates'] },
    { app: 'Excel', q: 'Trong vùng điều kiện của Advanced Filter, hai điều kiện "HOẶC" được đặt…', a: 'Trên hai dòng khác nhau', o: ['Trên cùng một dòng', 'Trong cùng một ô, cách bởi dấu phẩy', 'Không đặt được'] },
    { app: 'Excel', q: 'Vẽ biểu đồ so sánh?', a: 'Insert → Charts', o: ['Data → Chart', 'Formulas → Chart', 'View → Chart'] },
    { app: 'Excel', q: 'Đổi chiều biểu diễn của biểu đồ (hàng ↔ cột)?', a: 'Chart Design → Switch Row/Column', o: ['Format → Rotate', 'Data → Transpose', 'Layout → Axis'] },
    { app: 'Excel', q: 'Cố định dòng tiêu đề khi cuộn (Freeze Panes)?', a: 'View → Freeze Panes', o: ['Page Layout → Print Titles', 'Home → Format → Lock', 'Review → Protect Sheet'] },
    { app: 'Excel', q: 'Ẩn đường lưới (Gridlines) của sheet?', a: 'View → bỏ chọn Gridlines', o: ['Home → Borders → No Border', 'File → Options → Print', 'Insert → Table'] },
    { app: 'Excel', q: 'Tạo liên kết bấm vào để chuyển sang sheet khác?', a: 'Chuột phải → Hyperlink → Place in This Document (Bookmark)', o: ['Data → Connections', 'Formulas → Name Manager', 'Insert → Object'] },
    { app: 'Excel', q: 'Hàm dò tìm khi bảng tra nằm NGANG (mã ở dòng đầu)?', a: 'HLOOKUP', o: ['VLOOKUP', 'MATCH', 'SUMIF'] },
    { app: 'Excel', q: '=VLOOKUP(A2,$E$2:$F$5,2) thiếu đối số thứ 4 thì Excel…', a: 'Dò gần đúng (dễ ra sai nếu bảng chưa sắp xếp)', o: ['Báo lỗi #VALUE!', 'Dò chính xác', 'Trả về cột 1'] },
    { app: 'Excel', q: 'Công thức =B4*$F$1 ở C4, kéo xuống C5 sẽ thành…', a: '=B5*$F$1', o: ['=B5*$F$2', '=B4*$F$1', '=C5*$F$1'] },
    { app: 'Excel', q: 'Công thức =SUMIF($C$4:$C$13,$F17,G$16) ở G17 copy sang H18 thành…', a: '=SUMIF($C$4:$C$13,$F18,H$16)', o: ['=SUMIF($C$4:$C$13,$G18,H$17)', '=SUMIF($C$5:$C$14,$F18,H$16)', '=SUMIF($C$4:$C$13,$F17,G$16)'] },
    { app: 'Excel', q: 'Kết quả =MID("TL3ĐL",4,2)?', a: 'ĐL', o: ['3Đ', 'L3', 'TL'] },
    { app: 'Excel', q: '=RIGHT("1IN210",3) trả về gì?', a: 'Chuỗi "210" (cần VALUE để thành số)', o: ['Số 210', 'Số 1', 'Lỗi #VALUE!'] },
    { app: 'Excel', q: 'So sánh =LEFT(B4,1)=1 với B4="1IN210" cho kết quả…', a: 'FALSE vì "1" (chuỗi) khác 1 (số)', o: ['TRUE', 'Lỗi #N/A', 'Lỗi #NAME?'] },
    { app: 'Excel', q: 'Đếm số xe hãng MITSUBISHI trong cột tên xe kiểu "MITSUBISHI JOLIE"?', a: '=COUNTIF(B4:B12,"MITSUBISHI*")', o: ['=COUNT(B4:B12,"MITSUBISHI")', '=COUNTIF(B4:B12,MITSUBISHI)', '=SUMIF(B4:B12,"MITSUBISHI")'] },
    { app: 'Excel', q: 'Lỗi #N/A khi dùng VLOOKUP thường do…', a: 'Không tìm thấy giá trị dò trong cột đầu của bảng', o: ['Chia cho 0', 'Sai tên hàm', 'Cột trả về quá lớn'] },
    { app: 'Excel', q: 'Lỗi #NAME? xuất hiện khi…', a: 'Gõ sai tên hàm hoặc quên dấu nháy kép quanh chữ', o: ['Ô tham chiếu bị xóa', 'Chia cho 0', 'Kết quả quá lớn'] },
    { app: 'Excel', q: 'Ngày bán sau 20/12/2024 nên viết điều kiện thế nào?', a: 'E4>DATE(2024,12,20)', o: ['E4>20/12/2024', 'E4>"20/12/2024"', 'DAY(E4)>20'] },
    { app: 'Excel', q: 'Xếp hạng điểm trong G4:G13 (cao nhất hạng 1)?', a: '=RANK(G4,$G$4:$G$13)', o: ['=RANK(G4,G4:G13)', '=MAX(G4:G13)', '=RANK($G$4,G4:G13,1)'] },
  ];

  const shuffle = (a) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };

  global.DRILLS = { SHORTCUTS, QUIZ, comboOf, shuffle };
})(window);
