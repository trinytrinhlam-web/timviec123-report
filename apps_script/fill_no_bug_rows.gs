/**
 * SCRIPT CHẠY 1 LẦN (không phải onEdit) — KHÔNG dán vào file onEdit_NgayTest.gs.
 * -------------------------------------------------------------------------
 * Bổ sung thông tin cho các dòng 05_Lỗi_Tester (243-298): đây là danh sách
 * Trang/Tính năng Tester đã kiểm tra lại và CHƯA phát sinh lỗi (không phải
 * dòng khai báo bug). Dòng 243 (DS_08) là mẫu Tester đã viết tay đầy đủ;
 * dòng 244 (DS_19) gần như đầy đủ; các dòng 245-298 mới chỉ có DS_ID/Tên
 * trang (245-254) hoặc YC_ID (255-298).
 *
 * Script này:
 *  1. Với 245-254 (theo Trang): điền YC_ID liên quan (D) và Mô tả yêu cầu (F)
 *     — YC_ID lấy từ cột DS_ID ở 03_Tính_năng khớp với DS_ID của dòng; Mô tả
 *     là câu hướng dẫn truy cập trang, tổng hợp từ 02_UI_UX/03_Tính_năng.
 *  2. Với 255-298 (theo Tính năng/YC_ID):
 *     - Nếu YC_ID CHƯA từng xuất hiện ở phần khai báo bug phía trên (dòng
 *       5-242): điền đầy đủ DS_ID (A), Tên trang (B), Tính năng (C) và Mô tả
 *       yêu cầu (F) — tra thẳng từ 03_Tính_năng theo YC_ID.
 *     - Nếu YC_ID đã TRÙNG với một dòng bug có sẵn ở 5-242 (cùng YC_ID đã có
 *       BUG_ID riêng) — theo yêu cầu Tester, KHÔNG điền lại A/B/C/F/H/J/Q/R/S
 *       (tránh lặp/mâu thuẫn với dòng bug gốc), CHỈ ghi 1 dòng chú thích vào
 *       Ghi chú Tester (T) nêu rõ trùng với DS_ID/BUG_ID nào, ở dòng nào.
 *       Danh sách trùng: YC_32, YC_33, YC_34, YC_46, YC_69, YC_70, YC_74,
 *       YC_82, YC_19, YC_47, YC_48 (xem TINHNANG_TRUNG_LAP bên dưới).
 *  3. Với các dòng KHÔNG trùng (245-254 và phần còn lại của 255-298): điền
 *     các cột trạng thái "không lỗi" (H, J, Q, R, S) theo đúng khuôn mẫu
 *     Tester đã dùng ở dòng 243/244 (Tiêu đề lỗi = "Không lỗi", Kết quả
 *     mong đợi = "Đã pass", Trạng thái kiểm tra nguồn = "Hoàn thành",
 *     Tester = "Mr. Ẩn"). Ngày phát hiện (Q) = 31/08/2026 theo yêu cầu Tester
 *     (khớp ngày ở dòng mẫu 243).
 *  4. Với 243-298: điền công thức 2 cột U (Trạng thái lỗi hiện tại) và V
 *     (Đã gửi RTM?) — ĐÚNG công thức Tester đã dùng sẵn ở các dòng phía trên
 *     (240-243), chỉ neo lại theo số dòng tương ứng. Đây là formula thường
 *     (không phải công thức mảng spill của 06_RTM_Dev_Test) nên an toàn khi
 *     ghi từng ô bằng script.
 *
 * Script CHỈ ghi vào Ô ĐANG TRỐNG — nếu Tester đã tự nhập giá trị nào rồi
 * thì giữ nguyên, không ghi đè. Vì vậy an toàn khi chạy nhiều lần.
 *
 * CÁCH CHẠY:
 *  1. Vào Tiện ích mở rộng → Apps Script (cùng project với onEdit_NgayTest.gs).
 *  2. Tạo FILE MỚI tên tuỳ ý (vd: fill_no_bug_rows.gs), dán toàn bộ nội dung này vào.
 *  3. Trong thanh công cụ, chọn hàm "fillNoBugRows" → bấm Run (▶).
 *  4. Cấp quyền nếu được hỏi lần đầu.
 *  5. Kiểm tra lại 05_Lỗi_Tester dòng 243-298 cho đúng thực tế.
 */

var SHEET_LOI = '05_Lỗi_Tester';
var HANG_DAU_KHONG_LOI = 243;
var HANG_CUOI_KHONG_LOI = 298;
var NGAY_PHAT_HIEN = new Date(2026, 7, 31); // 31/08/2026, theo yêu cầu Tester

var TRANG_KHONG_LOI = [
  { row: 245, dsId: 'DS_31', tenTrang: 'Trang Xác thực tài khoản', ycId: 'YC_06', moTa: 'Truy cập vào trang Xác thực tài khoản để xác minh tài khoản qua mã hoặc liên kết email/OTP Zalo.' },
  { row: 246, dsId: 'DS_49', tenTrang: 'Trang Trung tâm xác thực', ycId: '', moTa: 'Truy cập vào Trung tâm xác thực (Nhà tuyển dụng) để xác thực thông tin doanh nghiệp.' },
  { row: 247, dsId: 'DS_54', tenTrang: 'Trang Gói dịch vụ', ycId: 'YC_77, YC_78', moTa: 'Truy cập vào trang Gói dịch vụ để xem, so sánh, mua gói dịch vụ tuyển dụng và áp mã giảm giá/khuyến mãi.' },
  { row: 248, dsId: 'DS_55', tenTrang: 'Trang Đơn hàng của tôi', ycId: '', moTa: 'Truy cập vào trang Đơn hàng của tôi (Nhà tuyển dụng) để xem lịch sử đơn hàng đã mua.' },
  { row: 249, dsId: 'DS_56', tenTrang: 'Popup nạp tiền', ycId: '', moTa: 'Mở popup Nạp tiền (Nhà tuyển dụng) để nạp tiền vào tài khoản.' },
  { row: 250, dsId: 'DS_58', tenTrang: 'Vị trí ứng tuyển- Chuẩn bị phỏng vấn', ycId: 'YC_73', moTa: 'Truy cập vào mục Chuẩn bị phỏng vấn (/employer/interviews) để đặt lịch phỏng vấn online/offline với ứng viên.' },
  { row: 251, dsId: 'DS_60', tenTrang: 'Vị trí ứng tuyển- Ứng viên quan tâm', ycId: 'YC_72', moTa: 'Truy cập vào mục Ứng viên quan tâm (/employer/analytics/saved-candidates) để xem danh sách ứng viên quan tâm tới vị trí tuyển dụng.' },
  { row: 252, dsId: 'DS_61', tenTrang: 'Hồ sơ đã nộp theo tin truyển dụng', ycId: '', moTa: 'Truy cập vào mục Hồ sơ đã nộp theo tin tuyển dụng (/employer/postings/[id]/applicants) để xem danh sách hồ sơ ứng viên đã nộp.' },
  { row: 253, dsId: 'DS_63', tenTrang: 'Xem hồ sơ ứng viên', ycId: '', moTa: 'Truy cập vào trang Xem hồ sơ ứng viên (/employer/applications/[id]) để xem chi tiết hồ sơ ứng tuyển.' },
  { row: 254, dsId: 'DS_64', tenTrang: 'Lịch phỏng vấn của tôi', ycId: 'YC_38', moTa: 'Truy cập vào trang Lịch phỏng vấn của tôi (Ứng viên) để theo dõi lịch phỏng vấn, trao đổi và kết quả ứng tuyển.' },
];

var TINHNANG_KHONG_LOI = [
  { row: 255, ycId: 'YC_02', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Đăng nhập bằng mạng xã hội', moTa: 'Đăng nhập nhanh qua tài khoản Google/Facebook.' },
  { row: 264, ycId: 'YC_03', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Đăng nhập bằng Zalo qua OTP', moTa: 'Nhập SĐT Zalo, hệ thống gửi và xác minh mã OTP.' },
  { row: 265, ycId: 'YC_05', dsId: 'DS_22', tenTrang: 'Trang Đăng ký', tenTinhNang: 'Đăng ký bằng Zalo', moTa: 'Tạo tài khoản bằng SĐT Zalo và xác minh OTP.' },
  { row: 266, ycId: 'YC_06', dsId: 'DS_31', tenTrang: 'Trang Xác thực tài khoản', tenTinhNang: 'Xác minh email / Zalo', moTa: 'Xác minh tài khoản qua mã hoặc liên kết email hay OTP Zalo.' },
  { row: 267, ycId: 'YC_07', dsId: 'DS_22', tenTrang: 'Trang Đăng ký', tenTinhNang: 'Thông báo đăng ký thành công', moTa: 'Xác nhận đăng ký thành công và điều hướng sang khảo sát.' },
  { row: 268, ycId: 'YC_09', dsId: 'DS_25', tenTrang: 'Trang Quên mật khẩu', tenTinhNang: 'Xác minh khi quên mật khẩu', moTa: 'Nhập mã xác minh trước khi đặt lại mật khẩu.' },
  { row: 269, ycId: 'YC_10', dsId: 'DS_18', tenTrang: 'Trang Đổi mật khẩu', tenTinhNang: 'Đổi mật khẩu', moTa: 'Nhập mật khẩu hiện tại và mật khẩu mới; thông báo thành công.' },
  { row: 270, ycId: 'YC_11', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Phân luồng Ứng viên và Nhà tuyển dụng', moTa: 'Điều hướng đúng luồng UV và NTD sau đăng nhập.' },
  { row: 271, ycId: 'YC_12', dsId: 'DS_23', tenTrang: 'Trang Đăng ký - Ứng viên', tenTinhNang: 'Khảo sát ứng viên khi onboarding', moTa: 'Thu thập công việc, kinh nghiệm và khu vực để cá nhân hoá gợi ý.' },
  { row: 272, ycId: 'YC_14', dsId: 'DS_23, DS_24', tenTrang: 'Trang Đăng ký - Ứng viên, Trang Đăng ký - Nhà tuyển dụng', tenTinhNang: 'Điều hướng bước và lưu khảo sát', moTa: 'Cho phép tiến/lùi, bỏ qua và lưu kết quả khảo sát.' },
  { row: 273, ycId: 'YC_16', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Trang chủ sau đăng nhập', moTa: 'Widget hồ sơ và việc phù hợp được cá nhân hoá.' },
  { row: 274, ycId: 'YC_17', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Widget độ hấp dẫn hồ sơ', moTa: 'Điểm hồ sơ và gợi ý bổ sung thông tin.' },
  { row: 276, ycId: 'YC_22', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Banner quảng cáo', moTa: 'Khu banner/carousel trên trang chủ.' },
  { row: 277, ycId: 'YC_29', dsId: 'DS_32', tenTrang: 'Trang Điểm mạnh', tenTinhNang: 'Khai báo điểm mạnh', moTa: 'Thêm và sửa các điểm mạnh của ứng viên.' },
  { row: 278, ycId: 'YC_30', dsId: 'DS_33', tenTrang: 'Trang Ưu tiên công việc', tenTinhNang: 'Ưu tiên công việc', moTa: 'Khai báo loại việc, lương và khu vực mong muốn.' },
  { row: 279, ycId: 'YC_36', dsId: 'DS_29', tenTrang: 'Trang Xem trước hồ sơ', tenTinhNang: 'Template hồ sơ', moTa: 'Chọn mẫu trình bày hồ sơ.' },
  { row: 280, ycId: 'YC_40', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Việc làm của doanh nghiệp', moTa: 'Danh sách việc theo từng doanh nghiệp.' },
  { row: 281, ycId: 'YC_43', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Trạng thái rỗng', moTa: 'Hiển thị đúng khi chưa có dữ liệu.' },
  { row: 282, ycId: 'YC_44', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Ứng tuyển việc làm', moTa: 'Gửi hồ sơ ứng tuyển vào tin tuyển dụng.' },
  { row: 285, ycId: 'YC_55', dsId: '', tenTrang: '', tenTinhNang: 'Trang tĩnh giới thiệu', moTa: 'Trang Về chúng tôi và các trang thông tin.' },
  { row: 286, ycId: 'YC_68', dsId: 'DS_50', tenTrang: 'Trang Quản lý tin', tenTinhNang: 'Cảnh báo thoát/xoá bài đăng', moTa: 'Xác nhận khi thoát hoặc xoá bài đăng chưa lưu.' },
  { row: 287, ycId: 'YC_72', dsId: 'DS_60', tenTrang: 'Vị trí ứng tuyển- Ứng viên quan tâm', tenTinhNang: 'Ứng viên quan tâm', moTa: 'Danh sách ứng viên quan tâm tới vị trí.' },
  { row: 288, ycId: 'YC_73', dsId: 'DS_58', tenTrang: 'Vị trí ứng tuyển- Chuẩn bị phỏng vấn', tenTinhNang: 'Đặt lịch phỏng vấn', moTa: 'Hẹn phỏng vấn online/offline với ngày và liên kết.' },
  { row: 289, ycId: 'YC_77', dsId: 'DS_54', tenTrang: 'Trang Gói dịch vụ', tenTinhNang: 'Gói dịch vụ tuyển dụng', moTa: 'Xem, so sánh và mua các gói dịch vụ.' },
  { row: 290, ycId: 'YC_78', dsId: 'DS_54', tenTrang: 'Trang Gói dịch vụ', tenTinhNang: 'Mã giảm giá / khuyến mãi', moTa: 'Áp mã giảm giá khi mua gói dịch vụ.' },
  { row: 291, ycId: 'YC_80', dsId: 'DS_53', tenTrang: 'Trang Hiệu quả tuyển dụng', tenTinhNang: 'Thống kê số lượng bài đăng', moTa: 'Theo dõi số bài đăng và lịch sử đăng bài.' },
  { row: 292, ycId: 'YC_81', dsId: 'DS_53', tenTrang: 'Trang Hiệu quả tuyển dụng', tenTinhNang: 'Thống kê tương tác bài đăng', moTa: 'Theo dõi tương tác của các bài đăng mới.' },
  { row: 293, ycId: 'YC_83', dsId: '', tenTrang: '', tenTinhNang: 'Chân trang', moTa: 'Danh sách ngành nghề, tuyển dụng, MXH và thông tin công ty.' },
  { row: 294, ycId: 'YC_84', dsId: '', tenTrang: '', tenTinhNang: 'Thanh điều hướng dưới NTD', moTa: 'Điều hướng nhanh các mục chính phía NTD.' },
  { row: 295, ycId: 'YC_85', dsId: '', tenTrang: '', tenTinhNang: 'Phân quyền và chuyển vai trò UV/NTD', moTa: 'Hiển thị đúng theo vai trò và cho phép chuyển vai trò.' },
  { row: 296, ycId: 'YC_86', dsId: '', tenTrang: '', tenTinhNang: 'Thông báo Notification', moTa: 'Hệ thống thông báo cho ứng viên và nhà tuyển dụng.' },
  { row: 297, ycId: 'YC_87', dsId: '', tenTrang: '', tenTinhNang: 'Responsive đa thiết bị', moTa: 'Bố cục đúng trên desktop, tablet và mobile.' },
  { row: 298, ycId: 'YC_88', dsId: 'DS_67, DS_68', tenTrang: 'Scan', tenTinhNang: 'Scan hồ sơ', moTa: 'Scan hồ sơ, scan cv' },
];

// Các YC_ID (255-298) đã TRÙNG với 1 dòng khai báo bug có sẵn ở 5-242 (cùng
// YC_ID, đã có BUG_ID riêng) — theo yêu cầu Tester, các dòng này KHÔNG điền
// lại đầy đủ dữ liệu, chỉ ghi 1 chú thích trùng vào cột T (Ghi chú Tester).
var TINHNANG_TRUNG_LAP = [
  { row: 256, ycId: 'YC_32', ghiChu: 'Trùng với DS_35 / BUG_LO_017 (đã ghi nhận ở dòng 174).' },
  { row: 257, ycId: 'YC_33', ghiChu: 'Trùng với DS_36 / BUG_LO_018 (đã ghi nhận ở dòng 175).' },
  { row: 258, ycId: 'YC_34', ghiChu: 'Trùng với DS_38 / BUG_UI_MOB_018 (đã ghi nhận ở dòng 178).' },
  { row: 259, ycId: 'YC_46', ghiChu: 'Trùng với DS_07 / BUG_FUNC_048 (đã ghi nhận ở dòng 167).' },
  { row: 260, ycId: 'YC_69', ghiChu: 'Trùng với DS_62 / BUG_UI_101 (đã ghi nhận ở dòng 231).' },
  { row: 261, ycId: 'YC_70', ghiChu: 'Trùng với DS_59 / BUG_UI_086 (đã ghi nhận ở dòng 227).' },
  { row: 262, ycId: 'YC_74', ghiChu: 'Trùng với DS_59 / BUG_UI_097 (đã ghi nhận ở dòng 230).' },
  { row: 263, ycId: 'YC_82', ghiChu: 'Trùng với DS_01 / BUG_UI_106 (đã ghi nhận ở dòng 238).' },
  { row: 275, ycId: 'YC_19', ghiChu: 'Trùng với DS_04 / BUG_UI_098 (đã ghi nhận ở dòng 166).' },
  { row: 283, ycId: 'YC_47', ghiChu: 'Trùng với DS_07 / BUG_UI_087 (đã ghi nhận ở dòng 168).' },
  { row: 284, ycId: 'YC_48', ghiChu: 'Trùng với DS_07 / BUG_UI_087 (đã ghi nhận ở dòng 168).' },
];

function fillNoBugRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOI);
  if (!sheet) throw new Error('Không tìm thấy sheet ' + SHEET_LOI + '.');

  var filled = 0;

  TRANG_KHONG_LOI.forEach(function (d) {
    filled += applyKhongLoiRow_(sheet, d.row, {
      D: d.ycId, F: d.moTa, H: 'Không lỗi', J: 'Đã pass',
      Q: NGAY_PHAT_HIEN, R: 'Hoàn thành', S: 'Mr. Ẩn'
    });
  });

  TINHNANG_KHONG_LOI.forEach(function (d) {
    filled += applyKhongLoiRow_(sheet, d.row, {
      A: d.dsId, B: d.tenTrang, C: d.tenTinhNang, F: d.moTa,
      H: 'Không lỗi', J: 'Đã pass', Q: NGAY_PHAT_HIEN, R: 'Hoàn thành', S: 'Mr. Ẩn'
    });
  });

  // Các dòng trùng YC_ID với bug đã khai báo phía trên: chỉ ghi chú, không
  // điền lại A/B/C/F/H/J/Q/R/S.
  TINHNANG_TRUNG_LAP.forEach(function (d) {
    filled += applyKhongLoiRow_(sheet, d.row, { T: d.ghiChu });
  });

  // Điền công thức U/V (Trạng thái lỗi hiện tại / Đã gửi RTM?) cho toàn bộ
  // khối 243-298 — mỗi dòng có formula riêng (không phải mảng spill), nên
  // ghi từng ô là an toàn.
  var formulaFilled = 0;
  for (var row = HANG_DAU_KHONG_LOI; row <= HANG_CUOI_KHONG_LOI; row++) {
    formulaFilled += setFormulaIfBlank_(sheet, row, 21,
      "=IFERROR(XLOOKUP(E" + row + ",'06_RTM_Dev_Test'!$K$5:$K$522,'06_RTM_Dev_Test'!$D$5:$D$522),\"Chưa gửi Dev\")");
    formulaFilled += setFormulaIfBlank_(sheet, row, 22,
      "=IF(COUNTIF('06_RTM_Dev_Test'!$K$5:$K$522,E" + row + ")>0,\"Đã gửi\",\"Chưa gửi\")");
  }

  Logger.log('Đã điền ' + filled + ' ô dữ liệu + ' + formulaFilled + ' ô công thức U/V.');
  SpreadsheetApp.getUi().alert('Hoàn tất!\nĐã điền ' + filled + ' ô dữ liệu (bỏ qua ô đã có sẵn).\nĐã điền ' + formulaFilled + ' công thức cột U/V (dòng 243-298).\n' + TINHNANG_TRUNG_LAP.length + ' dòng trùng YC_ID chỉ được ghi chú tham chiếu, không điền lại dữ liệu.');
}

var COL = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, J: 10, K: 11,
  L: 12, M: 13, N: 14, O: 15, P: 16, Q: 17, R: 18, S: 19, T: 20, U: 21, V: 22 };

/** Ghi các cột trong `fields` (map tên cột A-T -> giá trị) vào `row`, CHỈ khi ô đang trống. */
function applyKhongLoiRow_(sheet, row, fields) {
  var count = 0;
  Object.keys(fields).forEach(function (key) {
    var value = fields[key];
    if (value === undefined || value === null || value === '') return;
    var col = COL[key];
    var cell = sheet.getRange(row, col);
    if (String(cell.getValue()).trim() === '') {
      cell.setValue(value);
      if (key === 'Q') cell.setNumberFormat('dd/MM/yyyy');
      count++;
    }
  });
  return count;
}

/** Đặt formula vào (row, col) nếu ô đang trống, trả về 1 nếu có ghi, 0 nếu bỏ qua. */
function setFormulaIfBlank_(sheet, row, col, formula) {
  var cell = sheet.getRange(row, col);
  if (String(cell.getValue()).trim() === '') {
    cell.setFormula(formula);
    return 1;
  }
  return 0;
}
