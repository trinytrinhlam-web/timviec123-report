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
 *  2. Với 255-298 (theo Tính năng/YC_ID): điền DS_ID (A), Tên trang (B),
 *     Tính năng (C) và Mô tả yêu cầu (F) — tra thẳng từ 03_Tính_năng theo
 *     YC_ID (cột D).
 *  3. Với cả 2 nhóm: điền các cột trạng thái "không lỗi" (H, J, Q, R, S)
 *     theo đúng khuôn mẫu Tester đã dùng ở dòng 243/244 (Tiêu đề lỗi =
 *     "Không lỗi", Kết quả mong đợi = "Đã pass", Trạng thái kiểm tra nguồn =
 *     "Hoàn thành", Tester = "Mr. Ẩn"). Ngày phát hiện (Q) dùng NGÀY CHẠY
 *     SCRIPT vì không có ngày kiểm tra thực tế ghi lại cho các dòng này —
 *     Tester nên sửa lại cột Q cho đúng ngày đã kiểm tra thực tế nếu khác.
 *  4. Với 243-298: điền công thức 2 cột U (Trạng thái lỗi hiện tại) và V
 *     (Đã gửi RTM?) — ĐÚNG công thức Tester đã dùng sẵn ở các dòng phía trên
 *     (240-243), chỉ neo lại theo số dòng tương ứng. Đây là formula thường
 *     (không phải công thức mảng spill của 06_RTM_Dev_Test) nên an toàn khi
 *     ghi từng ô bằng script.
 *
 * Script CHỈ ghi vào Ô ĐANG TRỐNG — nếu Tester đã tự nhập giá trị nào rồi
 * thì giữ nguyên, không ghi đè. Vì vậy an toàn khi chạy nhiều lần.
 *
 * LƯU Ý: Dòng 298 (YC_88 - Scan hồ sơ) gắn với DS_67/DS_68, hai trang này ở
 * 02_UI_UX vẫn đang "Chưa kiểm tra" (chưa test) — script vẫn điền theo mẫu
 * "không lỗi" nhưng có ghi chú nhắc Tester xác nhận lại trước khi chốt.
 *
 * CÁCH CHẠY:
 *  1. Vào Tiện ích mở rộng → Apps Script (cùng project với onEdit_NgayTest.gs).
 *  2. Tạo FILE MỚI tên tuỳ ý (vd: fill_no_bug_rows.gs), dán toàn bộ nội dung này vào.
 *  3. Trong thanh công cụ, chọn hàm "fillNoBugRows" → bấm Run (▶).
 *  4. Cấp quyền nếu được hỏi lần đầu.
 *  5. Kiểm tra lại 05_Lỗi_Tester dòng 243-298, sửa cột Q (ngày) và Ghi chú
 *     Tester (T) nếu cần cho đúng thực tế.
 */

var SHEET_LOI = '05_Lỗi_Tester';
var HANG_DAU_KHONG_LOI = 243;
var HANG_CUOI_KHONG_LOI = 298;

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
  { row: 255, ycId: 'YC_02', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Đăng nhập bằng mạng xã hội', moTa: 'Đăng nhập nhanh qua tài khoản Google/Facebook.', ghiChu: '' },
  { row: 256, ycId: 'YC_32', dsId: 'DS_34', tenTrang: 'Trang Kinh nghiệm làm việc', tenTinhNang: 'Kinh nghiệm làm việc', moTa: 'Thêm/sửa quá trình làm việc và mô tả kinh nghiệm.', ghiChu: '' },
  { row: 257, ycId: 'YC_33', dsId: 'DS_36', tenTrang: 'Trang Giới thiệu bản thân', tenTinhNang: 'Giới thiệu bản thân', moTa: 'Viết đoạn giới thiệu bản thân.', ghiChu: '' },
  { row: 258, ycId: 'YC_34', dsId: 'DS_38', tenTrang: 'Trang Chứng nhận/Giải thưởng', tenTinhNang: 'Chứng nhận / Giải thưởng', moTa: 'Thêm chứng nhận và giải thưởng.', ghiChu: '' },
  { row: 259, ycId: 'YC_46', dsId: 'DS_08', tenTrang: 'Khung Tin nhắn Window user', tenTinhNang: 'Popup chat nổi', moTa: 'Khung chat popup truy cập nhanh từ mọi trang.', ghiChu: '' },
  { row: 260, ycId: 'YC_69', dsId: 'DS_62', tenTrang: 'Ứng viên gần đây', tenTinhNang: 'Ứng viên gần đây', moTa: 'Danh sách ứng viên mới hoặc hồ sơ phù hợp.', ghiChu: '' },
  { row: 261, ycId: 'YC_70', dsId: 'DS_59', tenTrang: 'Vị trí ứng tuyển- Hồ sơ ứng tuyển & Đề xuất', tenTinhNang: 'Xem hồ sơ ứng viên', moTa: 'NTD xem chi tiết hồ sơ và mời gửi hồ sơ.', ghiChu: '' },
  { row: 262, ycId: 'YC_74', dsId: 'DS_59', tenTrang: 'Vị trí ứng tuyển- Hồ sơ ứng tuyển & Đề xuất', tenTinhNang: 'Đề xuất ứng viên phù hợp', moTa: 'Gợi ý ứng viên phù hợp với tin tuyển dụng.', ghiChu: '' },
  { row: 263, ycId: 'YC_82', dsId: '', tenTrang: '', tenTinhNang: 'Điều hướng đầu trang đa cấp', moTa: 'Menu chính đa cấp, tin nhắn, thông báo và tài khoản.', ghiChu: '' },
  { row: 264, ycId: 'YC_03', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Đăng nhập bằng Zalo qua OTP', moTa: 'Nhập SĐT Zalo, hệ thống gửi và xác minh mã OTP.', ghiChu: '' },
  { row: 265, ycId: 'YC_05', dsId: 'DS_22', tenTrang: 'Trang Đăng ký', tenTinhNang: 'Đăng ký bằng Zalo', moTa: 'Tạo tài khoản bằng SĐT Zalo và xác minh OTP.', ghiChu: '' },
  { row: 266, ycId: 'YC_06', dsId: 'DS_31', tenTrang: 'Trang Xác thực tài khoản', tenTinhNang: 'Xác minh email / Zalo', moTa: 'Xác minh tài khoản qua mã hoặc liên kết email hay OTP Zalo.', ghiChu: '' },
  { row: 267, ycId: 'YC_07', dsId: 'DS_22', tenTrang: 'Trang Đăng ký', tenTinhNang: 'Thông báo đăng ký thành công', moTa: 'Xác nhận đăng ký thành công và điều hướng sang khảo sát.', ghiChu: '' },
  { row: 268, ycId: 'YC_09', dsId: 'DS_25', tenTrang: 'Trang Quên mật khẩu', tenTinhNang: 'Xác minh khi quên mật khẩu', moTa: 'Nhập mã xác minh trước khi đặt lại mật khẩu.', ghiChu: '' },
  { row: 269, ycId: 'YC_10', dsId: 'DS_18', tenTrang: 'Trang Đổi mật khẩu', tenTinhNang: 'Đổi mật khẩu', moTa: 'Nhập mật khẩu hiện tại và mật khẩu mới; thông báo thành công.', ghiChu: '' },
  { row: 270, ycId: 'YC_11', dsId: 'DS_21', tenTrang: 'Trang Đăng nhập', tenTinhNang: 'Phân luồng Ứng viên và Nhà tuyển dụng', moTa: 'Điều hướng đúng luồng UV và NTD sau đăng nhập.', ghiChu: '' },
  { row: 271, ycId: 'YC_12', dsId: 'DS_23', tenTrang: 'Trang Đăng ký - Ứng viên', tenTinhNang: 'Khảo sát ứng viên khi onboarding', moTa: 'Thu thập công việc, kinh nghiệm và khu vực để cá nhân hoá gợi ý.', ghiChu: '' },
  { row: 272, ycId: 'YC_14', dsId: 'DS_23, DS_24', tenTrang: 'Trang Đăng ký - Ứng viên, Trang Đăng ký - Nhà tuyển dụng', tenTinhNang: 'Điều hướng bước và lưu khảo sát', moTa: 'Cho phép tiến/lùi, bỏ qua và lưu kết quả khảo sát.', ghiChu: '' },
  { row: 273, ycId: 'YC_16', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Trang chủ sau đăng nhập', moTa: 'Widget hồ sơ và việc phù hợp được cá nhân hoá.', ghiChu: '' },
  { row: 274, ycId: 'YC_17', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Widget độ hấp dẫn hồ sơ', moTa: 'Điểm hồ sơ và gợi ý bổ sung thông tin.', ghiChu: '' },
  { row: 275, ycId: 'YC_19', dsId: 'DS_04', tenTrang: 'Trang Tìm kiếm', tenTinhNang: 'Bộ lọc tìm kiếm nâng cao', moTa: 'Lọc theo lương, khu vực, ngành nghề, kinh nghiệm và loại hình.', ghiChu: '' },
  { row: 276, ycId: 'YC_22', dsId: 'DS_01', tenTrang: 'Trang chủ', tenTinhNang: 'Banner quảng cáo', moTa: 'Khu banner/carousel trên trang chủ.', ghiChu: '' },
  { row: 277, ycId: 'YC_29', dsId: 'DS_32', tenTrang: 'Trang Điểm mạnh', tenTinhNang: 'Khai báo điểm mạnh', moTa: 'Thêm và sửa các điểm mạnh của ứng viên.', ghiChu: '' },
  { row: 278, ycId: 'YC_30', dsId: 'DS_33', tenTrang: 'Trang Ưu tiên công việc', tenTinhNang: 'Ưu tiên công việc', moTa: 'Khai báo loại việc, lương và khu vực mong muốn.', ghiChu: '' },
  { row: 279, ycId: 'YC_36', dsId: 'DS_29', tenTrang: 'Trang Xem trước hồ sơ', tenTinhNang: 'Template hồ sơ', moTa: 'Chọn mẫu trình bày hồ sơ.', ghiChu: '' },
  { row: 280, ycId: 'YC_40', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Việc làm của doanh nghiệp', moTa: 'Danh sách việc theo từng doanh nghiệp.', ghiChu: '' },
  { row: 281, ycId: 'YC_43', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Trạng thái rỗng', moTa: 'Hiển thị đúng khi chưa có dữ liệu.', ghiChu: '' },
  { row: 282, ycId: 'YC_44', dsId: 'DS_03', tenTrang: 'Trang DN địa phương chi tiết', tenTinhNang: 'Ứng tuyển việc làm', moTa: 'Gửi hồ sơ ứng tuyển vào tin tuyển dụng.', ghiChu: '' },
  { row: 283, ycId: 'YC_47', dsId: 'DS_07', tenTrang: 'Khung Hộp thư + tin nhắn', tenTinhNang: 'Tin nhắn nhà tuyển dụng', moTa: 'NTD chat với ứng viên và quản lý hội thoại.', ghiChu: '' },
  { row: 284, ycId: 'YC_48', dsId: 'DS_07', tenTrang: 'Khung Hộp thư + tin nhắn', tenTinhNang: 'Gửi/nhận và trạng thái tin nhắn', moTa: 'Gửi nhận realtime, trạng thái đã đọc và mời gửi/xem hồ sơ.', ghiChu: '' },
  { row: 285, ycId: 'YC_55', dsId: '', tenTrang: '', tenTinhNang: 'Trang tĩnh giới thiệu', moTa: 'Trang Về chúng tôi và các trang thông tin.', ghiChu: '' },
  { row: 286, ycId: 'YC_68', dsId: 'DS_50', tenTrang: 'Trang Quản lý tin', tenTinhNang: 'Cảnh báo thoát/xoá bài đăng', moTa: 'Xác nhận khi thoát hoặc xoá bài đăng chưa lưu.', ghiChu: '' },
  { row: 287, ycId: 'YC_72', dsId: 'DS_60', tenTrang: 'Vị trí ứng tuyển- Ứng viên quan tâm', tenTinhNang: 'Ứng viên quan tâm', moTa: 'Danh sách ứng viên quan tâm tới vị trí.', ghiChu: '' },
  { row: 288, ycId: 'YC_73', dsId: 'DS_58', tenTrang: 'Vị trí ứng tuyển- Chuẩn bị phỏng vấn', tenTinhNang: 'Đặt lịch phỏng vấn', moTa: 'Hẹn phỏng vấn online/offline với ngày và liên kết.', ghiChu: '' },
  { row: 289, ycId: 'YC_77', dsId: 'DS_54', tenTrang: 'Trang Gói dịch vụ', tenTinhNang: 'Gói dịch vụ tuyển dụng', moTa: 'Xem, so sánh và mua các gói dịch vụ.', ghiChu: '' },
  { row: 290, ycId: 'YC_78', dsId: 'DS_54', tenTrang: 'Trang Gói dịch vụ', tenTinhNang: 'Mã giảm giá / khuyến mãi', moTa: 'Áp mã giảm giá khi mua gói dịch vụ.', ghiChu: '' },
  { row: 291, ycId: 'YC_80', dsId: 'DS_53', tenTrang: 'Trang Hiệu quả tuyển dụng', tenTinhNang: 'Thống kê số lượng bài đăng', moTa: 'Theo dõi số bài đăng và lịch sử đăng bài.', ghiChu: '' },
  { row: 292, ycId: 'YC_81', dsId: 'DS_53', tenTrang: 'Trang Hiệu quả tuyển dụng', tenTinhNang: 'Thống kê tương tác bài đăng', moTa: 'Theo dõi tương tác của các bài đăng mới.', ghiChu: '' },
  { row: 293, ycId: 'YC_83', dsId: '', tenTrang: '', tenTinhNang: 'Chân trang', moTa: 'Danh sách ngành nghề, tuyển dụng, MXH và thông tin công ty.', ghiChu: '' },
  { row: 294, ycId: 'YC_84', dsId: '', tenTrang: '', tenTinhNang: 'Thanh điều hướng dưới NTD', moTa: 'Điều hướng nhanh các mục chính phía NTD.', ghiChu: '' },
  { row: 295, ycId: 'YC_85', dsId: '', tenTrang: '', tenTinhNang: 'Phân quyền và chuyển vai trò UV/NTD', moTa: 'Hiển thị đúng theo vai trò và cho phép chuyển vai trò.', ghiChu: '' },
  { row: 296, ycId: 'YC_86', dsId: '', tenTrang: '', tenTinhNang: 'Thông báo Notification', moTa: 'Hệ thống thông báo cho ứng viên và nhà tuyển dụng.', ghiChu: '' },
  { row: 297, ycId: 'YC_87', dsId: '', tenTrang: '', tenTinhNang: 'Responsive đa thiết bị', moTa: 'Bố cục đúng trên desktop, tablet và mobile.', ghiChu: '' },
  { row: 298, ycId: 'YC_88', dsId: 'DS_67, DS_68', tenTrang: 'Scan', tenTinhNang: 'Scan hồ sơ', moTa: 'Scan hồ sơ, scan cv', ghiChu: 'DS_67/DS_68 hiện vẫn "Chưa kiểm tra" ở 02_UI_UX — nhờ Tester xác nhận lại trạng thái trang trước khi chốt dòng này.' },
];

function fillNoBugRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOI);
  if (!sheet) throw new Error('Không tìm thấy sheet ' + SHEET_LOI + '.');

  var today = new Date();
  var filled = 0;

  TRANG_KHONG_LOI.forEach(function (d) {
    filled += applyKhongLoiRow_(sheet, d.row, {
      D: d.ycId, F: d.moTa, H: 'Không lỗi', J: 'Đã pass',
      Q: today, R: 'Hoàn thành', S: 'Mr. Ẩn'
    });
  });

  TINHNANG_KHONG_LOI.forEach(function (d) {
    filled += applyKhongLoiRow_(sheet, d.row, {
      A: d.dsId, B: d.tenTrang, C: d.tenTinhNang, F: d.moTa,
      H: 'Không lỗi', J: 'Đã pass', Q: today, R: 'Hoàn thành', S: 'Mr. Ẩn',
      T: d.ghiChu
    });
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
  SpreadsheetApp.getUi().alert('Hoàn tất!\nĐã điền ' + filled + ' ô dữ liệu (bỏ qua ô đã có sẵn).\nĐã điền ' + formulaFilled + ' công thức cột U/V (dòng 243-298).\n\nNhớ kiểm tra lại cột Q (Ngày phát hiện) và Ghi chú Tester (T) cho đúng thực tế.');
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
