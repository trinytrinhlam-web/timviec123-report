/**
 * onEdit HỢP NHẤT cho FILE LÀM VIỆC — gộp toàn bộ script cũ vào 1 hàm onEdit.
 * ---------------------------------------------------------------
 * LƯU Ý: Apps Script chỉ được có DUY NHẤT một hàm tên onEdit. Trước đây bạn có
 * nhiều hàm cùng tên onEdit nên chỉ hàm cuối chạy, các hàm khác bị ghi đè (xung đột).
 * => XÓA HẾT các onEdit cũ ở mọi file .gs, chỉ giữ đúng file này.
 *
 *  05_Lỗi_Tester : nhập BUG_ID (cột E) mà "Ngày phát hiện" (cột Q) đang trống -> điền hôm nay (không ghi đè).
 *  02_UI_UX      : "Trạng thái kiểm tra" (cột E) đổi sang Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (cột P) = hôm nay.
 *  03_Tính_năng  : "Trạng thái kiểm tra" (cột J) đổi sang Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (cột R) = hôm nay.
 */

var HEADER_ROW = 4;                                    // dòng tiêu đề; dữ liệu bắt đầu từ dòng 5
var TRIGGER_VALUES = ['Hoàn thành', 'Kiểm tra lại'];   // trạng thái được coi là "đã test"

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();

  if (name === '05_Lỗi_Tester') {
    // BUG_ID cột E(5) -> Ngày phát hiện cột Q(17)
    stampWhenFilled_(e, sheet, 5, 17);
  } else if (name === '02_UI_UX') {
    // Trạng thái kiểm tra cột E(5) -> Ngày test gần nhất cột P(16)
    stampWhenStatus_(e, sheet, 5, 16);
  } else if (name === '03_Tính_năng') {
    // Trạng thái kiểm tra cột J(10) -> Ngày test gần nhất cột R(18)
    stampWhenStatus_(e, sheet, 10, 18);
  }
}

/** Điền ngày hôm nay vào dateCol khi keyCol có dữ liệu và dateCol đang trống (KHÔNG ghi đè). */
function stampWhenFilled_(e, sheet, keyCol, dateCol) {
  if (!touches_(e.range, keyCol)) return;
  var r = editedRows_(e.range);
  var today = todayDate_();
  for (var row = r[0]; row <= r[1]; row++) {
    var key = sheet.getRange(row, keyCol).getValue();
    var dc = sheet.getRange(row, dateCol);
    if (String(key).trim() !== '' && String(dc.getValue()).trim() === '') {
      dc.setValue(today);
      dc.setNumberFormat('dd/MM/yyyy');
    }
  }
}

/** Điền ngày hôm nay vào dateCol khi statusCol đổi sang Hoàn thành/Kiểm tra lại (GHI ĐÈ). */
function stampWhenStatus_(e, sheet, statusCol, dateCol) {
  if (!touches_(e.range, statusCol)) return;
  var r = editedRows_(e.range);
  var today = todayDate_();
  for (var row = r[0]; row <= r[1]; row++) {
    var status = String(sheet.getRange(row, statusCol).getValue()).trim();
    if (TRIGGER_VALUES.indexOf(status) !== -1) {
      var dc = sheet.getRange(row, dateCol);
      dc.setValue(today);
      dc.setNumberFormat('dd/MM/yyyy');
    }
  }
}

/** Vùng vừa sửa có chạm cột col không (hỗ trợ sửa/kéo nhiều ô). */
function touches_(range, col) {
  var c1 = range.getColumn(), c2 = c1 + range.getNumColumns() - 1;
  return col >= c1 && col <= c2;
}

/** [dòng đầu, dòng cuối] của vùng sửa, đã bỏ qua vùng tiêu đề. */
function editedRows_(range) {
  return [Math.max(range.getRow(), HEADER_ROW + 1), range.getRow() + range.getNumRows() - 1];
}

/** Ngày hôm nay (bỏ giờ). */
function todayDate_() {
  var n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}
