/**
 * onEdit HỢP NHẤT cho FILE LÀM VIỆC (chỉ được có 1 hàm onEdit trong toàn project).
 * ---------------------------------------------------------------
 *  05_Lỗi_Tester : nhập BUG_ID (E) mà "Ngày phát hiện" (Q) đang trống -> điền hôm nay (không ghi đè).
 *  02_UI_UX      : "Trạng thái kiểm tra" (E) -> Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (P) = hôm nay.
 *  03_Tính_năng  : "Trạng thái kiểm tra" (J) -> Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (R) = hôm nay.
 *  06_RTM_Dev_Test: nhập/đổi BUG_ID (K) -> tra theo 05_Lỗi_Tester!E, GHI ĐÈ các cột tự điền;
 *                   "Ngày gửi lỗi" (R) = hôm nay (auto-stamp, không lấy từ 05);
 *                   "Cập nhật gần nhất" (AC) và "Tình trạng tổng" (AD) tự tính lại.
 */

var HEADER_ROW = 4;                                    // dòng tiêu đề; dữ liệu bắt đầu từ dòng 5
var TRIGGER_VALUES = ['Hoàn thành', 'Kiểm tra lại'];   // trạng thái được coi là "đã test"

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();

  if (name === '05_Lỗi_Tester') {
    stampWhenFilled_(e, sheet, 5, 17);          // BUG_ID (E) -> Ngày phát hiện (Q)
  } else if (name === '02_UI_UX') {
    stampWhenStatus_(e, sheet, 5, 16);          // Trạng thái (E) -> Ngày test gần nhất (P)
  } else if (name === '03_Tính_năng') {
    stampWhenStatus_(e, sheet, 10, 18);         // Trạng thái (J) -> Ngày test gần nhất (R)
  } else if (name === '06_RTM_Dev_Test') {
    handleRtmEdit_(e, sheet);
  }
}

/* ============================================================
 *  05_Lỗi_Tester / 02_UI_UX / 03_Tính_năng — tự đóng dấu ngày
 * ============================================================ */

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

/* ============================================================
 *  06_RTM_Dev_Test — nhập BUG_ID (K) -> tự điền theo 05_Lỗi_Tester
 * ============================================================ */

var RTM_COL_BUGID = 11; // K  BUG_ID (khoá tra cứu)
var RTM_COL_D      = 4; // D  Trạng thái lỗi        <- 05.U Trạng thái lỗi hiện tại
var RTM_COL_R      = 18; // R Ngày gửi lỗi           <- auto-stamp hôm nay (KHÔNG lấy từ 05)
var RTM_COL_X      = 24; // X Ngày Dev báo fix       (Dev nhập tay)
var RTM_COL_Z       = 26; // Z Ngày Retest            (Tester nhập tay)
var RTM_COL_AC     = 29; // AC Cập nhật gần nhất     <- tự tính = MAX(R, X, Z)
var RTM_COL_AD     = 30; // AD Tình trạng tổng       <- tự tính theo D

// Cột đích ở RTM -> cột nguồn ở 05_Lỗi_Tester (không gồm K và R, xử lý riêng ở trên).
var RTM_FROM_LOI = {
  1: 1,    // A  DS_ID              <- 05.A  DS_ID
  2: 2,    // B  Tên trang          <- 05.B  Tên trang
  3: 3,    // C  REQ_ID (thực chất là Tên tính năng) <- 05.C
  4: 21,   // D  Trạng thái lỗi     <- 05.U  Trạng thái lỗi hiện tại
  5: 6,    // E  Mô tả yêu cầu      <- 05.F  Mô tả yêu cầu
  6: 7,    // F  TC_ID(s)           <- 05.G  TC_ID(s)
  7: 9,    // G  Mô tả hiện tượng   <- 05.I  Mô tả hiện tượng
  8: 10,   // H  Kết quả mong đợi   <- 05.J  Kết quả mong đợi
  9: 11,   // I  Loại lỗi           <- 05.K  Loại lỗi
  10: 4,   // J  Tên yêu cầu/TN     <- 05.D  YC_ID
  12: 18,  // L  Trạng thái kiểm tra<- 05.R  Trạng thái kiểm tra nguồn
  13: 12,  // M  Mức độ             <- 05.L  Mức độ
  14: 13,  // N  Thiết bị           <- 05.M  Thiết bị
  15: 14,  // O  Trình duyệt        <- 05.N  Trình duyệt
  16: 15,  // P  Ảnh thiết kế       <- 05.O  Ảnh thiết kế (link)
  17: 16,  // Q  Bằng chứng thực tế <- 05.P  Bằng chứng thực tế (link)
  19: 19,  // S  Tester             <- 05.S  Tester
  27: 20   // AA Ghi chú Tester     <- 05.T  Ghi chú Tester
};

function handleRtmEdit_(e, sheet) {
  var range = e.range;
  var touchesBug = touches_(range, RTM_COL_BUGID);
  var touchesComputedTriggers =
      touches_(range, RTM_COL_D) || touches_(range, RTM_COL_R) ||
      touches_(range, RTM_COL_X) || touches_(range, RTM_COL_Z);
  if (!touchesBug && !touchesComputedTriggers) return;

  var rows = editedRows_(range);
  var bugIndex = touchesBug ? buildBugIndex_(sheet) : null;

  for (var row = rows[0]; row <= rows[1]; row++) {
    if (touchesBug) fillRtmFromBug_(sheet, row, bugIndex);
    updateRtmComputed_(sheet, row);
  }
}

/** Đọc toàn bộ 05_Lỗi_Tester một lần -> {BUG_ID: [DS_ID, Tên trang, ... cột A..V]}. */
function buildBugIndex_(rtmSheet) {
  var src = rtmSheet.getParent().getSheetByName('05_Lỗi_Tester');
  var lastRow = src.getLastRow();
  var idx = {};
  if (lastRow < HEADER_ROW + 1) return idx;
  var data = src.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 22).getValues(); // cột A..V
  for (var i = 0; i < data.length; i++) {
    var bugId = normalizeBugId_(data[i][4]); // cột E = BUG_ID (index 4)
    if (bugId !== '') idx[bugId] = data[i];
  }
  return idx;
}

/**
 * Ghi đè các cột tự động của 1 dòng RTM theo BUG_ID (K) — CHỈ khi tìm thấy khớp ở 05_Lỗi_Tester.
 * Không tìm thấy (gõ sai / chưa có trong 05 / BUG_ID để trống) -> KHÔNG đụng tới ô nào,
 * để tránh xóa mất dữ liệu đã có sẵn. Chỉ đánh dấu vào ô ghi chú (AB) để Dev/Tester biết.
 */
function fillRtmFromBug_(sheet, row, bugIndex) {
  var bugId = normalizeBugId_(sheet.getRange(row, RTM_COL_BUGID).getValue());
  var srcRow = bugId !== '' ? bugIndex[bugId] : null;

  if (!srcRow) {
    flagBugNotFound_(sheet, row, bugId);
    return; // không khớp -> giữ nguyên toàn bộ dữ liệu hiện có, không xóa gì cả
  }
  clearBugNotFoundFlag_(sheet, row);

  for (var targetCol in RTM_FROM_LOI) {
    var cell = sheet.getRange(row, Number(targetCol));
    cell.setValue(srcRow[RTM_FROM_LOI[targetCol] - 1]);
  }

  var rCell = sheet.getRange(row, RTM_COL_R); // Ngày gửi lỗi: auto-stamp, ghi đè
  rCell.setValue(todayDate_());
  rCell.setNumberFormat('dd/MM/yyyy');
}

/** Chuẩn hoá BUG_ID để so khớp: bỏ khoảng trắng đầu/cuối và khoảng trắng ẩn (non-breaking space). */
function normalizeBugId_(v) {
  return String(v).replace(/\u00A0/g, ' ').trim();
}

/** Không tìm thấy BUG_ID khớp -> ghi chú cảnh báo vào Ghi chú Dev (AB) để dễ phát hiện, không xóa dữ liệu khác. */
function flagBugNotFound_(sheet, row, bugId) {
  if (bugId === '') return; // K đang trống thì không cần cảnh báo
  var note = '⚠ Không tìm thấy BUG_ID "' + bugId + '" trong 05_Lỗi_Tester (kiểm tra chính tả/khoảng trắng).';
  var abCell = sheet.getRange(row, 28); // AB Ghi chú Dev
  var current = String(abCell.getValue());
  if (current.indexOf('⚠ Không tìm thấy BUG_ID') === -1) {
    abCell.setValue(current ? (note + '\n' + current) : note);
  }
}

/** Xoá cảnh báo "không tìm thấy" (nếu có) khi BUG_ID đã khớp lại được. */
function clearBugNotFoundFlag_(sheet, row) {
  var abCell = sheet.getRange(row, 28); // AB Ghi chú Dev
  var current = String(abCell.getValue());
  if (current.indexOf('⚠ Không tìm thấy BUG_ID') !== -1) {
    var cleaned = current.split('\n').filter(function (line) {
      return line.indexOf('⚠ Không tìm thấy BUG_ID') === -1;
    }).join('\n');
    abCell.setValue(cleaned);
  }
}

/** Tự tính AC = ngày mới nhất trong (R, X, Z); AD theo D (Trạng thái lỗi). */
function updateRtmComputed_(sheet, row) {
  var d = String(sheet.getRange(row, RTM_COL_D).getValue()).trim();
  var r = sheet.getRange(row, RTM_COL_R).getValue();
  var x = sheet.getRange(row, RTM_COL_X).getValue();
  var z = sheet.getRange(row, RTM_COL_Z).getValue();

  var dates = [r, x, z].filter(function (v) { return v instanceof Date; });
  var acCell = sheet.getRange(row, RTM_COL_AC);
  if (dates.length) {
    var latest = new Date(Math.max.apply(null, dates.map(function (dt) { return dt.getTime(); })));
    acCell.setValue(latest);
    acCell.setNumberFormat('dd/MM/yyyy');
  } else {
    acCell.clearContent();
  }

  var adCell = sheet.getRange(row, RTM_COL_AD);
  if (d === '') {
    adCell.clearContent();
  } else if (d === 'Fixed' || d === 'Verified' || d === 'Closed') {
    adCell.setValue('Đã xử lý');
  } else if (d === 'Deferred') {
    adCell.setValue('Tạm hoãn');
  } else {
    adCell.setValue('Đang xử lý');
  }
}

/* ============================================================
 *  Hàm dùng chung
 * ============================================================ */

/** Vùng vừa sửa có chạm cột col không (hỗ trợ sửa/kéo/dán nhiều ô). */
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
