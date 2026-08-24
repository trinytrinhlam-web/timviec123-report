/**
 * onEdit HỢP NHẤT cho FILE LÀM VIỆC (chỉ được có 1 hàm onEdit trong toàn project).
 * ---------------------------------------------------------------
 * LƯU Ý QUAN TRỌNG VỀ 06_RTM_Dev_Test:
 *   Sheet này ĐÃ CÓ SẴN hệ thống công thức mảng (XLOOKUP/MAP/LAMBDA) neo tại dòng 5,
 *   tự tràn (spill) xuống để điền A,B,C,D,E,F,G,H,I,J,L,M,N,O,P,Q,S,AA,AC,AD theo BUG_ID (K).
 *   TUYỆT ĐỐI KHÔNG dùng setValue()/clearContent() vào các cột đó bằng script — sẽ làm vỡ
 *   vùng tràn và gây lỗi #REF! toàn cột. Apps Script ở đây CHỈ xử lý cột R (Ngày gửi lỗi),
 *   vì R cần "đóng dấu đúng thời điểm gửi" — điều mà công thức không làm được (không có
 *   trạng thái/thời điểm để nhớ), còn lại để nguyên cho công thức mảng lo.
 *
 *  05_Lỗi_Tester : nhập BUG_ID (E) mà "Ngày phát hiện" (Q) đang trống -> điền hôm nay (không ghi đè).
 *  02_UI_UX      : "Trạng thái kiểm tra" (E) -> Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (P) = hôm nay.
 *  03_Tính_năng  : "Trạng thái kiểm tra" (J) -> Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" (R) = hôm nay.
 *  06_RTM_Dev_Test: nhập/đổi BUG_ID (K) -> "Ngày gửi lỗi" (R, cột 18) = hôm nay (ghi đè);
 *                   K bị xoá -> xoá R theo. KHÔNG đụng tới bất kỳ cột nào khác trong sheet này.
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
    stampRtmSubmission_(e, sheet, 11, 18);      // BUG_ID (K) -> Ngày gửi lỗi (R), GHI ĐÈ
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

/**
 * 06_RTM_Dev_Test: khi keyCol (BUG_ID) đổi -> dateCol (Ngày gửi lỗi) = hôm nay, GHI ĐÈ mỗi lần đổi.
 * keyCol trống -> xoá dateCol theo. CHỈ đụng tới dateCol, không đụng bất kỳ cột nào khác
 * (các cột còn lại do công thức mảng XLOOKUP/MAP/LAMBDA có sẵn trong sheet tự lo).
 */
function stampRtmSubmission_(e, sheet, keyCol, dateCol) {
  if (!touches_(e.range, keyCol)) return;
  var r = editedRows_(e.range);
  var today = todayDate_();
  for (var row = r[0]; row <= r[1]; row++) {
    var key = String(sheet.getRange(row, keyCol).getValue()).trim();
    var dc = sheet.getRange(row, dateCol);
    if (key !== '') {
      dc.setValue(today);
      dc.setNumberFormat('dd/MM/yyyy');
    } else {
      dc.clearContent();
    }
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
