/**
 * onEdit HỢP NHẤT cho FILE LÀM VIỆC.
 * ---------------------------------------------------------------
 * QUAN TRỌNG: Apps Script chỉ được có DUY NHẤT một hàm tên onEdit.
 * Nếu để nhiều hàm cùng tên onEdit, chỉ hàm định nghĩa CUỐI CÙNG chạy,
 * các hàm trước bị ghi đè -> gây "xung đột". Vì vậy hãy XÓA mọi hàm onEdit
 * khác (ở tất cả các file .gs trong project) và chỉ dùng file này.
 *
 * Gộp toàn bộ logic tự động:
 *   - 05_Lỗi_Tester : nhập BUG_ID (cột E) mà "Ngày phát hiện" (cột Q) đang trống -> điền hôm nay.
 *   - 02_UI_UX      : "Trạng thái kiểm tra" đổi sang Hoàn thành/Kiểm tra lại -> "Ngày test gần nhất" = hôm nay.
 *   - 03_Tính_năng  : như 02_UI_UX (tự dò cột theo tiêu đề dòng 4, không phụ thuộc thứ tự cột).
 */
var HEADER_ROW = 4;                                   // dòng tiêu đề của các sheet
var TRIGGER_VALUES = ['Hoàn thành', 'Kiểm tra lại'];  // trạng thái coi là "đã test"

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var name = sh.getName();

  if (name === '05_Lỗi_Tester') {
    // Nhập BUG_ID -> tự điền Ngày phát hiện (chỉ khi đang trống, KHÔNG ghi đè)
    stampWhenFilled_(e, sh, 'BUG_ID', 'Ngày phát hiện');
  } else if (name === '02_UI_UX' || name === '03_Tính_năng') {
    // Trạng thái -> Hoàn thành/Kiểm tra lại thì cập nhật Ngày test gần nhất (ghi đè)
    stampWhenStatus_(e, sh, 'Trạng thái kiểm tra', 'Ngày test gần nhất');
  }
}

/** Điền ngày hôm nay vào cột date khi cột key có dữ liệu và cột date đang trống. */
function stampWhenFilled_(e, sh, keyHeader, dateHeader) {
  var cols = findCols_(sh, [keyHeader, dateHeader]);
  if (!cols) return;
  var keyCol = cols[0], dateCol = cols[1];
  if (!touches_(e.range, keyCol)) return;

  var rows = editedRows_(e.range);
  var today = todayDate_();
  for (var r = rows[0]; r <= rows[1]; r++) {
    var key = sh.getRange(r, keyCol).getValue();
    var dateCell = sh.getRange(r, dateCol);
    if (String(key).trim() !== '' && String(dateCell.getValue()).trim() === '') {
      dateCell.setValue(today);
      dateCell.setNumberFormat('dd/MM/yyyy');
    }
  }
}

/** Điền ngày hôm nay vào cột date khi cột status đổi sang Hoàn thành/Kiểm tra lại (ghi đè). */
function stampWhenStatus_(e, sh, statusHeader, dateHeader) {
  var cols = findCols_(sh, [statusHeader, dateHeader]);
  if (!cols) return;
  var statusCol = cols[0], dateCol = cols[1];
  if (!touches_(e.range, statusCol)) return;

  var rows = editedRows_(e.range);
  var today = todayDate_();
  for (var r = rows[0]; r <= rows[1]; r++) {
    var status = String(sh.getRange(r, statusCol).getValue()).trim();
    if (TRIGGER_VALUES.indexOf(status) !== -1) {
      var dateCell = sh.getRange(r, dateCol);
      dateCell.setValue(today);
      dateCell.setNumberFormat('dd/MM/yyyy');
    }
  }
}

/** Vùng sửa có chạm cột col không? */
function touches_(range, col) {
  var c1 = range.getColumn(), c2 = c1 + range.getNumColumns() - 1;
  return col >= c1 && col <= c2;
}

/** Trả về [dòng đầu, dòng cuối] của vùng sửa (đã bỏ qua vùng tiêu đề). */
function editedRows_(range) {
  var start = Math.max(range.getRow(), HEADER_ROW + 1);
  var end = range.getRow() + range.getNumRows() - 1;
  return [start, end];
}

/** Tìm vị trí các cột theo tên tiêu đề ở HEADER_ROW; null nếu thiếu bất kỳ cột nào. */
function findCols_(sh, headerNames) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return null;
  var headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
                  .map(function (x) { return String(x).trim(); });
  var out = [];
  for (var i = 0; i < headerNames.length; i++) {
    var idx = headers.indexOf(headerNames[i]);
    if (idx < 0) return null;
    out.push(idx + 1);
  }
  return out;
}

/** Ngày hôm nay (bỏ giờ). */
function todayDate_() {
  var n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}
