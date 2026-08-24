/**
 * TỰ ĐỘNG CẬP NHẬT "Ngày test gần nhất"
 * ---------------------------------------------------------------
 * Khi cột "Trạng thái kiểm tra" đổi sang "Hoàn thành" hoặc "Kiểm tra lại",
 * script tự điền "Ngày test gần nhất" = ngày hôm nay (ngày thực hiện thay đổi).
 *
 * - Tự dò cột theo TÊN TIÊU ĐỀ ở dòng 4, nên không phụ thuộc thứ tự cột.
 * - Áp dụng cho mọi sheet có đủ 2 tiêu đề trên (02_UI_UX, 03_Tính_năng).
 * - Là simple trigger onEdit: tự chạy khi người dùng sửa, không cần cài đặt gì thêm.
 *
 * Cài đặt: Tiện ích mở rộng (Extensions) → Apps Script → dán toàn bộ file này → Lưu.
 */

var HEADER_ROW = 4;                              // dòng chứa tiêu đề cột
var STATUS_HEADER = 'Trạng thái kiểm tra';        // tên cột trạng thái
var DATE_HEADER   = 'Ngày test gần nhất';         // tên cột ngày test
var TRIGGER_VALUES = ['Hoàn thành', 'Kiểm tra lại']; // trạng thái coi là "đã test"

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  // Dò vị trí 2 cột theo tiêu đề
  var headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
                  .map(function (x) { return String(x).trim(); });
  var statusCol = headers.indexOf(STATUS_HEADER) + 1;
  var dateCol   = headers.indexOf(DATE_HEADER) + 1;
  if (statusCol < 1 || dateCol < 1) return;      // không phải sheet theo dõi → bỏ qua

  // Chỉ xử lý khi vùng sửa có chạm vào cột Trạng thái
  var startCol = e.range.getColumn();
  var endCol   = startCol + e.range.getNumColumns() - 1;
  if (statusCol < startCol || statusCol > endCol) return;

  var startRow = Math.max(e.range.getRow(), HEADER_ROW + 1);
  var endRow   = e.range.getRow() + e.range.getNumRows() - 1;
  if (endRow < startRow) return;

  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // bỏ giờ, chỉ lấy ngày

  for (var r = startRow; r <= endRow; r++) {
    var status = String(sh.getRange(r, statusCol).getValue()).trim();
    if (TRIGGER_VALUES.indexOf(status) !== -1) {
      var cell = sh.getRange(r, dateCol);
      cell.setValue(today);
      cell.setNumberFormat('dd/mm/yyyy');
    }
    // Nếu muốn XÓA ngày khi trạng thái quay lại "Chưa kiểm tra", bỏ ghi chú 2 dòng dưới:
    // else if (status === 'Chưa kiểm tra') {
    //   sh.getRange(r, dateCol).clearContent();
    // }
  }
}
