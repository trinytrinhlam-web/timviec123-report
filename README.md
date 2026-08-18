# Báo cáo tiến độ Website TimViec123 — dành cho BGD

File `BaoCao_BGD_TimViec123.xlsx` là báo cáo tiến độ website, **đồng bộ sống (live)** bằng
`IMPORTRANGE` từ **2 nguồn**:
- **File làm việc QA** (`QA Operations Master`) — tiến độ trang/tính năng, lỗi, công việc.
- **File Database** (`Hồ sơ`) — số DN / HR / Ứng viên / Tin tuyển dụng.

## Cách đưa lên Google Sheet
1. Kéo file `.xlsx` vào **Google Drive** → chuột phải → **Mở bằng → Google Sheets**
   (tự chuyển, giữ nguyên công thức + định dạng).
2. Vào sheet **00_Hướng_dẫn**, phần **KÍCH HOẠT KẾT NỐI**: nếu ô **① File làm việc QA**
   hoặc **② File Database** báo `#REF!`, bấm vào ô đó → **Cho phép truy cập / Allow access**.
   Làm 1 lần cho mỗi ô. Sau đó toàn bộ báo cáo có số.

## Sheet Tháng & Tuần (cùng một layout, chỉ khác kỳ lọc)
Hai sheet dùng chung 1 bố cục; sheet Tuần lọc theo tuần-trong-tháng (Tuần 1 = 1–7, …).
- **Bảng KPI** có cột **Mục tiêu kỳ** — tự tra từ sheet **`04_Mục_tiêu`** theo đúng
  tháng/tuần đang chọn (mỗi kỳ một dòng, không dùng chung). Tỷ lệ đạt = Kết quả / Mục tiêu.
  - I. Kiểm tra website: Trang/Tính năng hoàn thành, Công việc & Lỗi phát hiện trong kỳ.
  - II. Hoàn thiện: Lỗi đã xử lý (lũy kế), Số Trang / Tính năng đã xử lý.
  - III. Database (trong kỳ): DN / HR / Ứng viên / Tin tuyển dụng.
- **Tổng quan hiện tại**: trang, tính năng, lỗi đang mở, Critical/High + Database (tổng).
- **Tình trạng xử lý lỗi**: Open / Chưa gửi Dev / Fixed / Verified / Closed / Deferred.
- **Phân bố lỗi phát hiện trong kỳ**: theo mức độ & theo thiết bị.
- **Thanh tiến độ**.
- **Chi tiết Trang & Tính năng đã test trong kỳ** (hai bảng cạnh nhau): liệt kê từ
  `02_UI_UX` / `03_Tính_năng` các mục có **Ngày test gần nhất** trong kỳ, trạng thái
  Hoàn thành/Kiểm tra lại — cột thứ 3 là **Trạng thái** (báo cáo Tháng) hoặc
  **Ngày test/retest** (báo cáo Tuần).

## Sheet 04_Mục_tiêu
Nơi BGD/Tester nhập **mục tiêu KPI theo từng kỳ** (ô nền vàng):
- **Mục tiêu theo Tháng**: mỗi dòng là 1 tháng (1–12), cột là 11 chỉ số KPI.
- **Mục tiêu theo Tuần**: mỗi dòng là 1 (Tháng, Tuần), cột là 11 chỉ số KPI.

Báo cáo Tháng/Tuần tự tra đúng dòng theo kỳ đang chọn → đổi tháng thì mục tiêu đổi theo.

## Nguồn số liệu
| Chỉ số | Nguồn | Cột |
|---|---|---|
| Trang (tổng 63 / hoàn thành / …) | `02_UI_UX` | dòng tổng hợp `A3:N3` |
| Tính năng (88 / hoàn thành / …) | `03_Tính_năng` | dòng tổng hợp `A3:N3` |
| **Lỗi (toàn bộ)** | **`05_Lỗi_Tester`** | E=BUG_ID, L=Mức độ, M=Thiết bị, Q=Ngày phát hiện, U=Trạng thái lỗi |
| Công việc theo kỳ | `01_Cong_viec_ngay` | B=Ngày, O=Trạng thái |
| DB Doanh nghiệp | `Hồ sơ` → `HS DN` | D=Tên, N=Ngày Đăng/Tạo |
| DB HR | `Hồ sơ` → `HS HR` | D=Tên, E=Ngày Đăng/Tạo |
| DB Ứng viên | `Hồ sơ` → `HS ỨNG TUYỂN` | D=Họ tên (không có cột ngày → tính theo tổng) |
| Tin tuyển dụng | `Hồ sơ` → `Tin tuyển dụng` | C=Nội dung, D=Ngày viết |

### Quy ước tính lỗi (từ 05_Lỗi_Tester)
- **Tổng lỗi** = số dòng có BUG_ID.
- **Đang mở** = trạng thái `Open` + `Chưa gửi Dev` (chưa xử lý xong).
- **Đã xử lý** = `Fixed` + `Verified` + `Closed`. `Deferred` = tạm hoãn (tính riêng).

> Sheet Tuần và Nghiệm thu cũng đã đổi nguồn lỗi sang `05_Lỗi_Tester` cho nhất quán.

## Tạo lại file
```bash
python3 build_report.py   # cần openpyxl
```
