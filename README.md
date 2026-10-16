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

## Nghiệm thu GĐ2 — Nhà tuyển dụng, lần 2 (`nghiem_thu_gd2/`)

Biên bản lần 1 (29/09/2026, file `Biên bản nghiệm thu GD2.xlsx` trên Drive > Report Timviec123 >
Tháng 9) kết luận **CHƯA ĐẠT**. Lần 2 đối chiếu lại từng mã lỗi, luồng F1–F8, 23 yêu cầu,
19 trang và 116 kịch bản của lần 1 với trạng thái QA hiện tại và kết quả kiểm lại trên website.
Ngày 16/10 đã chạy lại toàn bộ 116 kịch bản và 52 lỗi Verified bằng các tài khoản demo:
kết luận lần 2 vẫn **CHƯA ĐẠT** (còn 9/21 mã chặn duyệt).

```bash
cd nghiem_thu_gd2
python3 extract_nghiemthu_gd2.py "<Biên bản nghiệm thu GD2.xlsx>" <file_QA.xlsx> ket_qua_kiem_lai_1610.json 16/10/2026
python3 build_nghiemthu_gd2.py      # -> BaoCao_NghiemThu_GD2_lan2.xlsx
```

- `ket_qua_kiem_lai_1610.json`: kết quả kiểm lại từng mã (`Pass` / `Fail` / `Không xác định`),
  nhận xét, ảnh (`anh_kiem_lai/`) và bằng chứng mã nguồn. Kiểm thêm thì bổ sung vào file này
  rồi chạy lại 2 lệnh trên.
- **Điều kiện ĐẠT**: mọi mã nhóm chặn duyệt (8 lỗi "Cần sửa trước khi duyệt" + 7 lỗi
  "QA - cần xử lý" của lần 1 + lỗi Critical/High phát sinh sau 29/09 trong khu NTD) phải
  kiểm lại đạt. Dev báo đã sửa (Fixed/Verified) mà chưa retest đạt thì chưa tính.
- **Kết luận luồng**: luồng Chưa đạt khi còn mã chặn của luồng / trang luồng đi qua, hoặc có
  kịch bản Lỗi khi chạy lại (giống cách tính lần 1).
- Sheet **Lỗi Verified (Dev báo sửa)**: kết quả kiểm các lỗi đang Verified trong
  `06_RTM_Dev_Test` kèm đề xuất cập nhật QA (chưa ghi vào file QA).
- `kiem_tra/perf4g.js`: đo tốc độ tải trên 4G giả lập (Playwright).
