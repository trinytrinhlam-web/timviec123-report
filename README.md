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

## Sheet 03_Nghiệm_thu — Biên bản nghiệm thu website

Biên bản nghiệm thu đầy đủ, 18 mục + phụ lục. Quy ước màu:
**ô nền vàng = nhập tay**, ô nền trắng/xám = công thức tự lấy từ dữ liệu QA (không sửa tay).

| Mục | Nội dung | Nguồn |
|---|---|---|
| I | Thông tin chung (số biên bản, dự án, tên miền, môi trường, phiên bản, hợp đồng) | nhập tay |
| II | Các bên tham gia (Bên A / Bên B / PM / QA / Dev / Người duyệt) | nhập tay |
| III | Phạm vi nghiệm thu (DS, YC, vai trò, tích hợp) + loại trừ | nhập tay + tổng tự động |
| IV | Môi trường kiểm thử (trình duyệt, thiết bị, OS, độ phân giải, tài khoản, dữ liệu) | nhập tay |
| V | **Tiêu chí nghiệm thu TC01–TC17** — ngưỡng nhập tay, kết quả & đánh giá tự động | `05_Lỗi_Tester`, `02_UI_UX`, `03_Tính_năng` |
| VI | Định nghĩa mức độ lỗi + SLA xử lý, kèm số lỗi thực tế từng mức | tự động |
| VII | Kết quả kiểm thử — tổng hợp trang / tính năng / lỗi | tự động |
| VIII | Thống kê lỗi theo trạng thái / mức độ / thiết bị / loại | tự động |
| IX | Kết quả nghiệm thu theo 9 hạng mục, kết luận từng hạng mục | tự động + nhập tay |
| X | Kiểm thử phi chức năng (hiệu năng, bảo mật, SEO, accessibility, backup) | nhập tay |
| XI | Lỗi tồn đọng chặn nghiệm thu (Critical/High còn mở, tối đa 30 dòng) | QUERY tự động |
| XII | Ngoại lệ được chấp nhận (waiver) — bắt buộc có người duyệt ký | nhập tay |
| XIII | Rủi ro và hạn chế đã biết | nhập tay |
| XIV | Hạng mục bàn giao (mã nguồn, DB, tài khoản, tài liệu, domain/SSL…) | nhập tay |
| XV | Bảo hành và hỗ trợ sau nghiệm thu | nhập tay |
| XVI | Kết luận nghiệm thu — có **đề xuất tự động** theo tiêu chí mục V | tự động + nhập tay |
| XVII | Việc cần làm để hoàn tất nghiệm thu | nhập tay |
| XVIII | Xác nhận / chữ ký các bên | nhập tay |
| PL | Phụ lục kèm theo (PL-01 → PL-08) | nhập tay |

### Thông tin còn thiếu cần bổ sung
Các mục sau chưa có nguồn dữ liệu trong file QA, phải nhập tay trước khi ký:
thông tin hành chính (mục I, II), phạm vi chốt (III), môi trường kiểm thử (IV),
tỷ lệ Test Case Pass & độ phủ RTM (TC09, TC10 — cần số từ `06_RTM_Dev_Test`),
số đo hiệu năng/bảo mật (X), hạng mục bàn giao (XIV), điều khoản bảo hành (XV).

## Nguồn số liệu
| Chỉ số | Nguồn | Cột |
|---|---|---|
| Trang (tổng 69 / hoàn thành / …) | `02_UI_UX` | dòng tổng hợp `A3:N3` |
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

## Biên bản nghiệm thu Giai đoạn 1 (`BaoCao_NghiemThu_GD1.xlsx`)

Báo cáo nghiệm thu riêng cho 5 màn hình **DS_04, DS_05, DS_21, DS_22, DS_25**
và 12 tính năng liên quan. Đây là văn bản **chốt tại một thời điểm** để ký,
không đồng bộ live như file BGD.

```bash
python3 build_nghiemthu_gd1.py   # đọc data_nghiemthu_gd1.json
```

Dữ liệu trong `data_nghiemthu_gd1.json` trích từ file QA `Tester Detail- Timviec123`
(`02_UI_UX`, `03_Tính_năng`, `05_Lỗi_Tester`, `06_RTM_Dev_Test`, `04_Test_Cases`),
chốt ngày 17/09/2026.

**Quy ước kết luận một DS = ĐẠT** khi đủ 4 điều kiện: trạng thái trang Hoàn thành;
không còn lỗi đang mở; không còn lỗi Critical/High chưa đóng; mọi lỗi đã xử lý có
`Kết quả Retest = Pass`.

Trích lại dữ liệu khi file QA thay đổi:

```bash
python3 extract_nghiemthu_gd1.py <file_QA.xlsx> 17/09/2026
python3 build_nghiemthu_gd1.py
```

Kết quả (chốt 17/09/2026): **5/5 ĐẠT** — 19/19 lỗi đã xử lý và retest Pass,
không còn lỗi Critical/High đang mở.
