# Báo cáo tiến độ Website TimViec123 — dành cho BGD

File `BaoCao_BGD_TimViec123.xlsx` là báo cáo tiến độ website, **đồng bộ sống (live)** từ
file làm việc của Tester/Dev bằng công thức `IMPORTRANGE`. Khi mở bằng Google Sheets, số
liệu tự cập nhật theo file làm việc.

## Cách đưa lên Google Sheet
1. Vào Google Drive → **Mới / New → Tải tệp lên / File upload** → chọn file `.xlsx` này
   (hoặc kéo–thả file vào Drive).
2. Bấm chuột phải file → **Mở bằng / Open with → Google Sheets**. Google tự chuyển thành
   Google Sheet, giữ nguyên công thức + định dạng.
3. Lần đầu mở: tại sheet **00_Hướng_dẫn**, ô **① Kích hoạt kết nối** sẽ báo `#REF!` —
   bấm vào ô đó → **Cho phép truy cập / Allow access**. Chỉ làm 1 lần; sau đó toàn bộ báo
   cáo có số.

## Các sheet
- **00_Hướng_dẫn** — người dùng, nguyên tắc tính số, liên kết, kích hoạt kết nối.
- **01_Báo_cáo_Tháng** — KPI theo tháng (chọn tháng ở ô B4) + tổng quan hiện tại + phân bố lỗi.
- **02_Báo_cáo_Tuần** — chỉ số theo tuần (Tuần 1 = ngày 1–7, Tuần 2 = 8–14…).
- **03_Nghiệm_thu** — checklist nghiệm thu theo hạng mục.
- **z_*** (ẩn) — sheet kéo dữ liệu thô bằng `IMPORTRANGE`, phục vụ tính COUNTIF/COUNTIFS.

## Nguồn dữ liệu (file làm việc)
Kéo trực tiếp từ các sheet của file làm việc QA:

| Chỉ số báo cáo | Sheet nguồn | Ô/Cột |
|---|---|---|
| Trang: tổng/hoàn thành/tiến độ/kiểm tra lại/chưa kiểm tra | `02_UI_UX` | dòng tổng hợp `A3:N3` |
| Tính năng: tổng/hoàn thành/tiến độ/kiểm tra lại/chưa kiểm tra | `03_Tính_năng` | dòng tổng hợp `A3:N3` |
| Lỗi mở theo mức độ (Critical/High/Medium/Low) | `00_Dashboard_QA` | `E17:E20` |
| Tổng lỗi / Lỗi đang mở / Ngày cập nhật | `00A_Tong_hop_Trang_Tinh_nang` | `A4:V5` |
| Lỗi phát hiện / đã fix theo kỳ, phân bố thiết bị | `04_Bug_Master_RTM` | `B`=Ngày phát hiện, `L`=Mức độ, `M`=Thiết bị, `W`=Ngày Dev báo fix |
| Công việc hoàn thành/đang làm theo kỳ | `01_Cong_viec_ngay` | `B`=Ngày, `O`=Trạng thái |

> **Lưu ý số liệu:** báo cáo lấy theo **mô hình dashboard hiện hành** của file làm việc
> (VD: 234 tổng lỗi / 157 đang mở, tính năng 20/88). Con số này **mới hơn** file báo cáo
> mẫu cũ (202 / 68, tính năng 3/88) vì file mẫu đã lâu không cập nhật.

## Tạo lại file
```bash
python3 build_report.py   # cần openpyxl
```
