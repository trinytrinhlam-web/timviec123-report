# Khôi phục công thức mảng ở 06_RTM_Dev_Test

## Nguyên nhân sự cố (17/08 - phát hiện lúc debug Apps Script)
Sheet `06_RTM_Dev_Test` vốn có sẵn một hệ thống **công thức mảng** (`XLOOKUP` +
`MAP`/`LAMBDA`) neo tại dòng 5, tự "tràn" (spill) xuống để điền toàn bộ thông tin
theo `BUG_ID` (cột K) — tra từ `05_Lỗi_Tester`. Đây chính là tính năng tự động gốc
của sheet.

Apps Script (bản cũ) đã `setValue()`/`clearContent()` vào các dòng nằm giữa vùng
tràn của công thức mảng này → Google Sheets không cho mảng tràn vào ô đã có nội
dung khác → toàn cột báo lỗi `#REF!` ("Array result could not be expanded").

**Từ nay Apps Script CHỈ được đụng vào cột R (Ngày gửi lỗi)** của sheet này — mọi
cột khác (A,B,C,D,E,F,G,H,I,J,L,M,N,O,P,Q,S,AA,AC,AD) do công thức mảng có sẵn lo,
KHÔNG được ghi bằng script.

## Cách khôi phục nếu bị vỡ lại
1. Chọn vùng `A5:AA500` (không cần đụng K, D không cần nếu đã đúng, T, U-Z, AB) →
   **Delete** để xoá sạch nội dung/giá trị còn sót (kể cả cột R — R không dùng công
   thức nữa, để trống cho script quản lý).
2. Dán từng công thức sau vào **đúng dòng 5** của cột tương ứng (chỉ 1 lần, nó tự
   tràn xuống theo dữ liệu ở cột K):

| Ô  | Nguồn 05_Lỗi_Tester | Công thức |
|----|---------------------|-----------|
| A5 | A (DS_ID)           | `=IF($K$5:$K$500="","",IFNA(XLOOKUP($K$5:$K$500,'05_Lỗi_Tester'!$E$5:$E$498,'05_Lỗi_Tester'!$A$5:$A$498),""))` |
| B5 | B (Tên trang)       | thay `$A$5:$A$498` → `$B$5:$B$498` |
| C5 | C (Tính năng)       | → `$C$5:$C$498` |
| D5 | U (Trạng thái lỗi hiện tại) | → `$U$5:$U$498` |
| E5 | F (Mô tả yêu cầu)   | → `$F$5:$F$498` |
| F5 | G (TC_ID(s))        | → `$G$5:$G$498` |
| G5 | I (Mô tả hiện tượng)| → `$I$5:$I$498` |
| H5 | J (Kết quả mong đợi)| → `$J$5:$J$498` |
| I5 | K (Loại lỗi)        | → `$K$5:$K$498` |
| J5 | D (YC_ID)           | → `$D$5:$D$498` |
| L5 | R (Trạng thái kiểm tra nguồn) | → `$R$5:$R$498` |
| M5 | L (Mức độ)          | → `$L$5:$L$498` |
| N5 | M (Thiết bị)        | → `$M$5:$M$498` |
| O5 | N (Trình duyệt)     | → `$N$5:$N$498` |
| P5 | O (Ảnh thiết kế)    | → `$O$5:$O$498` |
| Q5 | P (Bằng chứng thực tế) | → `$P$5:$P$498` |
| S5 | S (Tester)          | → `$S$5:$S$498` |

Cột **R** (Ngày gửi lỗi): **không dán công thức** — để trống, Apps Script
(`stampRtmSubmission_`) sẽ tự ghi ngày hôm nay khi cột K (BUG_ID) thay đổi.

Cột **AC** (`=MAP(K5:K500,R5:R500,X5:X500,Z5:Z500,LAMBDA(...))`) và **AD**
(`=IF(K5:K500="","",IF((L5:L500="Hoàn thành")*...))`) vốn đã có công thức riêng,
tự khỏi khi các cột trên hết `#REF!` — không cần đụng vào.
