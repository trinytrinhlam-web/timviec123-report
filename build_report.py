# -*- coding: utf-8 -*-
"""Build BGD progress report (.xlsx) with live IMPORTRANGE links to the QA working file."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

WORK_URL = "https://docs.google.com/spreadsheets/d/1J2A_OiWz7IeYsgafLDY3mIfZsG-xCNGhkNaSQkhicHg/edit"

# ---------- Palette ----------
NAVY   = "1F3A5F"   # title band
BLUE   = "2E5A88"   # section band
LBLUE  = "DCE6F1"   # light header
CARD   = "EEF3F9"   # kpi card bg
ZEBRA  = "F5F7FA"
GREEN  = "2E7D32"; GREEN_BG="E8F5E9"
AMBER  = "B8860B"; AMBER_BG="FFF8E1"
RED    = "C62828"; RED_BG ="FDECEA"
GREY   = "6B7280"
WHITE  = "FFFFFF"
BORDER_CLR = "C9D6E5"

thin = Side(style="thin", color=BORDER_CLR)
box  = Border(left=thin, right=thin, top=thin, bottom=thin)

def F(size=11, bold=False, color="1A1A1A", italic=False):
    return Font(name="Arial", size=size, bold=bold, color=color, italic=italic)
def Fill(c):
    return PatternFill("solid", fgColor=c)
def Al(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def cell(ws, ref, val=None, font=None, fill=None, al=None, nfmt=None, border=None):
    c = ws[ref]
    if val is not None: c.value = val
    if font: c.font = font
    if fill: c.fill = fill
    if al: c.alignment = al
    if nfmt: c.number_format = nfmt
    if border: c.border = border
    return c

def merge(ws, rng, val=None, font=None, fill=None, al=None):
    ws.merge_cells(rng)
    tl = rng.split(":")[0]
    cell(ws, tl, val, font, fill, al)
    # apply fill across merged range for consistent background
    if fill:
        from openpyxl.utils.cell import range_boundaries
        c1,r1,c2,r2 = range_boundaries(rng)
        for r in range(r1,r2+1):
            for c in range(c1,c2+1):
                ws.cell(r,c).fill = fill

def band(ws, rng, text, big=False):
    merge(ws, rng, text,
          font=F(15 if big else 11.5, True, WHITE),
          fill=Fill(NAVY if big else BLUE),
          al=Al("left","center"))
    # row height
    r = int(''.join(ch for ch in rng.split(":")[0] if ch.isdigit()))
    ws.row_dimensions[r].height = 30 if big else 22

PCT = '0.0%'
INT = '#,##0'

wb = openpyxl.Workbook()

# =====================================================================
#  HELPER SHEETS (hidden) — one IMPORTRANGE each
# =====================================================================
def U():  # url reference
    return "'z_Config'!$B$1"

zcfg = wb.active; zcfg.title = "z_Config"
zcfg["A1"]="URL file làm việc"; zcfg["B1"]=WORK_URL
zcfg["A2"]="Năm báo cáo";       zcfg["B2"]=2026
zcfg["A1"].font=F(bold=True); zcfg["A2"].font=F(bold=True)
zcfg.column_dimensions["A"].width=22; zcfg.column_dimensions["B"].width=70

def src(name, rng):
    ws = wb.create_sheet(name)
    # RAW importrange (no IFERROR) so the "Allow access" prompt can appear / connection works
    ws["A1"] = f'=IMPORTRANGE({U()},"{rng}")'
    ws.sheet_state = "hidden"
    return ws

src("z_Trang",    "02_UI_UX!A3:N3")
src("z_TinhNang", "03_Tính_năng!A3:N3")
src("z_Dash",     "00_Dashboard_QA!A1:L22")
src("z_Tong",     "00A_Tong_hop_Trang_Tinh_nang!A4:V5")
src("z_Bug",      "04_Bug_Master_RTM!A4:AE1004")
src("z_CV",       "01_Cong_viec_ngay!A4:P1004")
src("z_Week",     "09_Release_Tuan!A4:D54")
zcfg.sheet_state = "hidden"

# shorthand refs to live cells
T_TOTAL="'z_Trang'!$B$1"; T_DONE="'z_Trang'!$D$1"; T_PROG="'z_Trang'!$F$1"
T_INTEST="'z_Trang'!$H$1"; T_RETEST="'z_Trang'!$J$1"; T_UNTEST="'z_Trang'!$L$1"
F_TOTAL="'z_TinhNang'!$B$1"; F_DONE="'z_TinhNang'!$D$1"; F_PROG="'z_TinhNang'!$F$1"
F_INTEST="'z_TinhNang'!$H$1"; F_RETEST="'z_TinhNang'!$J$1"; F_UNTEST="'z_TinhNang'!$L$1"; F_FLOW="'z_TinhNang'!$N$1"
BUG_TOTAL="'z_Tong'!$J$2"; BUG_OPEN="'z_Tong'!$L$2"; UPD="'z_Tong'!$T$2"
SEV_C="'z_Dash'!$E$17"; SEV_H="'z_Dash'!$E$18"; SEV_M="'z_Dash'!$E$19"; SEV_L="'z_Dash'!$E$20"
CV_PLAN="'z_Dash'!$A$8"; CHO_DEV="'z_Dash'!$G$12"

# =====================================================================
#  SHEET: 00_Hướng_dẫn
# =====================================================================
g = wb.create_sheet("00_Hướng_dẫn")
for col,w in {"A":26,"B":30,"C":30,"D":22,"E":22,"F":18}.items():
    g.column_dimensions[col].width=w
band(g,"A1:F1","BÁO CÁO TIẾN ĐỘ WEBSITE TIMVIEC123 — DÀNH CHO BAN GIÁM ĐỐC",big=True)
merge(g,"A2:F2","Số liệu đồng bộ trực tiếp (live) từ file làm việc của Tester/Dev. BGD chỉ xem, chọn kỳ ở sheet Tháng/Tuần.",
      font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
g.row_dimensions[2].height=20

merge(g,"A4:B4","NGƯỜI DÙNG",font=F(11,True,WHITE),fill=Fill(BLUE))
merge(g,"C4:F4","CÁCH SỬ DỤNG",font=F(11,True,WHITE),fill=Fill(BLUE))
rows=[("Ban Giám đốc","Xem tiến độ trang/tính năng, tình trạng lỗi và nghiệm thu. Chọn Tháng/Tuần để lọc kỳ."),
      ("Tester","Cập nhật dữ liệu tại file làm việc; báo cáo này tự đồng bộ."),
      ("Dev","Cập nhật xử lý lỗi tại 04_Bug_Master_RTM / 06_RTM_Dev_Test của file làm việc.")]
r=5
for who,how in rows:
    merge(g,f"A{r}:B{r}",who,font=F(11,True),fill=Fill(LBLUE),al=Al("left"))
    merge(g,f"C{r}:F{r}",how,font=F(10.5),al=Al("left",wrap=True))
    g.row_dimensions[r].height=32; r+=1

r+=1
band(g,f"A{r}:F{r}","NGUYÊN TẮC TÍNH SỐ")
r+=1
principles=[
 "Tiến độ trang = số trang Hoàn thành / tổng 56 trang.",
 "Tiến độ tính năng = số tính năng Hoàn thành / tổng 88 tính năng (gồm YC_88 Scan đang chờ chốt flow).",
 "Lỗi đang mở, phân bố mức độ lấy theo dashboard QA hiện hành của file làm việc.",
 "Chỉ số theo Tháng/Tuần dùng Ngày phát hiện lỗi và Ngày công việc trong file làm việc.",
 "Một lỗi có thể ở nhiều thiết bị, nên tổng theo thiết bị có thể lớn hơn tổng số lỗi."]
for i,p in enumerate(principles,1):
    cell(g,f"A{r}",i,F(11,True,BLUE),Fill(ZEBRA),Al("center"),border=box)
    merge(g,f"B{r}:F{r}",p,font=F(10.5),al=Al("left",wrap=True));
    for cc in "BCDEF": g[f"{cc}{r}"].border=box
    g.row_dimensions[r].height=28; r+=1

r+=1
band(g,f"A{r}:F{r}","LƯU Ý ĐỒNG BỘ (QUAN TRỌNG)")
r+=1
merge(g,f"A{r}:F{r}","Lần đầu mở file: nếu thấy ô báo #REF! hoặc \"CHUA_KET_NOI\", bấm vào ô đó → chọn \"Cho phép truy cập / Allow access\" để kết nối tới file làm việc. Chỉ cần làm 1 lần.",
      font=F(10.5,color=RED),fill=Fill(RED_BG),al=Al("left",wrap=True))
g.row_dimensions[r].height=40; r+=2

band(g,f"A{r}:F{r}","LIÊN KẾT & TÌNH TRẠNG")
r+=1
cell(g,f"A{r}","File làm việc (nguồn)",F(11,True),Fill(LBLUE),border=box)
merge(g,f"B{r}:F{r}",WORK_URL,font=F(10.5,color="1155CC"),al=Al("left"));
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=1
cell(g,f"A{r}","① Kích hoạt kết nối",F(11,True,RED),Fill(AMBER_BG),Al("left",wrap=True),border=box)
merge(g,f"B{r}:F{r}",f'=IMPORTRANGE({U()},"00A_Tong_hop_Trang_Tinh_nang!T5")',
      font=F(11,True),fill=Fill(AMBER_BG),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
g[f"B{r}"].number_format="yyyy-mm-dd"
g_activate_row=r
r+=1
merge(g,f"A{r}:F{r}","→ Nếu ô ① báo #REF!: bấm vào ô đó, chọn \"Cho phép truy cập / Allow access\". Sau đó toàn bộ báo cáo tự có số. Chỉ làm 1 lần.",
      font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left",wrap=True))
g.row_dimensions[r].height=28
r+=1
cell(g,f"A{r}","Trạng thái kết nối",F(11,True),Fill(LBLUE),border=box)
merge(g,f"B{r}:F{r}",f'=IF(ISERROR({T_TOTAL}),"⚠ Chưa kết nối — hãy Cho phép truy cập ở ô ①","✔ Đã kết nối — báo cáo đang cập nhật live")',
      font=F(10.5,True,GREEN),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box

# =====================================================================
#  SHEET: 01_Báo_cáo_Tháng
# =====================================================================
m = wb.create_sheet("01_Báo_cáo_Tháng")
widths={"A":34,"B":15,"C":13,"D":15,"E":13,"F":30,"G":3,"H":16,"I":10}
for col,w in widths.items(): m.column_dimensions[col].width=w
# hidden helper cols for date bounds
m.column_dimensions["N"].hidden=True; m.column_dimensions["O"].hidden=True

band(m,"A1:F1","BÁO CÁO TIẾN ĐỘ THEO THÁNG",big=True)
merge(m,"A2:F2","Chọn tháng ở ô B4. Số liệu tổng thể là thời điểm hiện tại; chỉ số trong tháng lọc theo tháng đã chọn.",
      font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
# selector row
cell(m,"A4","Tháng báo cáo",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"B4",8,F(12,True,NAVY),Fill(AMBER_BG),Al("center"),border=box)
cell(m,"C4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"D4",2026,F(12,True,NAVY),Fill(AMBER_BG),Al("center"),border=box)
cell(m,"E4","Ngày chốt",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"F4","=TODAY()",F(11,True),Fill(AMBER_BG),Al("center"),border=box); m["F4"].number_format="yyyy-mm-dd"
# date bounds helpers
m["N4"]="=DATE(D4,B4,1)"; m["O4"]="=EOMONTH(N4,0)"
mF="$N$4"; mT="$O$4"

dv_m = DataValidation(type="list", formula1='"1,2,3,4,5,6,7,8,9,10,11,12"', allow_blank=False); m.add_data_validation(dv_m); dv_m.add(m["B4"])

def kpi_header(ws, r):
    hdrs=["KPI / CHỈ SỐ","Mục tiêu","Kết quả","Tỷ lệ / %","Đánh giá","Ghi chú"]
    for i,h in enumerate(hdrs):
        c=get_column_letter(1+i)
        cell(ws,f"{c}{r}",h,F(10.5,True,WHITE),Fill(NAVY),Al("center" if i else "left"),border=box)
    ws.row_dimensions[r].height=22

def section(ws,r,txt):
    merge(ws,f"A{r}:F{r}",txt,font=F(11,True,NAVY),fill=Fill(LBLUE),al=Al("left"))
    ws.row_dimensions[r].height=20

def kpi_row(ws,r,name,target,result,ratio,verdict,note,rfmt=INT,zebra=False):
    bg = Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(ws,f"A{r}",name,F(10.5),bg,Al("left",wrap=True),border=box)
    cell(ws,f"B{r}",target,F(10.5),bg,Al("center"),border=box)
    cell(ws,f"C{r}",result,F(11,True),bg,Al("center"),border=box, nfmt=rfmt)
    cell(ws,f"D{r}",ratio,F(10.5),bg,Al("center"),border=box,nfmt=PCT if ratio is not None else None)
    cell(ws,f"E{r}",verdict,F(10.5,True),bg,Al("center"),border=box)
    cell(ws,f"F{r}",note,F(9.5,color=GREY),bg,Al("left",wrap=True),border=box)

r=6
kpi_header(m,r); r+=1
section(m,r,"I. KIỂM TRA WEBSITE"); r+=1
kpi_row(m,r,"1. Trang hoàn thành","100%",f"={T_DONE}",f"={T_PROG}",
        f'=IF({T_PROG}>=1,"Đạt","Đang làm")',"Trên tổng 56 trang",INT); r+=1
kpi_row(m,r,"2. Tính năng hoàn thành","100%",f"={F_DONE}",f"={F_PROG}",
        f'=IF({F_PROG}>=1,"Đạt","Đang làm")',"Trên tổng 88 tính năng",INT,zebra=True); r+=1
kpi_row(m,r,"3. Công việc hoàn thành trong tháng","—",
        f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{mF},\'z_CV\'!$B:$B,"<="&{mT},\'z_CV\'!$O:$O,"Hoàn thành")',
        None,"Theo dõi","Nhật ký công việc theo ngày",INT); r+=1
kpi_row(m,r,"4. Lỗi phát hiện trong tháng","—",
        f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{mF},\'z_Bug\'!$B:$B,"<="&{mT})',
        None,"Theo dõi","Theo Ngày phát hiện lỗi",INT,zebra=True); r+=1
section(m,r,"II. HOÀN THIỆN WEBSITE"); r+=1
kpi_row(m,r,"1. Lỗi đã xử lý (Dev báo fix) trong tháng","—",
        f'=COUNTIFS(\'z_Bug\'!$W:$W,">="&{mF},\'z_Bug\'!$W:$W,"<="&{mT})',
        None,"Theo dõi","Theo Ngày Dev báo fix",INT); r+=1
r+=1

# CURRENT STATUS block
band(m,f"A{r}:F{r}","TỔNG QUAN HIỆN TẠI (thời điểm xem)"); r+=1
hdr=["Chỉ số","Số lượng","Tỷ lệ"]
cell(m,f"A{r}",hdr[0],F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"B{r}",hdr[1],F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
merge(m,f"C{r}:D{r}",hdr[2],font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("center"))
for cc in "CD": m[f"{cc}{r}"].border=box
r+=1
def stat(ws,r,label,val,ratio=None,zebra=False,strong=None):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(ws,f"A{r}",label,F(10.5),bg,Al("left"),border=box)
    cell(ws,f"B{r}",val,F(11,True,strong or "1A1A1A"),bg,Al("center"),border=box,nfmt=INT)
    if ratio is not None:
        merge(ws,f"C{r}:D{r}",ratio,font=F(10.5,True),fill=bg,al=Al("center"));
    else:
        merge(ws,f"C{r}:D{r}","",fill=bg)
    for cc in "CD": ws[f"{cc}{r}"].border=box
stat(m,r,"Trang hoàn thành",f"={T_DONE}",f"={T_PROG}"); m[f"C{r}"].number_format=PCT; r+=1
stat(m,r,"Trang cần kiểm tra lại",f"={T_RETEST}",f"={T_RETEST}/{T_TOTAL}",zebra=True); m[f"C{r}"].number_format=PCT; r+=1
stat(m,r,"Trang chưa kiểm tra",f"={T_UNTEST}",f"={T_UNTEST}/{T_TOTAL}"); m[f"C{r}"].number_format=PCT; r+=1
stat(m,r,"Tính năng hoàn thành",f"={F_DONE}",f"={F_PROG}",zebra=True); m[f"C{r}"].number_format=PCT; r+=1
stat(m,r,"Tính năng cần kiểm tra lại",f"={F_RETEST}",f"={F_RETEST}/{F_TOTAL}"); m[f"C{r}"].number_format=PCT; r+=1
stat(m,r,"Tổng lỗi ghi nhận",f"={BUG_TOTAL}",None,zebra=True); r+=1
stat(m,r,"Lỗi đang mở",f"={BUG_OPEN}",None,strong=RED); r+=1
stat(m,r,"Lỗi Critical/High đang mở",f"={SEV_C}+{SEV_H}",None,zebra=True,strong=RED); r+=1
r+=1

# BUG DISTRIBUTION
band(m,f"A{r}:F{r}","PHÂN BỐ LỖI ĐANG MỞ"); r+=1
cell(m,f"A{r}","Theo mức độ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"B{r}","Số lỗi mở",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,f"D{r}","Theo thiết bị (tổng lỗi)",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"E{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
r0=r; r+=1
sev=[("Critical",f"={SEV_C}",RED),("High",f"={SEV_H}",RED),("Medium",f"={SEV_M}",AMBER),("Low",f"={SEV_L}",GREEN)]
dev=[("Desktop","Desktop"),("Mobile","Mobile"),("Tablet","Tablet")]
for i,(nm,fx,clr) in enumerate(sev):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(m,f"A{r}",nm,F(10.5,True,clr),bg,Al("left"),border=box)
    cell(m,f"B{r}",fx,F(11,True),bg,Al("center"),border=box,nfmt=INT)
    r+=1
r=r0+1
for i,(nm,key) in enumerate(dev):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(m,f"D{r}",nm,F(10.5),bg,Al("left"),border=box)
    cell(m,f"E{r}",f'=COUNTIF(\'z_Bug\'!$M:$M,"*{key}*")',F(11,True),bg,Al("center"),border=box,nfmt=INT)
    r+=1
r=r0+5
# progress bars
merge(m,f"A{r}:B{r}","THANH TIẾN ĐỘ",font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("left"))
for cc in "AB": m[f"{cc}{r}"].border=box
merge(m,f"C{r}:F{r}","",fill=Fill(BLUE))
r+=1
cell(m,f"A{r}","Trang",F(10.5,True),Fill(WHITE),Al("left"),border=box)
merge(m,f"B{r}:F{r}",f'=REPT("█",ROUND({T_PROG}*30,0))&" "&TEXT({T_PROG},"0.0%")',
      font=F(11,True,BLUE),fill=Fill(WHITE),al=Al("left"))
for cc in "BCDEF": m[f"{cc}{r}"].border=box
r+=1
cell(m,f"A{r}","Tính năng",F(10.5,True),Fill(ZEBRA),Al("left"),border=box)
merge(m,f"B{r}:F{r}",f'=REPT("█",ROUND({F_PROG}*30,0))&" "&TEXT({F_PROG},"0.0%")',
      font=F(11,True,GREEN),fill=Fill(ZEBRA),al=Al("left"))
for cc in "BCDEF": m[f"{cc}{r}"].border=box

# =====================================================================
#  SHEET: 02_Báo_cáo_Tuần
# =====================================================================
w = wb.create_sheet("02_Báo_cáo_Tuần")
for col,wd in {"A":34,"B":16,"C":14,"D":14,"E":14,"F":28}.items(): w.column_dimensions[col].width=wd
w.column_dimensions["N"].hidden=True; w.column_dimensions["O"].hidden=True
band(w,"A1:F1","BÁO CÁO TIẾN ĐỘ THEO TUẦN",big=True)
merge(w,"A2:F2","Tuần tính theo tuần-trong-tháng: Tuần 1 = ngày 1–7, Tuần 2 = 8–14, ... (khớp cách chia của công ty).",
      font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
cell(w,"A4","Tháng",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(w,"B4",8,F(12,True,NAVY),Fill(AMBER_BG),Al("center"),border=box)
cell(w,"C4","Tuần",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(w,"D4",2,F(12,True,NAVY),Fill(AMBER_BG),Al("center"),border=box)
cell(w,"E4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(w,"F4",2026,F(12,True,NAVY),Fill(AMBER_BG),Al("center"),border=box)
dv_wm=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); w.add_data_validation(dv_wm); dv_wm.add(w["B4"])
dv_ww=DataValidation(type="list",formula1='"1,2,3,4,5"'); w.add_data_validation(dv_ww); dv_ww.add(w["D4"])
# from/to for week-of-month
w["N4"]="=DATE(F4,B4,(D4-1)*7+1)"
w["O4"]="=MIN(EOMONTH(DATE(F4,B4,1),0),DATE(F4,B4,D4*7))"
wF="$N$4"; wT="$O$4"
cell(w,"A5","Từ ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box)
cell(w,"B5","=N4",F(10.5,True),Fill(WHITE),Al("center"),border=box); w["B5"].number_format="yyyy-mm-dd"
cell(w,"C5","Đến ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box)
cell(w,"D5","=O4",F(10.5,True),Fill(WHITE),Al("center"),border=box); w["D5"].number_format="yyyy-mm-dd"
merge(w,"E5:F5",f'=IF(ISERROR({T_TOTAL}),"⚠ Chưa kết nối","✔ Đã kết nối")',font=F(10.5,True,GREEN),fill=Fill(WHITE),al=Al("center"))

r=7
band(w,f"A{r}:F{r}","CHỈ SỐ TRONG TUẦN"); r+=1
cell(w,f"A{r}","Chỉ số",F(10.5,True,WHITE),Fill(NAVY),Al("left"),border=box)
cell(w,f"B{r}","Kết quả",F(10.5,True,WHITE),Fill(NAVY),Al("center"),border=box)
merge(w,f"C{r}:F{r}","Ghi chú",font=F(10.5,True,WHITE),fill=Fill(NAVY),al=Al("left"))
for cc in "CDEF": w[f"{cc}{r}"].border=box
r+=1
def wrow(r,label,formula,note,strong=None,zebra=False,indent=False):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(w,f"A{r}",("    "+label) if indent else label,F(10.5),bg,Al("left"),border=box)
    cell(w,f"B{r}",formula,F(11,True,strong or "1A1A1A"),bg,Al("center"),border=box,nfmt=INT)
    merge(w,f"C{r}:F{r}",note,font=F(9.5,color=GREY),fill=bg,al=Al("left",wrap=True))
    for cc in "CDEF": w[f"{cc}{r}"].border=box
wrow(r,"Công việc hoàn thành",f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{wF},\'z_CV\'!$B:$B,"<="&{wT},\'z_CV\'!$O:$O,"Hoàn thành")',"Nhật ký công việc theo ngày"); r+=1
wrow(r,"Công việc đang tiến hành",f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{wF},\'z_CV\'!$B:$B,"<="&{wT},\'z_CV\'!$O:$O,"Đang làm")',"",zebra=True); r+=1
wrow(r,"Lỗi phát hiện",f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{wF},\'z_Bug\'!$B:$B,"<="&{wT})',"Theo Ngày phát hiện"); r+=1
wrow(r,"Critical",f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{wF},\'z_Bug\'!$B:$B,"<="&{wT},\'z_Bug\'!$L:$L,"Critical")',"",strong=RED,zebra=True,indent=True); r+=1
wrow(r,"High",f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{wF},\'z_Bug\'!$B:$B,"<="&{wT},\'z_Bug\'!$L:$L,"High")',"",strong=RED,indent=True); r+=1
wrow(r,"Medium",f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{wF},\'z_Bug\'!$B:$B,"<="&{wT},\'z_Bug\'!$L:$L,"Medium")',"",strong=AMBER,zebra=True,indent=True); r+=1
wrow(r,"Low",f'=COUNTIFS(\'z_Bug\'!$B:$B,">="&{wF},\'z_Bug\'!$B:$B,"<="&{wT},\'z_Bug\'!$L:$L,"Low")',"",strong=GREEN,indent=True); r+=1
wrow(r,"Lỗi Dev báo fix",f'=COUNTIFS(\'z_Bug\'!$W:$W,">="&{wF},\'z_Bug\'!$W:$W,"<="&{wT})',"Theo Ngày Dev báo fix",zebra=True); r+=1
wrow(r,"Lỗi còn mở hiện tại",f"={BUG_OPEN}","Tồn tại tại thời điểm xem",strong=RED); r+=1
r+=1
band(w,f"A{r}:F{r}","TIẾN ĐỘ TOÀN WEBSITE (thời điểm xem)"); r+=1
cell(w,f"A{r}","Chỉ số",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(w,f"B{r}","Số lượng",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
merge(w,f"C{r}:F{r}","Tỷ lệ",font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("center"))
for cc in "CDEF": w[f"{cc}{r}"].border=box
r+=1
def wstat(r,label,val,ratio,zebra=False,strong=None):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(w,f"A{r}",label,F(10.5),bg,Al("left"),border=box)
    cell(w,f"B{r}",val,F(11,True,strong or "1A1A1A"),bg,Al("center"),border=box,nfmt=INT)
    merge(w,f"C{r}:F{r}",ratio,font=F(10.5,True),fill=bg,al=Al("center")); w[f"C{r}"].number_format=PCT
    for cc in "CDEF": w[f"{cc}{r}"].border=box
wstat(r,"Trang hoàn thành",f"={T_DONE}",f"={T_PROG}"); r+=1
wstat(r,"Tính năng hoàn thành",f"={F_DONE}",f"={F_PROG}",zebra=True); r+=1
wstat(r,"Trang cần kiểm tra lại",f"={T_RETEST}",f"={T_RETEST}/{T_TOTAL}"); r+=1
wstat(r,"Tính năng cần kiểm tra lại",f"={F_RETEST}",f"={F_RETEST}/{F_TOTAL}",zebra=True); r+=1

# =====================================================================
#  SHEET: 03_Nghiệm_thu
# =====================================================================
n = wb.create_sheet("03_Nghiệm_thu")
for col,wd in {"A":20,"B":24,"C":10,"D":12,"E":10,"F":16,"G":18,"H":26}.items(): n.column_dimensions[col].width=wd
band(n,"A1:H1","BIÊN BẢN / CHECKLIST NGHIỆM THU",big=True)
merge(n,"A2:H2","Dùng để chốt nghiệm thu theo kỳ. Cột Kết luận & Người duyệt do người duyệt điền.",
      font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
cell(n,"A4","Kỳ nghiệm thu",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(n,"B4","Tháng 8 - Tuần 2",F(10.5,True),Fill(AMBER_BG),Al("center"),border=box)
cell(n,"C4","Ngày lập",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
merge(n,"D4:E4","=TODAY()",font=F(10.5,True),fill=Fill(AMBER_BG),al=Al("center")); n["D4"].number_format="yyyy-mm-dd"
cell(n,"F4","Người lập",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(n,"G4","",F(10.5),Fill(AMBER_BG),Al("center"),border=box)
cell(n,"H4","",Fill(WHITE))
r=6
heads=["Hạng mục","Phạm vi","Tổng","Hoàn thành","Tỷ lệ","Critical/High mở","Kết luận","Ghi chú"]
for i,h in enumerate(heads):
    c=get_column_letter(1+i)
    cell(n,f"{c}{r}",h,F(10,True,WHITE),Fill(NAVY),Al("center"),border=box)
n.row_dimensions[r].height=30
r+=1
items=[
 ("UI/UX","56 trang",f"={T_TOTAL}",f"={T_DONE}",f"={T_PROG}",f"={SEV_C}+{SEV_H}"),
 ("Chức năng","88 tính năng",f"={F_TOTAL}",f"={F_DONE}",f"={F_PROG}",f"={SEV_C}+{SEV_H}"),
 ("Responsive","YC_87 + lỗi responsive","","","",f'=COUNTIF(\'z_Bug\'!$K:$K,"*Responsive*")'),
 ("Bảo mật","Test Case Security","","","",f'=COUNTIF(\'z_Bug\'!$K:$K,"*Security*")'),
 ("Dữ liệu","Test Case API/Database","","","",""),
]
for i,(cat,scope,tot,done,ratio,ch) in enumerate(items):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(n,f"A{r}",cat,F(10.5,True),bg,Al("left"),border=box)
    cell(n,f"B{r}",scope,F(10),bg,Al("left",wrap=True),border=box)
    cell(n,f"C{r}",tot,F(10.5),bg,Al("center"),border=box,nfmt=INT)
    cell(n,f"D{r}",done,F(10.5,True),bg,Al("center"),border=box,nfmt=INT)
    cell(n,f"E{r}",ratio,F(10.5),bg,Al("center"),border=box,nfmt=PCT if ratio else None)
    cell(n,f"F{r}",ch,F(10.5,True,RED),bg,Al("center"),border=box,nfmt=INT if ch else None)
    cell(n,f"G{r}","Chưa nghiệm thu",F(10,italic=True,color=GREY),bg,Al("center"),border=box)
    cell(n,f"H{r}","",F(10),bg,Al("left"),border=box)
    n.row_dimensions[r].height=26; r+=1
r+=1
merge(n,f"A{r}:B{r}","KẾT LUẬN CHUNG",font=F(11,True,WHITE),fill=Fill(BLUE),al=Al("left"))
merge(n,f"C{r}:H{r}","Chưa nghiệm thu",font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
for cc in "ABCDEFGH": n[f"{cc}{r}"].border=box
r+=1
merge(n,f"A{r}:H{r}","Điều kiện khuyến nghị: không còn Critical/High chưa được phê duyệt; phạm vi và bằng chứng đầy đủ; ngoại lệ phải có người duyệt và lý do rõ ràng.",
      font=F(9.5,color=GREY),fill=Fill(ZEBRA),al=Al("left",wrap=True))
n.row_dimensions[r].height=36

# tab colors
for ws,clr in [(g,NAVY),(m,BLUE),(w,"2E7D32"),(n,"8B5E00")]:
    ws.sheet_properties.tabColor=clr
# freeze panes
g.sheet_view.showGridLines=False
for ws in [g,m,w,n]:
    ws.sheet_view.showGridLines=False

wb.active = wb.sheetnames.index("00_Hướng_dẫn")
out="/tmp/BaoCao_BGD_TimViec123.xlsx"
wb.save(out)
print("SAVED:", out)
print("Sheets:", wb.sheetnames)
