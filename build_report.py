# -*- coding: utf-8 -*-
"""Báo cáo tiến độ Website TimViec123 cho BGD — đồng bộ live qua IMPORTRANGE.
Nguồn lỗi: 05_Lỗi_Tester. Database: file 'Hồ sơ' riêng."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, range_boundaries

WORK_URL = "https://docs.google.com/spreadsheets/d/1J2A_OiWz7IeYsgafLDY3mIfZsG-xCNGhkNaSQkhicHg/edit"
DB_URL   = "https://docs.google.com/spreadsheets/d/1i7-xLYiLHg60EqmCpfeA9B4Y0T-8YsgXk-_0qJ4eACo/edit"

# ---------- Palette ----------
NAVY="1F3A5F"; BLUE="2E5A88"; LBLUE="DCE6F1"; CARD="EEF3F9"; ZEBRA="F5F7FA"
GREEN="2E7D32"; GREEN_BG="D9F0DD"; AMBER="B8860B"; AMBER_BG="FFF3CD"
RED="C62828"; RED_BG="FBE3E3"; GREY="6B7280"; WHITE="FFFFFF"; INPUT="FFF9DB"; BORDER_CLR="C9D6E5"
thin=Side(style="thin",color=BORDER_CLR); box=Border(left=thin,right=thin,top=thin,bottom=thin)
def F(size=11,bold=False,color="1A1A1A",italic=False): return Font(name="Arial",size=size,bold=bold,color=color,italic=italic)
def Fill(c): return PatternFill("solid",fgColor=c)
def Al(h="left",v="center",wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
PCT='0.0%'; INT='#,##0'

def cell(ws,ref,val=None,font=None,fill=None,al=None,nfmt=None,border=None):
    c=ws[ref]
    if val is not None: c.value=val
    if font: c.font=font
    if fill: c.fill=fill
    if al: c.alignment=al
    if nfmt: c.number_format=nfmt
    if border: c.border=border
    return c
def fillrange(ws,rng,fill):
    c1,r1,c2,r2=range_boundaries(rng)
    for r in range(r1,r2+1):
        for c in range(c1,c2+1): ws.cell(r,c).fill=fill
def merge(ws,rng,val=None,font=None,fill=None,al=None,border=False):
    ws.merge_cells(rng); tl=rng.split(":")[0]
    cell(ws,tl,val,font,fill,al)
    if fill: fillrange(ws,rng,fill)
    if border:
        c1,r1,c2,r2=range_boundaries(rng)
        for r in range(r1,r2+1):
            for c in range(c1,c2+1): ws.cell(r,c).border=box
def band(ws,rng,text,big=False):
    merge(ws,rng,text,font=F(15 if big else 11.5,True,WHITE),fill=Fill(NAVY if big else BLUE),al=Al("left"))
    r=int(''.join(ch for ch in rng.split(":")[0] if ch.isdigit()))
    ws.row_dimensions[r].height=30 if big else 22

wb=openpyxl.Workbook()

# ============ HELPER SHEETS ============
def UW(): return "'z_Config'!$B$1"   # working url
def UD(): return "'z_Config'!$B$3"   # database url
zc=wb.active; zc.title="z_Config"
zc["A1"]="URL file làm việc"; zc["B1"]=WORK_URL
zc["A2"]="Năm báo cáo";       zc["B2"]=2026
zc["A3"]="URL file Database";  zc["B3"]=DB_URL
for r in (1,2,3): zc[f"A{r}"].font=F(bold=True)
zc.column_dimensions["A"].width=20; zc.column_dimensions["B"].width=70

def src(name, url_ref, rng):
    ws=wb.create_sheet(name); ws["A1"]=f'=IMPORTRANGE({url_ref},"{rng}")'; ws.sheet_state="hidden"; return ws
# working-file sources
src("z_Trang",    UW(), "02_UI_UX!A3:N3")
src("z_TinhNang", UW(), "03_Tính_năng!A3:N3")
src("z_Loi",      UW(), "05_Lỗi_Tester!A4:V606")     # bug list (header row1 + data)
src("z_CV",       UW(), "01_Cong_viec_ngay!A4:P1004")# work log
# database sources (data only, no header)
src("z_DB_DN",    UD(), "'HS DN'!D4:N970")            # A=Tên DN ... K=Ngày Đăng/Tạo(N)
src("z_DB_HR",    UD(), "'HS HR'!D4:E966")            # A=Tên HR, B=Ngày Đăng/Tạo(E)
src("z_DB_UV",    UD(), "'HS ỨNG TUYỂN'!D5:D990")     # A=Họ tên (không có cột ngày)
src("z_DB_Tin",   UD(), "'Tin tuyển dụng'!C4:D1006")  # A=Nội dung, B=Ngày viết(D)
zc.sheet_state="hidden"

# ---- live refs ----
T_TOT="'z_Trang'!$B$1"; T_DONE="'z_Trang'!$D$1"; T_PROG="'z_Trang'!$F$1"; T_ING="'z_Trang'!$H$1"; T_RE="'z_Trang'!$J$1"; T_UN="'z_Trang'!$L$1"
F_TOT="'z_TinhNang'!$B$1"; F_DONE="'z_TinhNang'!$D$1"; F_PROG="'z_TinhNang'!$F$1"; F_ING="'z_TinhNang'!$H$1"; F_RE="'z_TinhNang'!$J$1"; F_UN="'z_TinhNang'!$L$1"
U="'z_Loi'!$U:$U"; L="'z_Loi'!$L:$L"; M="'z_Loi'!$M:$M"; Q="'z_Loi'!$Q:$Q"; E="'z_Loi'!$E:$E"
def cif(rng,crit): return f'COUNTIF({rng},"{crit}")'
BUG_TOTAL=f'=COUNTA({E})-1'
OPEN=f'=({cif(U,"Open")}+{cif(U,"Chưa gửi Dev")})'
RESOLVED=f'=({cif(U,"Fixed")}+{cif(U,"Verified")}+{cif(U,"Closed")})'
CH_OPEN=(f'=(COUNTIFS({U},"Open",{L},"Critical")+COUNTIFS({U},"Open",{L},"High")'
         f'+COUNTIFS({U},"Chưa gửi Dev",{L},"Critical")+COUNTIFS({U},"Chưa gửi Dev",{L},"High"))')
def SEV(x): return f'=COUNTIF({L},"{x}")'
def DEV(x): return f'=COUNTIF({M},"*{x}*")'
# month bounds live at 01_Báo_cáo_Tháng!$N$4 / $O$4
MF="'01_Báo_cáo_Tháng'!$N$4"; MT="'01_Báo_cáo_Tháng'!$O$4"
def found_month(): return f'=COUNTIFS({Q},">="&{MF},{Q},"<="&{MT})'
def sev_month(x): return f'=COUNTIFS({Q},">="&{MF},{Q},"<="&{MT},{L},"{x}")'
def dev_month(x): return f'=COUNTIFS({Q},">="&{MF},{Q},"<="&{MT},{M},"*{x}*")'
def cv_done_month(): return f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{MF},\'z_CV\'!$B:$B,"<="&{MT},\'z_CV\'!$O:$O,"Hoàn thành")'
# DB
DBDN="'z_DB_DN'!$A:$A"; DBDN_D="'z_DB_DN'!$K:$K"
DBHR="'z_DB_HR'!$A:$A"; DBHR_D="'z_DB_HR'!$B:$B"
DBUV="'z_DB_UV'!$A:$A"
DBTIN="'z_DB_Tin'!$A:$A"; DBTIN_D="'z_DB_Tin'!$B:$B"
def db_total(a): return f'=COUNTA({a})'
def db_month(d): return f'=COUNTIFS({d},">="&{MF},{d},"<="&{MT})'

# ============ 00_Hướng_dẫn ============
g=wb.create_sheet("00_Hướng_dẫn")
for col,w in {"A":26,"B":30,"C":28,"D":20,"E":20,"F":16}.items(): g.column_dimensions[col].width=w
band(g,"A1:F1","BÁO CÁO TIẾN ĐỘ WEBSITE TIMVIEC123 — DÀNH CHO BAN GIÁM ĐỐC",big=True)
merge(g,"A2:F2","Số liệu đồng bộ live từ file làm việc QA và file Database. BGD chỉ xem, chọn kỳ ở sheet Tháng/Tuần.",
      font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left")); g.row_dimensions[2].height=20
merge(g,"A4:B4","NGƯỜI DÙNG",font=F(11,True,WHITE),fill=Fill(BLUE)); merge(g,"C4:F4","CÁCH SỬ DỤNG",font=F(11,True,WHITE),fill=Fill(BLUE))
r=5
for who,how in [("Ban Giám đốc","Xem tiến độ trang/tính năng, tình trạng lỗi, database và nghiệm thu. Chọn Tháng/Tuần để lọc kỳ."),
                ("Tester","Cập nhật dữ liệu tại file làm việc; báo cáo tự đồng bộ."),
                ("Dev","Cập nhật xử lý lỗi tại 06_RTM_Dev_Test của file làm việc.")]:
    merge(g,f"A{r}:B{r}",who,font=F(11,True),fill=Fill(LBLUE),al=Al("left")); merge(g,f"C{r}:F{r}",how,font=F(10.5),al=Al("left",wrap=True)); g.row_dimensions[r].height=30; r+=1
r+=1; band(g,f"A{r}:F{r}","NGUYÊN TẮC TÍNH SỐ"); r+=1
for i,p in enumerate([
 "Tiến độ trang = số trang Hoàn thành / tổng số trang (hiện 63).",
 "Tiến độ tính năng = số tính năng Hoàn thành / 88 tính năng (gồm YC_88 chờ flow).",
 "Số lỗi lấy toàn bộ từ sheet 05_Lỗi_Tester. \"Đang mở\" = trạng thái Open + Chưa gửi Dev (chưa xử lý xong).",
 "\"Đã xử lý\" = Fixed + Verified + Closed. Riêng Deferred (tạm hoãn) tính riêng.",
 "Số liệu Database (DN/HR/Ứng viên/Tin TD) lấy từ file Hồ sơ, đếm theo cột ngày.",
 "Chỉ số \"trong tháng\" lọc theo Ngày phát hiện lỗi và Ngày đăng/tạo của kỳ đã chọn."],1):
    cell(g,f"A{r}",i,F(11,True,BLUE),Fill(ZEBRA),Al("center"),border=box); merge(g,f"B{r}:F{r}",p,font=F(10.5),al=Al("left",wrap=True))
    for cc in "BCDEF": g[f"{cc}{r}"].border=box
    g.row_dimensions[r].height=26; r+=1
r+=1; band(g,f"A{r}:F{r}","KÍCH HOẠT KẾT NỐI (làm 1 lần)")
r+=1
cell(g,f"A{r}","① File làm việc QA",F(11,True,RED),Fill(AMBER_BG),Al("left"),border=box)
merge(g,f"B{r}:F{r}",f'=IMPORTRANGE({UW()},"05_Lỗi_Tester!E4")',font=F(11,True),fill=Fill(AMBER_BG),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=1
cell(g,f"A{r}","② File Database",F(11,True,RED),Fill(AMBER_BG),Al("left"),border=box)
merge(g,f"B{r}:F{r}",f'=IMPORTRANGE({UD()},"\'HS DN\'!D3")',font=F(11,True),fill=Fill(AMBER_BG),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=1
merge(g,f"A{r}:F{r}","→ Nếu ô ① hoặc ② báo #REF!: bấm vào ô đó → \"Cho phép truy cập / Allow access\". Làm 1 lần cho mỗi ô.",
      font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left",wrap=True)); g.row_dimensions[r].height=28
r+=1
cell(g,f"A{r}","Trạng thái kết nối",F(11,True),Fill(LBLUE),border=box)
merge(g,f"B{r}:F{r}",f'=IF(ISERROR({T_TOT}),"⚠ Chưa kết nối file làm việc","✔ Đã kết nối")&"   |   "&IF(ISERROR({DBDN}),"⚠ Chưa kết nối Database","✔ Database OK")',
      font=F(10.5,True,GREEN),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=2
band(g,f"A{r}:F{r}","LIÊN KẾT"); r+=1
for lbl,url in [("File làm việc (QA)",WORK_URL),("File Database (Hồ sơ)",DB_URL)]:
    cell(g,f"A{r}",lbl,F(11,True),Fill(LBLUE),border=box); merge(g,f"B{r}:F{r}",url,font=F(10,color="1155CC"),al=Al("left"))
    for cc in "BCDEF": g[f"{cc}{r}"].border=box
    r+=1

# ============ 01_Báo_cáo_Tháng ============
m=wb.create_sheet("01_Báo_cáo_Tháng")
for col,w in {"A":36,"B":14,"C":12,"D":12,"E":12,"F":30,"G":3,"H":18,"I":10}.items(): m.column_dimensions[col].width=w
m.column_dimensions["N"].hidden=True; m.column_dimensions["O"].hidden=True
band(m,"A1:F1","BÁO CÁO TIẾN ĐỘ THEO THÁNG",big=True)
merge(m,"A2:F2","Chọn tháng ở ô B4. Ô nền vàng là chỉ tiêu nhập tay. Số tổng quan là thời điểm hiện tại; chỉ số \"trong tháng\" lọc theo tháng đã chọn.",
      font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
cell(m,"A4","Tháng báo cáo",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
cell(m,"C4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"D4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
cell(m,"E4","Ngày chốt",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,"F4","=TODAY()",F(11,True),Fill(INPUT),Al("center"),border=box); m["F4"].number_format="dd/mm/yyyy"
m["N4"]="=DATE(D4,B4,1)"; m["O4"]="=EOMONTH(N4,0)"
dv=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); m.add_data_validation(dv); dv.add(m["B4"])

def kpi_header(r):
    for i,h in enumerate(["KPI / CHỈ SỐ","Mục tiêu tháng","Kết quả","Tỷ lệ","Đạt KPI?","Ghi chú"]):
        c=get_column_letter(1+i); cell(m,f"{c}{r}",h,F(10.5,True,WHITE),Fill(NAVY),Al("center" if i else "left"),border=box)
    m.row_dimensions[r].height=22
def sect(r,txt): merge(m,f"A{r}:F{r}",txt,font=F(11,True,NAVY),fill=Fill(LBLUE),al=Al("left"),border=True); m.row_dimensions[r].height=19
def kpi(r,name,target,result,ratio,note,zebra=False,rfmt=INT):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(m,f"A{r}",name,F(10.5),bg,Al("left",wrap=True),border=box)
    tfill=Fill(INPUT) if isinstance(target,(int,float)) else bg
    cell(m,f"B{r}",target,F(10.5,True),tfill,Al("center"),border=box, nfmt=INT if isinstance(target,(int,float)) else None)
    cell(m,f"C{r}",result,F(11,True),bg,Al("center"),border=box,nfmt=rfmt)
    cell(m,f"D{r}",ratio,F(10.5),bg,Al("center"),border=box,nfmt=PCT if ratio is not None else None)
    cell(m,f"E{r}",f'=IF(ISNUMBER(B{r}),IF(C{r}>=B{r},"Đạt","Chưa đạt"),"Theo dõi")',F(10.5,True),bg,Al("center"),border=box)
    cell(m,f"F{r}",note,F(9.5,color=GREY),bg,Al("left",wrap=True),border=box)

r=6; kpi_header(r); r+=1
sect(r,"I. KIỂM TRA WEBSITE"); r+=1
kpi(r,"1. Trang hoàn thành",63,f"={T_DONE}",f"={T_PROG}","Trên tổng số trang website"); r+=1
kpi(r,"2. Tính năng hoàn thành",88,f"={F_DONE}",f"={F_PROG}","Trên tổng 88 tính năng",zebra=True); r+=1
kpi(r,"3. Công việc hoàn thành trong tháng","—",cv_done_month(),None,"Nhật ký công việc theo ngày"); r+=1
kpi(r,"4. Lỗi phát hiện trong tháng","—",found_month(),None,"Theo Ngày phát hiện (05_Lỗi_Tester)",zebra=True); r+=1
sect(r,"II. HOÀN THIỆN WEBSITE"); r+=1
kpi(r,"1. Lỗi đã xử lý (lũy kế)","—",RESOLVED,None,"Fixed + Verified + Closed"); r+=1
kpi(r,"2. Số Trang đã xử lý","—",f"={T_TOT}-{T_UN}",None,"Đã kiểm thử = Tổng − Chưa kiểm tra",zebra=True); r+=1
kpi(r,"3. Số Tính năng đã xử lý","—",f"={F_TOT}-{F_UN}",None,"Đã kiểm thử = Tổng − Chưa kiểm tra"); r+=1
sect(r,"III. DATABASE (trong tháng)"); r+=1
kpi(r,"1. DB Doanh nghiệp đã thêm",100,db_month(DBDN_D),None,"Theo Ngày Đăng/Tạo"); r+=1
kpi(r,"2. DB HR đã thêm",100,db_month(DBHR_D),None,"Theo Ngày Đăng/Tạo",zebra=True); r+=1
kpi(r,"3. DB Ứng viên đã thêm",100,f"={db_total(DBUV)[1:]}",None,"Theo tổng (chưa có cột ngày nhập)"); r+=1
kpi(r,"4. Tin tuyển dụng đã đăng",50,db_month(DBTIN_D),None,"Theo Ngày viết",zebra=True); r+=1
r+=1

# TỔNG QUAN HIỆN TẠI
band(m,f"A{r}:F{r}","TỔNG QUAN HIỆN TẠI (thời điểm xem)"); r+=1
for i,h in enumerate(["Chỉ số","Số lượng","Tổng/Mục tiêu","Tỷ lệ"]):
    c=get_column_letter(1+i); cell(m,f"{c}{r}",h,F(10.5,True,WHITE),Fill(BLUE),Al("center" if i else "left"),border=box)
r+=1
def stat(r,label,val,total,ratio,hi=None,zebra=False):
    bg = Fill(GREEN_BG) if hi=="g" else Fill(RED_BG) if hi=="r" else (Fill(ZEBRA) if zebra else Fill(WHITE))
    fc = GREEN if hi=="g" else RED if hi=="r" else "1A1A1A"
    cell(m,f"A{r}",label,F(10.5,True if hi else False,fc),bg,Al("left"),border=box)
    cell(m,f"B{r}",val,F(12 if hi else 11,True,fc),bg,Al("center"),border=box,nfmt=INT)
    cell(m,f"C{r}",total,F(10.5,color=fc),bg,Al("center"),border=box,nfmt=INT if total is not None else None)
    cell(m,f"D{r}",ratio,F(11 if hi else 10.5,True,fc),bg,Al("center"),border=box,nfmt=PCT if ratio is not None else None)
sect(r,"I. KIỂM TRA WEBSITE"); r+=1
stat(r,"Trang hoàn thành",f"={T_DONE}",f"={T_TOT}",f"={T_PROG}",hi="g"); r+=1
stat(r,"Trang đang kiểm tra",f"={T_ING}",f"={T_TOT}",f"={T_ING}/{T_TOT}",zebra=True); r+=1
stat(r,"Trang cần kiểm tra lại",f"={T_RE}",f"={T_TOT}",f"={T_RE}/{T_TOT}"); r+=1
stat(r,"Trang chưa kiểm tra",f"={T_UN}",f"={T_TOT}",f"={T_UN}/{T_TOT}",zebra=True); r+=1
stat(r,"Tính năng hoàn thành",f"={F_DONE}",f"={F_TOT}",f"={F_PROG}",hi="g"); r+=1
stat(r,"Tính năng cần kiểm tra lại",f"={F_RE}",f"={F_TOT}",f"={F_RE}/{F_TOT}",zebra=True); r+=1
stat(r,"Tính năng chưa kiểm tra",f"={F_UN}",f"={F_TOT}",f"={F_UN}/{F_TOT}"); r+=1
stat(r,"Tổng lỗi ghi nhận",BUG_TOTAL,None,None,zebra=True); r+=1
stat(r,"Lỗi đang mở (Open + Chưa gửi Dev)",OPEN,None,None,hi="r"); r+=1
stat(r,"Lỗi Critical/High đang mở",CH_OPEN,None,None,hi="r"); r+=1
sect(r,"II. DATABASE (tổng đã thêm)"); r+=1
stat(r,"DB Doanh nghiệp",db_total(DBDN),None,None); r+=1
stat(r,"DB HR",db_total(DBHR),None,None,zebra=True); r+=1
stat(r,"DB Ứng viên",db_total(DBUV),None,None); r+=1
stat(r,"Tin tuyển dụng đã đăng",db_total(DBTIN),None,None,zebra=True); r+=1
r+=1

# TÌNH TRẠNG XỬ LÝ LỖI (status breakdown)
band(m,f"A{r}:F{r}","TÌNH TRẠNG XỬ LÝ LỖI (toàn bộ)"); r+=1
cell(m,f"A{r}","Trạng thái",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
merge(m,f"C{r}:F{r}","Ý nghĩa",font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("left"));
for cc in "CDEF": m[f"{cc}{r}"].border=box
r+=1
statuses=[("Open","Lỗi đang mở, đã ghi nhận",RED,False),("Chưa gửi Dev","Chưa chuyển Dev xử lý",AMBER,True),
          ("Fixed","Dev báo đã sửa",GREEN,False),("Verified","Đã retest đạt",GREEN,True),
          ("Closed","Đã đóng",GREY,False),("Deferred","Tạm hoãn / không sửa",GREY,True)]
for nm,mean,clr,zb in statuses:
    bg=Fill(ZEBRA) if zb else Fill(WHITE)
    cell(m,f"A{r}",nm,F(10.5,True,clr),bg,Al("left"),border=box)
    cell(m,f"B{r}",cif(U,nm).join(["=",""]),F(11,True),bg,Al("center"),border=box,nfmt=INT)
    merge(m,f"C{r}:F{r}",mean,font=F(9.5,color=GREY),fill=bg,al=Al("left"));
    for cc in "CDEF": m[f"{cc}{r}"].border=box
    r+=1
r+=1

# PHÂN BỐ LỖI TRONG THÁNG
band(m,f"A{r}:F{r}","PHÂN BỐ LỖI TRONG THÁNG"); r+=1
cell(m,f"A{r}","Theo mức độ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,f"D{r}","Theo thiết bị",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"E{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
r0=r; r+=1
sev=[("Critical",RED),("High",RED),("Medium",AMBER),("Low",GREEN)]
for i,(nm,clr) in enumerate(sev):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(m,f"A{r}",nm,F(10.5,True,clr),bg,Al("left"),border=box)
    cell(m,f"B{r}",sev_month(nm),F(11,True),bg,Al("center"),border=box,nfmt=INT); r+=1
rr=r0+1
for i,nm in enumerate(["Desktop","Mobile","Tablet"]):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(m,f"D{rr}",nm,F(10.5),bg,Al("left"),border=box)
    cell(m,f"E{rr}",dev_month(nm),F(11,True),bg,Al("center"),border=box,nfmt=INT); rr+=1
r=r0+5; r+=1

# THANH TIẾN ĐỘ
merge(m,f"A{r}:F{r}","THANH TIẾN ĐỘ HOÀN THÀNH",font=F(11,True,WHITE),fill=Fill(BLUE),al=Al("left"),border=True); r+=1
cell(m,f"A{r}","Trang",F(10.5,True),Fill(WHITE),Al("left"),border=box)
merge(m,f"B{r}:F{r}",f'=REPT("█",ROUND({T_PROG}*30,0))&" "&TEXT({T_PROG},"0.0%")',font=F(11,True,BLUE),fill=Fill(WHITE),al=Al("left"));
for cc in "BCDEF": m[f"{cc}{r}"].border=box
r+=1
cell(m,f"A{r}","Tính năng",F(10.5,True),Fill(ZEBRA),Al("left"),border=box)
merge(m,f"B{r}:F{r}",f'=REPT("█",ROUND({F_PROG}*30,0))&" "&TEXT({F_PROG},"0.0%")',font=F(11,True,GREEN),fill=Fill(ZEBRA),al=Al("left"))
for cc in "BCDEF": m[f"{cc}{r}"].border=box
r+=2

# BẢNG LỖI THEO TRANG / TÍNH NĂNG (QUERY, month-filtered)
band(m,f"A{r}:F{r}","CHI TIẾT LỖI TRONG THÁNG THEO TRANG & TÍNH NĂNG"); r+=1
hdrp=r
cell(m,f"A{r}","Trang đã kiểm thử trong tháng",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
cell(m,f"D{r}","Tính năng đã kiểm thử trong tháng",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
cell(m,f"E{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
r+=1
q_from='TEXT($N$4,"yyyy-mm-dd")'; q_to='TEXT($O$4,"yyyy-mm-dd")'
qp=(f'=IFERROR(QUERY(\'z_Loi\'!$A$1:$U$606,"select B, count(E) where E is not null '
    f'and Q >= date \'"&{q_from}&"\' and Q <= date \'"&{q_to}&"\' group by B order by count(E) desc label count(E) \'\', B \'\'",1),"(không có lỗi trong kỳ)")')
qf=(f'=IFERROR(QUERY(\'z_Loi\'!$A$1:$U$606,"select D, count(E) where E is not null '
    f'and Q >= date \'"&{q_from}&"\' and Q <= date \'"&{q_to}&"\' group by D order by count(E) desc label count(E) \'\', D \'\'",1),"(không có lỗi trong kỳ)")')
cell(m,f"A{r}",qp,F(10),Fill(WHITE),Al("left"))
cell(m,f"D{r}",qf,F(10),Fill(WHITE),Al("left"))
# light border frame for query zones
for rr2 in range(hdrp+1, hdrp+16):
    for cc in ("A","B"): m[f"{cc}{rr2}"].border=box
    for cc in ("D","E"): m[f"{cc}{rr2}"].border=box

# ============ 02_Báo_cáo_Tuần ============
w=wb.create_sheet("02_Báo_cáo_Tuần")
for col,wd in {"A":34,"B":16,"C":14,"D":14,"E":14,"F":28}.items(): w.column_dimensions[col].width=wd
w.column_dimensions["N"].hidden=True; w.column_dimensions["O"].hidden=True
band(w,"A1:F1","BÁO CÁO TIẾN ĐỘ THEO TUẦN",big=True)
merge(w,"A2:F2","Tuần theo tuần-trong-tháng: Tuần 1 = ngày 1–7, Tuần 2 = 8–14, ...",font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
cell(w,"A4","Tháng",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(w,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
cell(w,"C4","Tuần",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(w,"D4",2,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
cell(w,"E4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(w,"F4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
DataValidation
dv1=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); w.add_data_validation(dv1); dv1.add(w["B4"])
dv2=DataValidation(type="list",formula1='"1,2,3,4,5"'); w.add_data_validation(dv2); dv2.add(w["D4"])
w["N4"]="=DATE(F4,B4,(D4-1)*7+1)"; w["O4"]="=MIN(EOMONTH(DATE(F4,B4,1),0),DATE(F4,B4,D4*7))"
wF="$N$4"; wT="$O$4"
cell(w,"A5","Từ ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); cell(w,"B5","=N4",F(10.5,True),Fill(WHITE),Al("center"),border=box); w["B5"].number_format="dd/mm/yyyy"
cell(w,"C5","Đến ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); cell(w,"D5","=O4",F(10.5,True),Fill(WHITE),Al("center"),border=box); w["D5"].number_format="dd/mm/yyyy"
merge(w,"E5:F5",f'=IF(ISERROR({T_TOT}),"⚠ Chưa kết nối","✔ Đã kết nối")',font=F(10.5,True,GREEN),fill=Fill(WHITE),al=Al("center"))
QW="'z_Loi'!$Q:$Q"; LW="'z_Loi'!$L:$L"
r=7; band(w,f"A{r}:F{r}","CHỈ SỐ TRONG TUẦN"); r+=1
cell(w,f"A{r}","Chỉ số",F(10.5,True,WHITE),Fill(NAVY),Al("left"),border=box); cell(w,f"B{r}","Kết quả",F(10.5,True,WHITE),Fill(NAVY),Al("center"),border=box)
merge(w,f"C{r}:F{r}","Ghi chú",font=F(10.5,True,WHITE),fill=Fill(NAVY),al=Al("left"));
for cc in "CDEF": w[f"{cc}{r}"].border=box
r+=1
def wrow(r,label,formula,note,strong=None,zebra=False,indent=False):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(w,f"A{r}",("    "+label) if indent else label,F(10.5),bg,Al("left"),border=box)
    cell(w,f"B{r}",formula,F(11,True,strong or "1A1A1A"),bg,Al("center"),border=box,nfmt=INT)
    merge(w,f"C{r}:F{r}",note,font=F(9.5,color=GREY),fill=bg,al=Al("left",wrap=True));
    for cc in "CDEF": w[f"{cc}{r}"].border=box
wrow(r,"Công việc hoàn thành",f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{wF},\'z_CV\'!$B:$B,"<="&{wT},\'z_CV\'!$O:$O,"Hoàn thành")',"Nhật ký công việc theo ngày"); r+=1
wrow(r,"Công việc đang tiến hành",f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{wF},\'z_CV\'!$B:$B,"<="&{wT},\'z_CV\'!$O:$O,"Đang làm")',"",zebra=True); r+=1
wrow(r,"Lỗi phát hiện",f'=COUNTIFS({QW},">="&{wF},{QW},"<="&{wT})',"Theo Ngày phát hiện (05_Lỗi_Tester)"); r+=1
for sv,cl,zb in [("Critical",RED,True),("High",RED,False),("Medium",AMBER,True),("Low",GREEN,False)]:
    wrow(r,sv,f'=COUNTIFS({QW},">="&{wF},{QW},"<="&{wT},{LW},"{sv}")',"",strong=cl,zebra=zb,indent=True); r+=1
wrow(r,"Lỗi còn mở hiện tại",OPEN,"Open + Chưa gửi Dev (toàn bộ)",strong=RED,zebra=True); r+=1
r+=1
band(w,f"A{r}:F{r}","TIẾN ĐỘ TOÀN WEBSITE (thời điểm xem)"); r+=1
cell(w,f"A{r}","Chỉ số",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box); cell(w,f"B{r}","Số lượng",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
merge(w,f"C{r}:F{r}","Tỷ lệ",font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("center"));
for cc in "CDEF": w[f"{cc}{r}"].border=box
r+=1
def wstat(r,label,val,ratio,zebra=False):
    bg=Fill(ZEBRA) if zebra else Fill(WHITE)
    cell(w,f"A{r}",label,F(10.5),bg,Al("left"),border=box); cell(w,f"B{r}",val,F(11,True),bg,Al("center"),border=box,nfmt=INT)
    merge(w,f"C{r}:F{r}",ratio,font=F(10.5,True),fill=bg,al=Al("center")); w[f"C{r}"].number_format=PCT
    for cc in "CDEF": w[f"{cc}{r}"].border=box
wstat(r,"Trang hoàn thành",f"={T_DONE}",f"={T_PROG}"); r+=1
wstat(r,"Tính năng hoàn thành",f"={F_DONE}",f"={F_PROG}",zebra=True); r+=1
wstat(r,"Trang cần kiểm tra lại",f"={T_RE}",f"={T_RE}/{T_TOT}"); r+=1
wstat(r,"Tính năng cần kiểm tra lại",f"={F_RE}",f"={F_RE}/{F_TOT}",zebra=True); r+=1

# ============ 03_Nghiệm_thu ============
n=wb.create_sheet("03_Nghiệm_thu")
for col,wd in {"A":20,"B":24,"C":10,"D":12,"E":10,"F":16,"G":18,"H":24}.items(): n.column_dimensions[col].width=wd
band(n,"A1:H1","BIÊN BẢN / CHECKLIST NGHIỆM THU",big=True)
merge(n,"A2:H2","Cột Kết luận & Người duyệt do người duyệt điền.",font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
cell(n,"A4","Kỳ nghiệm thu",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(n,"B4","Tháng 8",F(10.5,True),Fill(INPUT),Al("center"),border=box)
cell(n,"C4","Ngày lập",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box); merge(n,"D4:E4","=TODAY()",font=F(10.5,True),fill=Fill(INPUT),al=Al("center")); n["D4"].number_format="dd/mm/yyyy"
cell(n,"F4","Người lập",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(n,"G4","",F(10.5),Fill(INPUT),Al("center"),border=box)
r=6
for i,h in enumerate(["Hạng mục","Phạm vi","Tổng","Hoàn thành","Tỷ lệ","Critical/High mở","Kết luận","Ghi chú"]):
    c=get_column_letter(1+i); cell(n,f"{c}{r}",h,F(10,True,WHITE),Fill(NAVY),Al("center"),border=box)
n.row_dimensions[r].height=30; r+=1
items=[("UI/UX","Trang website",f"={T_TOT}",f"={T_DONE}",f"={T_PROG}",CH_OPEN),
       ("Chức năng","88 tính năng",f"={F_TOT}",f"={F_DONE}",f"={F_PROG}",CH_OPEN),
       ("Responsive","YC_87 + lỗi responsive","","","",f'=COUNTIF(\'z_Loi\'!$K:$K,"*Responsive*")'),
       ("Bảo mật","Test Case Security","","","",f'=COUNTIF(\'z_Loi\'!$K:$K,"*Security*")'),
       ("Dữ liệu","Test Case API/Database","","","","")]
for i,(cat,scope,tot,done,ratio,ch) in enumerate(items):
    bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    cell(n,f"A{r}",cat,F(10.5,True),bg,Al("left"),border=box); cell(n,f"B{r}",scope,F(10),bg,Al("left",wrap=True),border=box)
    cell(n,f"C{r}",tot,F(10.5),bg,Al("center"),border=box,nfmt=INT); cell(n,f"D{r}",done,F(10.5,True),bg,Al("center"),border=box,nfmt=INT)
    cell(n,f"E{r}",ratio,F(10.5),bg,Al("center"),border=box,nfmt=PCT if ratio else None)
    cell(n,f"F{r}",ch,F(10.5,True,RED),bg,Al("center"),border=box,nfmt=INT if ch else None)
    cell(n,f"G{r}","Chưa nghiệm thu",F(10,italic=True,color=GREY),bg,Al("center"),border=box); cell(n,f"H{r}","",F(10),bg,Al("left"),border=box)
    n.row_dimensions[r].height=26; r+=1
r+=1
merge(n,f"A{r}:B{r}","KẾT LUẬN CHUNG",font=F(11,True,WHITE),fill=Fill(BLUE),al=Al("left"),border=True)
merge(n,f"C{r}:H{r}","Chưa nghiệm thu",font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"));
for cc in "CDEFGH": n[f"{cc}{r}"].border=box
r+=1
merge(n,f"A{r}:H{r}","Điều kiện khuyến nghị: không còn Critical/High chưa được phê duyệt; phạm vi và bằng chứng đầy đủ; ngoại lệ phải có người duyệt và lý do rõ ràng.",
      font=F(9.5,color=GREY),fill=Fill(ZEBRA),al=Al("left",wrap=True)); n.row_dimensions[r].height=34

# tab colors & gridlines
for ws,clr in [(g,NAVY),(m,BLUE),(w,"2E7D32"),(n,"8B5E00")]:
    ws.sheet_properties.tabColor=clr; ws.sheet_view.showGridLines=False
wb.active=wb.sheetnames.index("01_Báo_cáo_Tháng")
out="/tmp/BaoCao_BGD_TimViec123.xlsx"; wb.save(out)
print("SAVED",out); print("Sheets:",wb.sheetnames)
