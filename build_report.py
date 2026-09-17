# -*- coding: utf-8 -*-
"""Báo cáo tiến độ Website TimViec123 cho BGD — đồng bộ live qua IMPORTRANGE.
Nguồn lỗi: 05_Lỗi_Tester. Database: file 'Hồ sơ'. Tháng & Tuần dùng chung layout.
Mục tiêu lưu theo từng kỳ ở sheet 04_Mục_tiêu."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, range_boundaries

WORK_URL="https://docs.google.com/spreadsheets/d/1J2A_OiWz7IeYsgafLDY3mIfZsG-xCNGhkNaSQkhicHg/edit"
DB_URL  ="https://docs.google.com/spreadsheets/d/1i7-xLYiLHg60EqmCpfeA9B4Y0T-8YsgXk-_0qJ4eACo/edit"

NAVY="1F3A5F"; BLUE="2E5A88"; LBLUE="DCE6F1"; ZEBRA="F5F7FA"
GREEN="2E7D32"; GREEN_BG="D9F0DD"; AMBER="B8860B"; RED="C62828"; RED_BG="FBE3E3"
GREY="6B7280"; WHITE="FFFFFF"; INPUT="FFF9DB"; BORDER_CLR="C9D6E5"
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
    ws.merge_cells(rng); tl=rng.split(":")[0]; cell(ws,tl,val,font,fill,al)
    if fill: fillrange(ws,rng,fill)
    if border:
        c1,r1,c2,r2=range_boundaries(rng)
        for r in range(r1,r2+1):
            for c in range(c1,c2+1): ws.cell(r,c).border=box
def band(ws,rng,text,big=False):
    merge(ws,rng,text,font=F(15 if big else 11.5,True,WHITE),fill=Fill(NAVY if big else BLUE),al=Al("left"))
    r=int(''.join(ch for ch in rng.split(":")[0] if ch.isdigit())); ws.row_dimensions[r].height=30 if big else 22

wb=openpyxl.Workbook()
def UW(): return "'z_Config'!$B$1"
def UD(): return "'z_Config'!$B$3"
zc=wb.active; zc.title="z_Config"
zc["A1"]="URL file làm việc"; zc["B1"]=WORK_URL; zc["A2"]="Năm"; zc["B2"]=2026; zc["A3"]="URL file Database"; zc["B3"]=DB_URL
for rr in (1,2,3): zc[f"A{rr}"].font=F(bold=True)
zc.column_dimensions["A"].width=20; zc.column_dimensions["B"].width=70
def src(name,u,rng):
    ws=wb.create_sheet(name); ws["A1"]=f'=IMPORTRANGE({u},"{rng}")'; ws.sheet_state="hidden"; return ws
src("z_Trang",UW(),"02_UI_UX!A3:N3")
src("z_TinhNang",UW(),"03_Tính_năng!A3:N3")
src("z_TrangFull",UW(),"02_UI_UX!A4:R200")
src("z_TNFull",UW(),"03_Tính_năng!A4:T200")
src("z_Loi",UW(),"05_Lỗi_Tester!A4:V606")
src("z_CV",UW(),"01_Cong_viec_ngay!A4:P1004")
src("z_DB_DN",UD(),"'HS DN'!D4:N970")
src("z_DB_HR",UD(),"'HS HR'!D4:E966")
src("z_DB_UV",UD(),"'HS ỨNG TUYỂN'!D5:D990")
src("z_DB_Tin",UD(),"'Tin tuyển dụng'!C4:D1006")
zc.sheet_state="hidden"

T_TOT="'z_Trang'!$B$1"; T_DONE="'z_Trang'!$D$1"; T_PROG="'z_Trang'!$F$1"; T_ING="'z_Trang'!$H$1"; T_RE="'z_Trang'!$J$1"; T_UN="'z_Trang'!$L$1"
F_TOT="'z_TinhNang'!$B$1"; F_DONE="'z_TinhNang'!$D$1"; F_PROG="'z_TinhNang'!$F$1"; F_ING="'z_TinhNang'!$H$1"; F_RE="'z_TinhNang'!$J$1"; F_UN="'z_TinhNang'!$L$1"
U="'z_Loi'!$U:$U"; L="'z_Loi'!$L:$L"; M="'z_Loi'!$M:$M"; Q="'z_Loi'!$Q:$Q"; E="'z_Loi'!$E:$E"
def cif(rng,crit): return f'COUNTIF({rng},"{crit}")'
BUG_TOTAL=f'=COUNTA({E})-1'
OPEN=f'=({cif(U,"Open")}+{cif(U,"Chưa gửi Dev")})'
RESOLVED=f'=({cif(U,"Fixed")}+{cif(U,"Verified")}+{cif(U,"Closed")})'
CH_OPEN=(f'=(COUNTIFS({U},"Open",{L},"Critical")+COUNTIFS({U},"Open",{L},"High")'
         f'+COUNTIFS({U},"Chưa gửi Dev",{L},"Critical")+COUNTIFS({U},"Chưa gửi Dev",{L},"High"))')
DBDN="'z_DB_DN'!$A:$A"; DBDN_D="'z_DB_DN'!$K:$K"; DBHR="'z_DB_HR'!$A:$A"; DBHR_D="'z_DB_HR'!$B:$B"
DBUV="'z_DB_UV'!$A:$A"; DBTIN="'z_DB_Tin'!$A:$A"; DBTIN_D="'z_DB_Tin'!$B:$B"
LN="$N$4"; LO="$O$4"
def foundp(): return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO})'
def sevp(x):  return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO},{L},"{x}")'
def devp(x):  return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO},{M},"*{x}*")'
def cvdone(): return f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{LN},\'z_CV\'!$B:$B,"<="&{LO},\'z_CV\'!$O:$O,"Hoàn thành")'
def dbp(d):   return f'=COUNTIFS({d},">="&{LN},{d},"<="&{LO})'

def detail_query(rng,namec,bugc,datec,thirdc,statc,order_col,third_is_date,fb):
    fmt=f"{bugc} '#,##0'"+(f", {thirdc} 'dd/mm/yyyy'" if third_is_date else "")
    S1=f'select {namec}, {bugc}, {thirdc} where {datec} >= date \''
    S2=f'\' and {datec} <= date \''
    S3=(f'\' and ({statc} = \'Hoàn thành\' or {statc} = \'Kiểm tra lại\') order by {order_col} desc '
        f"label {namec} '', {bugc} '', {thirdc} '' format {fmt}")
    return ('=IFERROR(QUERY('+rng+',"'+S1+'"&TEXT($N$4,"yyyy-mm-dd")&"'+S2+
            '"&TEXT($O$4,"yyyy-mm-dd")&"'+S3+'",1),"'+fb+'")')

KPI_NAMES=["Trang hoàn thành","Tính năng hoàn thành","Công việc HT trong kỳ","Lỗi phát hiện trong kỳ",
           "Lỗi đã xử lý (lũy kế)","Số Trang đã xử lý","Số Tính năng đã xử lý",
           "DB Doanh nghiệp","DB HR","DB Ứng viên","Tin tuyển dụng"]

# ================= build_period =================
def build_period(sheetname, mode):
    s=wb.create_sheet(sheetname)
    for col,wd in {"A":32,"B":12,"C":15,"D":11,"E":11,"F":28,"G":12,"H":15}.items(): s.column_dimensions[col].width=wd
    s.column_dimensions["N"].hidden=True; s.column_dimensions["O"].hidden=True
    P="tháng" if mode=="month" else "tuần"
    def tref(k):
        if mode=="month":
            return f'=IFERROR(INDEX(\'04_Mục_tiêu\'!$B$5:$L$16,MATCH($B$4,\'04_Mục_tiêu\'!$A$5:$A$16,0),{k}),"")'
        return f'=IFERROR(INDEX(\'04_Mục_tiêu\'!$D$20:$N$79,MATCH($B$4*10+$D$4,\'04_Mục_tiêu\'!$C$20:$C$79,0),{k}),"")'
    band(s,"A1:F1",f"BÁO CÁO TIẾN ĐỘ THEO {P.upper()}",big=True)
    merge(s,"A2:F2",f"Chỉ số \"trong {P}\" lọc theo kỳ đã chọn; số Tổng quan là thời điểm hiện tại. Mục tiêu nhập tại sheet 04_Mục_tiêu (mỗi kỳ một dòng).",
          font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
    if mode=="month":
        cell(s,"A4","Tháng báo cáo",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"C4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"D4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"E4","Ngày chốt",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"F4","=TODAY()",F(11,True),Fill(INPUT),Al("center"),border=box); s["F4"].number_format="dd/mm/yyyy"
        s["N4"]="=DATE(D4,B4,1)"; s["O4"]="=EOMONTH(N4,0)"
        dv=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); s.add_data_validation(dv); dv.add(s["B4"]); selrow=5
    else:
        cell(s,"A4","Tháng",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"C4","Tuần",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"D4",2,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"E4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"F4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        s["N4"]="=DATE(F4,B4,(D4-1)*7+1)"; s["O4"]="=MIN(EOMONTH(DATE(F4,B4,1),0),DATE(F4,B4,D4*7))"
        d1=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); s.add_data_validation(d1); d1.add(s["B4"])
        d2=DataValidation(type="list",formula1='"1,2,3,4,5"'); s.add_data_validation(d2); d2.add(s["D4"])
        cell(s,"A5","Từ ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); cell(s,"B5","=N4",F(10.5,True),Fill(WHITE),Al("center"),border=box); s["B5"].number_format="dd/mm/yyyy"
        cell(s,"C5","Đến ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); merge(s,"D5:E5","=O4",font=F(10.5,True),fill=Fill(WHITE),al=Al("center")); s["D5"].number_format="dd/mm/yyyy"
        cell(s,"F5",f'=IF(ISERROR({T_TOT}),"⚠ Chưa kết nối","✔ Đã kết nối")',F(10,True,GREEN),Fill(WHITE),Al("center"),border=box); selrow=6
    def kpih(r):
        for i,h in enumerate(["KPI / CHỈ SỐ","Mục tiêu kỳ","Kết quả","Tỷ lệ đạt","Đạt KPI?","Ghi chú"]):
            c=get_column_letter(1+i); cell(s,f"{c}{r}",h,F(10.5,True,WHITE),Fill(NAVY),Al("center" if i else "left"),border=box)
        s.row_dimensions[r].height=22
    def sect(r,txt): merge(s,f"A{r}:F{r}",txt,font=F(11,True,NAVY),fill=Fill(LBLUE),al=Al("left"),border=True); s.row_dimensions[r].height=19
    def kpi(r,k,name,result,note,zebra=False):
        bg=Fill(ZEBRA) if zebra else Fill(WHITE)
        cell(s,f"A{r}",name,F(10.5),bg,Al("left",wrap=True),border=box)
        cell(s,f"B{r}",tref(k),F(10.5,True,NAVY),bg,Al("center"),border=box,nfmt=INT)
        cell(s,f"C{r}",result,F(11,True),bg,Al("center"),border=box,nfmt=INT)
        cell(s,f"D{r}",f'=IFERROR(C{r}/B{r},"")',F(10.5),bg,Al("center"),border=box,nfmt=PCT)
        cell(s,f"E{r}",f'=IF(ISNUMBER(B{r}),IF(C{r}>=B{r},"Đạt","Chưa đạt"),"Theo dõi")',F(10.5,True),bg,Al("center"),border=box)
        cell(s,f"F{r}",note,F(9.5,color=GREY),bg,Al("left",wrap=True),border=box)
    r=selrow+1; kpih(r); r+=1
    sect(r,"I. KIỂM TRA WEBSITE"); r+=1
    kpi(r,1,"1. Trang hoàn thành",f"={T_DONE}","Hoàn thành / tổng số trang"); r+=1
    kpi(r,2,"2. Tính năng hoàn thành",f"={F_DONE}","Hoàn thành / 88 tính năng",zebra=True); r+=1
    kpi(r,3,f"3. Công việc hoàn thành trong {P}",cvdone(),"Nhật ký công việc theo ngày"); r+=1
    kpi(r,4,f"4. Lỗi phát hiện trong {P}",foundp(),"Theo Ngày phát hiện (05_Lỗi_Tester)",zebra=True); r+=1
    sect(r,"II. HOÀN THIỆN WEBSITE"); r+=1
    kpi(r,5,"1. Lỗi đã xử lý (lũy kế)",RESOLVED,"Fixed + Verified + Closed"); r+=1
    kpi(r,6,"2. Số Trang đã xử lý",f"={T_TOT}-{T_UN}","Đã kiểm thử = Tổng − Chưa kiểm tra",zebra=True); r+=1
    kpi(r,7,"3. Số Tính năng đã xử lý",f"={F_TOT}-{F_UN}","Đã kiểm thử = Tổng − Chưa kiểm tra"); r+=1
    sect(r,f"III. DATABASE (trong {P})"); r+=1
    kpi(r,8,"1. DB Doanh nghiệp đã thêm",dbp(DBDN_D),"Theo Ngày Đăng/Tạo"); r+=1
    kpi(r,9,"2. DB HR đã thêm",dbp(DBHR_D),"Theo Ngày Đăng/Tạo",zebra=True); r+=1
    kpi(r,10,"3. DB Ứng viên đã thêm",f"=COUNTA({DBUV})","Theo tổng (chưa có cột ngày nhập)"); r+=1
    kpi(r,11,"4. Tin tuyển dụng đã đăng",dbp(DBTIN_D),"Theo Ngày viết",zebra=True); r+=1
    r+=1
    band(s,f"A{r}:F{r}","TỔNG QUAN HIỆN TẠI (thời điểm xem)"); r+=1
    for i,h in enumerate(["Chỉ số","Số lượng","Tổng","Tỷ lệ"]):
        c=get_column_letter(1+i); cell(s,f"{c}{r}",h,F(10.5,True,WHITE),Fill(BLUE),Al("center" if i else "left"),border=box)
    r+=1
    def stat(r,label,val,total,ratio,hi=None,zebra=False):
        bg=Fill(GREEN_BG) if hi=="g" else Fill(RED_BG) if hi=="r" else (Fill(ZEBRA) if zebra else Fill(WHITE))
        fc=GREEN if hi=="g" else RED if hi=="r" else "1A1A1A"
        cell(s,f"A{r}",label,F(10.5,bool(hi),fc),bg,Al("left"),border=box)
        cell(s,f"B{r}",val,F(12 if hi else 11,True,fc),bg,Al("center"),border=box,nfmt=INT)
        cell(s,f"C{r}",total,F(10.5,color=fc),bg,Al("center"),border=box,nfmt=INT if total is not None else None)
        cell(s,f"D{r}",ratio,F(11 if hi else 10.5,True,fc),bg,Al("center"),border=box,nfmt=PCT if ratio is not None else None)
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
    stat(r,"DB Doanh nghiệp",f"=COUNTA({DBDN})",None,None); r+=1
    stat(r,"DB HR",f"=COUNTA({DBHR})",None,None,zebra=True); r+=1
    stat(r,"DB Ứng viên",f"=COUNTA({DBUV})",None,None); r+=1
    stat(r,"Tin tuyển dụng đã đăng",f"=COUNTA({DBTIN})",None,None,zebra=True); r+=1
    r+=1
    band(s,f"A{r}:F{r}","TÌNH TRẠNG XỬ LÝ LỖI (toàn bộ)"); r+=1
    cell(s,f"A{r}","Trạng thái",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box); cell(s,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    merge(s,f"C{r}:F{r}","Ý nghĩa",font=F(10.5,True,WHITE),fill=Fill(BLUE),al=Al("left"))
    for cc in "CDEF": s[f"{cc}{r}"].border=box
    r+=1
    for nm,mean,clr,zb in [("Open","Lỗi đang mở, đã ghi nhận",RED,False),("Chưa gửi Dev","Chưa chuyển Dev xử lý",AMBER,True),
                           ("Fixed","Dev báo đã sửa",GREEN,False),("Verified","Đã retest đạt",GREEN,True),
                           ("Closed","Đã đóng",GREY,False),("Deferred","Tạm hoãn / không sửa",GREY,True)]:
        bg=Fill(ZEBRA) if zb else Fill(WHITE)
        cell(s,f"A{r}",nm,F(10.5,True,clr),bg,Al("left"),border=box)
        cell(s,f"B{r}",f'={cif(U,nm)}',F(11,True),bg,Al("center"),border=box,nfmt=INT)
        merge(s,f"C{r}:F{r}",mean,font=F(9.5,color=GREY),fill=bg,al=Al("left"))
        for cc in "CDEF": s[f"{cc}{r}"].border=box
        r+=1
    r+=1
    band(s,f"A{r}:F{r}",f"PHÂN BỐ LỖI PHÁT HIỆN TRONG {P.upper()}"); r+=1
    cell(s,f"A{r}","Theo mức độ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box); cell(s,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"D{r}","Theo thiết bị",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box); cell(s,f"E{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    r0=r; r+=1
    for i,(nm,clr) in enumerate([("Critical",RED),("High",RED),("Medium",AMBER),("Low",GREEN)]):
        bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
        cell(s,f"A{r}",nm,F(10.5,True,clr),bg,Al("left"),border=box); cell(s,f"B{r}",sevp(nm),F(11,True),bg,Al("center"),border=box,nfmt=INT); r+=1
    rr=r0+1
    for i,nm in enumerate(["Desktop","Mobile","Tablet"]):
        bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
        cell(s,f"D{rr}",nm,F(10.5),bg,Al("left"),border=box); cell(s,f"E{rr}",devp(nm),F(11,True),bg,Al("center"),border=box,nfmt=INT); rr+=1
    r=r0+5; r+=1
    merge(s,f"A{r}:F{r}","THANH TIẾN ĐỘ HOÀN THÀNH",font=F(11,True,WHITE),fill=Fill(BLUE),al=Al("left"),border=True); r+=1
    cell(s,f"A{r}","Trang",F(10.5,True),Fill(WHITE),Al("left"),border=box)
    merge(s,f"B{r}:F{r}",f'=REPT("█",ROUND({T_PROG}*30,0))&" "&TEXT({T_PROG},"0.0%")',font=F(11,True,BLUE),fill=Fill(WHITE),al=Al("left"))
    for cc in "BCDEF": s[f"{cc}{r}"].border=box
    r+=1
    cell(s,f"A{r}","Tính năng",F(10.5,True),Fill(ZEBRA),Al("left"),border=box)
    merge(s,f"B{r}:F{r}",f'=REPT("█",ROUND({F_PROG}*30,0))&" "&TEXT({F_PROG},"0.0%")',font=F(11,True,GREEN),fill=Fill(ZEBRA),al=Al("left"))
    for cc in "BCDEF": s[f"{cc}{r}"].border=box
    r+=2
    # ---- CHI TIẾT: Trang (A:C) và Tính năng (F:H) cạnh nhau ----
    band(s,f"A{r}:H{r}",f"CHI TIẾT: TRANG & TÍNH NĂNG ĐÃ TEST TRONG {P.upper()}"); r+=1
    merge(s,f"A{r}:H{r}","Trang/tính năng có Ngày test gần nhất trong kỳ (trạng thái Hoàn thành hoặc Kiểm tra lại).",
          font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left")); r+=1
    third_hdr="Trạng thái" if mode=="month" else "Ngày test/retest"
    cell(s,f"A{r}","TRANG đã test trong kỳ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
    cell(s,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"C{r}",third_hdr,F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"F{r}","TÍNH NĂNG đã test trong kỳ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
    cell(s,f"G{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"H{r}",third_hdr,F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    r+=1
    isd = (mode=="week")
    if mode=="month":
        qp=detail_query("'z_TrangFull'!$A$1:$R$197","B","I","P","E","E","I",False,"(không có trang test trong kỳ)")
        qf=detail_query("'z_TNFull'!$A$1:$T$197","B","K","R","J","J","K",False,"(không có tính năng test trong kỳ)")
    else:
        qp=detail_query("'z_TrangFull'!$A$1:$R$197","B","I","P","P","E","P",True,"(không có trang test trong kỳ)")
        qf=detail_query("'z_TNFull'!$A$1:$T$197","B","K","R","R","J","R",True,"(không có tính năng test trong kỳ)")
    cell(s,f"A{r}",qp,F(10),Fill(WHITE),Al("left"))
    cell(s,f"F{r}",qf,F(10),Fill(WHITE),Al("left"))
    for rr2 in range(r,r+66):
        for cc in "ABC": s[f"{cc}{rr2}"].border=box
        for cc in "FGH": s[f"{cc}{rr2}"].border=box
        if isd:
            s[f"C{rr2}"].number_format="dd/mm/yyyy"; s[f"H{rr2}"].number_format="dd/mm/yyyy"
        s[f"B{rr2}"].number_format=INT; s[f"G{rr2}"].number_format=INT
    s.sheet_view.showGridLines=False
    return s

# ============ 04_Mục_tiêu ============
def build_targets():
    t=wb.create_sheet("04_Mục_tiêu")
    for col,wd in {"A":10,"B":10,"C":8}.items(): t.column_dimensions[col].width=wd
    band(t,"A1:N1","MỤC TIÊU KPI THEO KỲ — nhập tay (ô nền vàng)",big=True)
    # ---- Monthly grid: A5:A16 = tháng 1..12 ; B..L = 11 KPI ----
    band(t,"A3:N3","MỤC TIÊU THEO THÁNG")
    cell(t,"A4","Tháng",F(10,True,WHITE),Fill(NAVY),Al("center"),border=box)
    for i,nm in enumerate(KPI_NAMES):
        c=get_column_letter(2+i); cell(t,f"{c}4",nm,F(9,True,WHITE),Fill(NAVY),Al("center",wrap=True),border=box); t.column_dimensions[c].width=13
    t.row_dimensions[4].height=40
    for mth in range(1,13):
        rr=4+mth
        cell(t,f"A{rr}",mth,F(10,True),Fill(LBLUE),Al("center"),border=box)
        for i in range(11):
            c=get_column_letter(2+i); cell(t,f"{c}{rr}",None,F(10),Fill(INPUT),Al("center"),border=box,nfmt=INT)
    # ---- Weekly grid: A20:A79 tháng, B week, C key ; D..N = 11 KPI ----
    band(t,"A18:N18","MỤC TIÊU THEO TUẦN")
    cell(t,"A19","Tháng",F(10,True,WHITE),Fill(NAVY),Al("center"),border=box)
    cell(t,"B19","Tuần",F(10,True,WHITE),Fill(NAVY),Al("center"),border=box)
    cell(t,"C19","Key",F(9,True,WHITE),Fill(GREY),Al("center"),border=box)
    for i,nm in enumerate(KPI_NAMES):
        c=get_column_letter(4+i); cell(t,f"{c}19",nm,F(9,True,WHITE),Fill(NAVY),Al("center",wrap=True),border=box); t.column_dimensions[c].width=13
    t.row_dimensions[19].height=40
    rr=20
    for mth in range(1,13):
        for wk in range(1,6):
            cell(t,f"A{rr}",mth,F(10),Fill(LBLUE),Al("center"),border=box)
            cell(t,f"B{rr}",wk,F(10),Fill(LBLUE),Al("center"),border=box)
            cell(t,f"C{rr}",f"=A{rr}*10+B{rr}",F(9,color=GREY),Fill(WHITE),Al("center"),border=box)
            for i in range(11):
                c=get_column_letter(4+i); cell(t,f"{c}{rr}",None,F(10),Fill(INPUT),Al("center"),border=box,nfmt=INT)
            rr+=1
    t.sheet_view.showGridLines=False; t.sheet_properties.tabColor="7A5AA6"
    return t

# ============ 00_Hướng_dẫn ============
g=wb.create_sheet("00_Hướng_dẫn")
for col,w in {"A":26,"B":30,"C":28,"D":20,"E":20,"F":16}.items(): g.column_dimensions[col].width=w
band(g,"A1:F1","BÁO CÁO TIẾN ĐỘ WEBSITE TIMVIEC123 — DÀNH CHO BAN GIÁM ĐỐC",big=True)
merge(g,"A2:F2","Số liệu đồng bộ live từ file làm việc QA và file Database. BGD chỉ xem, chọn kỳ ở sheet Tháng/Tuần.",
      font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left")); g.row_dimensions[2].height=20
merge(g,"A4:B4","NGƯỜI DÙNG",font=F(11,True,WHITE),fill=Fill(BLUE)); merge(g,"C4:F4","CÁCH SỬ DỤNG",font=F(11,True,WHITE),fill=Fill(BLUE))
r=5
for who,how in [("Ban Giám đốc","Xem tiến độ trang/tính năng, tình trạng lỗi, database và nghiệm thu. Chọn Tháng/Tuần để lọc kỳ."),
                ("Tester","Cập nhật dữ liệu tại file làm việc; nhập Mục tiêu tại sheet 04_Mục_tiêu."),
                ("Dev","Cập nhật xử lý lỗi tại 06_RTM_Dev_Test của file làm việc.")]:
    merge(g,f"A{r}:B{r}",who,font=F(11,True),fill=Fill(LBLUE),al=Al("left")); merge(g,f"C{r}:F{r}",how,font=F(10.5),al=Al("left",wrap=True)); g.row_dimensions[r].height=30; r+=1
r+=1; band(g,f"A{r}:F{r}","NGUYÊN TẮC TÍNH SỐ"); r+=1
for i,p in enumerate([
 "Tiến độ trang = số trang Hoàn thành / tổng số trang (hiện 63). Tính năng / 88.",
 "Số lỗi lấy toàn bộ từ 05_Lỗi_Tester. \"Đang mở\" = Open + Chưa gửi Dev; \"Đã xử lý\" = Fixed + Verified + Closed.",
 "Mục tiêu nhập tại sheet 04_Mục_tiêu theo từng tháng/tuần (không cố định, không dùng chung).",
 "Chi tiết theo Trang/Tính năng = mục có Ngày test gần nhất trong kỳ, trạng thái Hoàn thành/Kiểm tra lại.",
 "Chỉ số Database lấy từ file Hồ sơ, đếm theo ngày đăng/tạo của kỳ."],1):
    cell(g,f"A{r}",i,F(11,True,BLUE),Fill(ZEBRA),Al("center"),border=box); merge(g,f"B{r}:F{r}",p,font=F(10.5),al=Al("left",wrap=True))
    for cc in "BCDEF": g[f"{cc}{r}"].border=box
    g.row_dimensions[r].height=26; r+=1
r+=1; band(g,f"A{r}:F{r}","KÍCH HOẠT KẾT NỐI (làm 1 lần)"); r+=1
cell(g,f"A{r}","① File làm việc QA",F(11,True,RED),Fill("FFF3CD"),Al("left"),border=box)
merge(g,f"B{r}:F{r}",f'=IMPORTRANGE({UW()},"05_Lỗi_Tester!E4")',font=F(11,True),fill=Fill("FFF3CD"),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=1
cell(g,f"A{r}","② File Database",F(11,True,RED),Fill("FFF3CD"),Al("left"),border=box)
merge(g,f"B{r}:F{r}",f'=IMPORTRANGE({UD()},"\'HS DN\'!D3")',font=F(11,True),fill=Fill("FFF3CD"),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=1
merge(g,f"A{r}:F{r}","→ Nếu ô ① hoặc ② báo #REF!: bấm vào ô đó → \"Cho phép truy cập / Allow access\". Làm 1 lần cho mỗi ô.",
      font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left",wrap=True)); g.row_dimensions[r].height=28
r+=1
cell(g,f"A{r}","Trạng thái kết nối",F(11,True),Fill(LBLUE),border=box)
merge(g,f"B{r}:F{r}",f'=IF(ISERROR({T_TOT}),"⚠ Chưa kết nối file làm việc","✔ Đã kết nối")&"   |   "&IF(ISERROR({DBDN}),"⚠ Chưa kết nối Database","✔ Database OK")',
      font=F(10.5,True,GREEN),al=Al("left"))
for cc in "BCDEF": g[f"{cc}{r}"].border=box
r+=2; band(g,f"A{r}:F{r}","LIÊN KẾT"); r+=1
for lbl,url in [("File làm việc (QA)",WORK_URL),("File Database (Hồ sơ)",DB_URL)]:
    cell(g,f"A{r}",lbl,F(11,True),Fill(LBLUE),border=box); merge(g,f"B{r}:F{r}",url,font=F(10,color="1155CC"),al=Al("left"))
    for cc in "BCDEF": g[f"{cc}{r}"].border=box
    r+=1
g.sheet_view.showGridLines=False

m=build_period("01_Báo_cáo_Tháng","month")
w=build_period("02_Báo_cáo_Tuần","week")

# ============ 03_Nghiệm_thu ============
KT="'z_Loi'!$K:$K"
def _open(col,crit):
    return f'(COUNTIFS({U},"Open",{col},"{crit}")+COUNTIFS({U},"Chưa gửi Dev",{col},"{crit}"))'
def sev_total(x): return f'=COUNTIF({L},"{x}")'
def sev_open(x):  return "="+_open(L,x)
def sta_cnt(x):   return f'=COUNTIF({U},"{x}")'
def typ_total(x): return f'=COUNTIF({KT},"*{x}*")'
def typ_open(x):  return "="+_open(KT,f"*{x}*")
def typ_ch(x):
    return ('=('+f'COUNTIFS({U},"Open",{KT},"*{x}*",{L},"Critical")+COUNTIFS({U},"Chưa gửi Dev",{KT},"*{x}*",{L},"Critical")'
            f'+COUNTIFS({U},"Open",{KT},"*{x}*",{L},"High")+COUNTIFS({U},"Chưa gửi Dev",{KT},"*{x}*",{L},"High")'+')')
def dev_open(x):  return "="+_open(M,f"*{x}*")
BLOCKING_QUERY=('=IFERROR(QUERY(\'z_Loi\'!$A:$V,'
 '"select E,B,C,L,K,Q,U where (L = \'Critical\' or L = \'High\') and (U = \'Open\' or U = \'Chưa gửi Dev\') '
 'order by L asc, Q desc limit 30 '
 'label E \'\', B \'\', C \'\', L \'\', K \'\', Q \'\', U \'\' format Q \'dd/mm/yyyy\'",1),'
 '"Không còn lỗi Critical/High đang mở — đủ điều kiện về lỗi chặn")')

def build_acceptance():
    n=wb.create_sheet("03_Nghiệm_thu")
    for col,wd in {"A":22,"B":34,"C":17,"D":17,"E":14,"F":15,"G":17,"H":22,"I":28}.items():
        n.column_dimensions[col].width=wd
    dv_kl=DataValidation(type="list",formula1='"Đạt,Đạt có điều kiện,Không đạt,Chưa nghiệm thu"',allow_blank=True)
    dv_dg=DataValidation(type="list",formula1='"Đạt,Chưa đạt,Không áp dụng,Chưa đo"',allow_blank=True)
    dv_bg=DataValidation(type="list",formula1='"Đã bàn giao,Chưa bàn giao,Không áp dụng"',allow_blank=True)
    for dv in (dv_kl,dv_dg,dv_bg): n.add_data_validation(dv)
    row=[1]
    def gap(h=8):
        n.row_dimensions[row[0]].height=h; row[0]+=1
    def title(text):
        r=row[0]; band(n,f"A{r}:I{r}",text,big=True); row[0]+=1; return r
    def sec(text):
        r=row[0]; band(n,f"A{r}:I{r}",text); row[0]+=1; return r
    def note(text,h=26):
        r=row[0]; merge(n,f"A{r}:I{r}",text,font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left",wrap=True))
        n.row_dimensions[r].height=h; row[0]+=1; return r
    def head(labels,h=32):
        r=row[0]
        for i,lb in enumerate(labels):
            cell(n,f"{get_column_letter(1+i)}{r}",lb,F(9.5,True,WHITE),Fill(NAVY),Al("center",wrap=True),border=box)
        n.row_dimensions[r].height=h; row[0]+=1; return r
    def line(vals,i=0,inputs=(),nfmts=None,h=24,left=(0,1),bold=()):
        r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
        for j in range(9):
            c=get_column_letter(1+j); v=vals[j] if j<len(vals) else ""
            cell(n,f"{c}{r}",v if v!="" else None,F(10,j in bold),Fill(INPUT) if j in inputs else bg,
                 Al("left" if j in left else "center",wrap=True),border=box,nfmt=(nfmts or {}).get(j))
        n.row_dimensions[r].height=h; row[0]+=1; return r
    def kv(pairs,h=24):
        """Tối đa 3 cặp Nhãn/Giá trị trên 1 dòng: A|B:C — D|E:F — G|H:I"""
        r=row[0]; slots=[("A","B:C"),("D","E:F"),("G","H:I")]
        for k,(lc,vr) in enumerate(slots):
            if k<len(pairs):
                item=pairs[k]; lb,val=item[0],item[1]; nf=item[2] if len(item)>2 else None
                cell(n,f"{lc}{r}",lb,F(10,True,WHITE),Fill(BLUE),Al("left",wrap=True),border=box)
                a,b=vr.split(":"); merge(n,f"{a}{r}:{b}{r}",val if val!="" else None,font=F(10),fill=Fill(INPUT),al=Al("left"))
                for cc in (a,b): n[f"{cc}{r}"].border=box
                if nf: n[f"{a}{r}"].number_format=nf
            else:
                for cc in [lc]+vr.split(":"): cell(n,f"{cc}{r}","",F(10),Fill(WHITE),Al("left"),border=box)
        n.row_dimensions[r].height=h; row[0]+=1; return r

    title("BIÊN BẢN NGHIỆM THU WEBSITE TIMVIEC123")
    note("Ô nền VÀNG = nhập tay. Ô nền trắng/xám có công thức = tự động lấy từ dữ liệu QA (02_UI_UX, 03_Tính_năng, 05_Lỗi_Tester) — không sửa tay. "
         "Phần nào chưa có dữ liệu nguồn thì để trống và ghi rõ trong mục 'Thông tin còn thiếu'.",h=30)
    gap()

    # ---------- I. THÔNG TIN CHUNG ----------
    sec("I. THÔNG TIN CHUNG")
    kv([("Số biên bản",""),("Ngày lập","=TODAY()","dd/mm/yyyy"),("Địa điểm lập","")])
    kv([("Tên dự án","Website TimViec123"),("Tên miền / URL",""),("Môi trường nghiệm thu","")])
    kv([("Giai đoạn nghiệm thu","Giai đoạn 1"),("Phiên bản / Release",""),("Kỳ kiểm thử (từ)","","dd/mm/yyyy")])
    kv([("Kỳ kiểm thử (đến)","","dd/mm/yyyy"),("Hợp đồng số",""),("Ngày hợp đồng","","dd/mm/yyyy")])
    kv([("Tài liệu yêu cầu (SRS/BRD)",""),("Kế hoạch kiểm thử",""),("Lần nghiệm thu thứ","")])
    gap()

    # ---------- II. CÁC BÊN THAM GIA ----------
    sec("II. CÁC BÊN THAM GIA")
    head(["Vai trò","Đơn vị / Bộ phận","Họ và tên","Chức vụ","Email","Điện thoại","Trách nhiệm","Ghi chú",""])
    parties=[("Bên A — Chủ đầu tư","","","","","","Phê duyệt & ký nghiệm thu"),
             ("Bên B — Đơn vị thực hiện","","","","","","Bàn giao sản phẩm, xử lý lỗi"),
             ("Quản lý dự án (PM)","","","","","","Điều phối, chốt phạm vi"),
             ("QA / Tester","","Mr. Ẩn","","","","Thực hiện kiểm thử, lập hồ sơ lỗi"),
             ("Đại diện Dev","","","","","","Sửa lỗi, xác nhận kỹ thuật"),
             ("Người duyệt cuối","","","","","","Ra kết luận nghiệm thu")]
    for i,p in enumerate(parties): line(list(p),i,inputs=(1,2,3,4,5),left=(0,1,2,6,7))
    gap()

    # ---------- III. PHẠM VI NGHIỆM THU ----------
    sec("III. PHẠM VI NGHIỆM THU")
    note("Nghiệm thu chỉ xét đúng các hạng mục liệt kê trong phạm vi. Hạng mục ngoài phạm vi không dùng làm lý do từ chối nghiệm thu.",h=22)
    head(["Nhóm","Nội dung trong phạm vi","Số lượng trong phạm vi","Tổng trên hệ thống","Tỷ lệ phạm vi","Tài liệu tham chiếu","Ngoài phạm vi (loại trừ)","Ghi chú",""])
    line(["Trang (DS)","Danh sách DS được chốt nghiệm thu","",f"={T_TOT}","=IFERROR(C/D,\"\")","","",""],0,
         inputs=(2,5,6,7),nfmts={3:INT,4:PCT},left=(0,1,5,6,7))
    n[f"E{row[0]-1}"]=f"=IFERROR(C{row[0]-1}/D{row[0]-1},\"\")"
    line(["Tính năng (YC)","Danh sách YC được chốt nghiệm thu","",f"={F_TOT}","","","",""],1,
         inputs=(2,5,6,7),nfmts={3:INT,4:PCT},left=(0,1,5,6,7))
    n[f"E{row[0]-1}"]=f"=IFERROR(C{row[0]-1}/D{row[0]-1},\"\")"
    line(["Vai trò người dùng","Khách / Ứng viên / Nhà tuyển dụng / Admin","","","","","",""],2,inputs=(2,3,5,6,7),left=(0,1,5,6,7))
    line(["Tích hợp / API","Danh sách tích hợp bên thứ ba","","","","","",""],3,inputs=(2,3,5,6,7),left=(0,1,5,6,7))
    gap()

    # ---------- IV. MÔI TRƯỜNG KIỂM THỬ ----------
    sec("IV. MÔI TRƯỜNG KIỂM THỬ")
    head(["Hạng mục","Cấu hình / danh sách cụ thể","Phiên bản","Đã kiểm thử","Ghi chú","","","",""])
    envs=["Trình duyệt Desktop","Trình duyệt Mobile","Thiết bị di động","Máy tính bảng","Độ phân giải chuẩn",
          "Hệ điều hành","Tài khoản kiểm thử","Dữ liệu kiểm thử","Đường truyền / mạng"]
    for i,e in enumerate(envs): line([e,"","","",""],i,inputs=(1,2,3,4),left=(0,1,4))
    gap()

    # ---------- V. TIÊU CHÍ NGHIỆM THU ----------
    sec("V. TIÊU CHÍ NGHIỆM THU (ACCEPTANCE / EXIT CRITERIA)")
    note("Ngưỡng bắt buộc do hai bên chốt trước khi kiểm thử. Kết quả thực tế tự động tính; đánh giá tự so sánh với ngưỡng.",h=22)
    head(["Mã","Tiêu chí","Ngưỡng bắt buộc","Kết quả thực tế","Đánh giá","Bắt buộc?","Nguồn số liệu","Ghi chú",""])
    crit_start=row[0]
    def crit(code,name,thr,res,cmp_,must,srcv,nf=None,pct=False):
        r=row[0]; i=r-crit_start
        ins=(2,7) if res!="" else (2,3,7)
        line([code,name,thr,res,"",must,srcv,""],i,inputs=ins,
             nfmts={2:(PCT if pct else INT),3:(PCT if pct else (nf or INT))},left=(0,1,6,7))
        if cmp_=="le":  n[f"E{r}"]=f'=IF(C{r}="","Chưa chốt ngưỡng",IF(D{r}<=C{r},"Đạt","Chưa đạt"))'
        elif cmp_=="ge":n[f"E{r}"]=f'=IF(C{r}="","Chưa chốt ngưỡng",IF(D{r}>=C{r},"Đạt","Chưa đạt"))'
        else:
            n[f"E{r}"]=None; n[f"E{r}"].fill=Fill(INPUT); dv_dg.add(n[f"E{r}"])
        n[f"E{r}"].font=F(10,True); n[f"E{r}"].alignment=Al("center")
        return r
    crit("TC01","Lỗi Critical đang mở",0,sev_open("Critical"),"le","Bắt buộc","05_Lỗi_Tester (L, U)")
    crit("TC02","Lỗi High đang mở",0,sev_open("High"),"le","Bắt buộc","05_Lỗi_Tester (L, U)")
    crit("TC03","Lỗi Medium đang mở",5,sev_open("Medium"),"le","Bắt buộc","05_Lỗi_Tester (L, U)")
    crit("TC04","Lỗi Low đang mở","",sev_open("Low"),"le","Khuyến nghị","05_Lỗi_Tester (L, U)")
    crit("TC05","Tỷ lệ trang hoàn thành kiểm thử",1,f"={T_PROG}","ge","Bắt buộc","02_UI_UX",pct=True)
    crit("TC06","Tỷ lệ tính năng hoàn thành kiểm thử",1,f"={F_PROG}","ge","Bắt buộc","03_Tính_năng",pct=True)
    crit("TC07","Tỷ lệ lỗi đã xử lý (Fixed/Verified/Closed)",0.95,
         f'=IFERROR({RESOLVED[1:]}/(COUNTA({E})-1),0)',"ge","Bắt buộc","05_Lỗi_Tester (U)",pct=True)
    crit("TC08","Tỷ lệ lỗi đã retest đạt (Verified/Closed)",0.9,
         f'=IFERROR((COUNTIF({U},"Verified")+COUNTIF({U},"Closed"))/(COUNTA({E})-1),0)',"ge","Bắt buộc","05_Lỗi_Tester (U)",pct=True)
    crit("TC09","Tỷ lệ Test Case Pass","","","ge","Bắt buộc","06_RTM_Dev_Test — CHƯA CÓ SỐ",pct=True)
    crit("TC10","Độ phủ yêu cầu (RTM: YC có Test Case)","","","ge","Bắt buộc","06_RTM_Dev_Test — CHƯA CÓ SỐ",pct=True)
    crit("TC11","Lỗi bảo mật Critical/High đang mở",0,typ_ch("Security"),"le","Bắt buộc","05_Lỗi_Tester (K, L, U)")
    crit("TC12","Lỗi Responsive Critical/High đang mở",0,typ_ch("Responsive"),"le","Bắt buộc","05_Lỗi_Tester (K, L, U)")
    crit("TC13","Hiệu năng trang chủ (điểm PageSpeed / LCP)","","","man","Bắt buộc","Đo bằng công cụ — CHƯA CÓ SỐ")
    crit("TC14","Tương thích đủ danh mục trình duyệt/thiết bị đã chốt","","","man","Bắt buộc","Mục IV — CHƯA CÓ SỐ")
    crit("TC15","Bằng chứng kiểm thử đầy đủ (ảnh/video/link retest)","","","man","Bắt buộc","05_Lỗi_Tester (O, P)")
    crit("TC16","Không còn lỗi hồi quy (regression) đang mở","","","man","Bắt buộc","Kết quả retest — CHƯA CÓ SỐ")
    crit("TC17","Tài liệu bàn giao đầy đủ","","","man","Bắt buộc","Mục XII")
    crit_end=row[0]-1
    gap()

    # ---------- VI. ĐỊNH NGHĨA MỨC ĐỘ LỖI & SLA ----------
    sec("VI. ĐỊNH NGHĨA MỨC ĐỘ LỖI VÀ SLA XỬ LÝ")
    head(["Mức độ","Định nghĩa","Ảnh hưởng","SLA xử lý (giờ/ngày)","Chặn nghiệm thu?","Số lỗi hiện có","Đang mở","Ghi chú",""])
    sevs=[("Critical","Chặn luồng nghiệp vụ chính, mất dữ liệu, lỗ hổng bảo mật nghiêm trọng","Không thể sử dụng","","Có"),
          ("High","Sai nghiệp vụ, tính năng chính không hoạt động đúng","Ảnh hưởng lớn","","Có"),
          ("Medium","Sai lệch hiển thị/nghiệp vụ phụ, có cách làm thay thế","Ảnh hưởng vừa","","Theo ngưỡng TC03"),
          ("Low","Lỗi nhỏ về giao diện, chính tả, trải nghiệm","Ảnh hưởng thấp","","Không")]
    for i,(s,d,a,sla,blk) in enumerate(sevs):
        line([s,d,a,sla,blk,sev_total(s),sev_open(s),""],i,inputs=(3,7),nfmts={5:INT,6:INT},left=(0,1,2,7))
    gap()

    # ---------- VII. KẾT QUẢ KIỂM THỬ ----------
    sec("VII. KẾT QUẢ KIỂM THỬ — TỔNG HỢP (tự động)")
    head(["Chỉ số","Diễn giải","Số lượng","Tổng","Tỷ lệ","Chỉ số","Diễn giải","Số lượng","Tỷ lệ"])
    stats=[("Trang hoàn thành","Đã test xong, không còn lỗi chặn",f"={T_DONE}",f"={T_TOT}",f"={T_PROG}",
            "Tổng lỗi ghi nhận","Số dòng có BUG_ID",BUG_TOTAL,""),
           ("Trang cần kiểm tra lại","Đã test, còn lỗi phải retest",f"={T_RE}",f"={T_TOT}","",
            "Lỗi đang mở","Open + Chưa gửi Dev",OPEN,""),
           ("Trang chưa kiểm tra","Chưa bắt đầu kiểm thử",f"={T_UN}",f"={T_TOT}","",
            "Lỗi đã xử lý","Fixed + Verified + Closed",RESOLVED,""),
           ("Tính năng hoàn thành","Đã test xong",f"={F_DONE}",f"={F_TOT}",f"={F_PROG}",
            "Lỗi Critical/High đang mở","Chặn nghiệm thu",CH_OPEN,""),
           ("Tính năng cần kiểm tra lại","Còn lỗi phải retest",f"={F_RE}",f"={F_TOT}","",
            "Lỗi tạm hoãn (Deferred)","Cần ghi nhận ngoại lệ",sta_cnt("Deferred"),""),
           ("Tính năng chưa kiểm tra","Chưa bắt đầu kiểm thử",f"={F_UN}",f"={F_TOT}","",
            "Lỗi đã retest đạt","Verified + Closed",f'=(COUNTIF({U},"Verified")+COUNTIF({U},"Closed"))',"")]
    for i,(a,b,c,d,e,f2,g2,h2,i2) in enumerate(stats):
        r=line([a,b,c,d,e,f2,g2,h2,i2],i,nfmts={2:INT,3:INT,4:PCT,7:INT,8:PCT},left=(0,1,5,6))
        if not e: n[f"E{r}"]=f'=IFERROR(C{r}/D{r},"")'; n[f"E{r}"].number_format=PCT
        if i>0:
            n[f"I{r}"]=f'=IFERROR(H{r}/({BUG_TOTAL[1:]}),"")'; n[f"I{r}"].number_format=PCT
    gap()

    # ---------- VIII. THỐNG KÊ LỖI ----------
    sec("VIII. THỐNG KÊ LỖI (tự động)")
    head(["Theo trạng thái","Số lỗi","Tỷ lệ","Theo mức độ","Tổng","Đang mở","Theo thiết bị / loại","Tổng","Đang mở"])
    st=[("Open","Critical","Desktop"),("Chưa gửi Dev","High","Mobile"),("Fixed","Medium","Tablet"),
        ("Verified","Low","UI/UX"),("Closed",None,"Functional"),("Deferred",None,"Security")]
    for i,(s1,s2,s3) in enumerate(st):
        vals=[s1,sta_cnt(s1),"",s2 or "",sev_total(s2) if s2 else "",sev_open(s2) if s2 else "",s3,"",""]
        r=line(vals,i,nfmts={1:INT,2:PCT,4:INT,5:INT,7:INT,8:INT},left=(0,3,6))
        n[f"C{r}"]=f'=IFERROR(B{r}/({BUG_TOTAL[1:]}),"")'; n[f"C{r}"].number_format=PCT
        if s3 in ("Desktop","Mobile","Tablet"):
            n[f"H{r}"]=f'=COUNTIF({M},"*{s3}*")'; n[f"I{r}"]=dev_open(s3)
        else:
            n[f"H{r}"]=typ_total(s3); n[f"I{r}"]=typ_open(s3)
        n[f"H{r}"].number_format=INT; n[f"I{r}"].number_format=INT
    gap()

    # ---------- IX. KẾT QUẢ NGHIỆM THU THEO HẠNG MỤC ----------
    sec("IX. KẾT QUẢ NGHIỆM THU THEO HẠNG MỤC")
    head(["Hạng mục","Phạm vi","Tổng","Hoàn thành","Tỷ lệ","Lỗi đang mở","Critical/High mở","Kết luận","Ghi chú"])
    cat_start=row[0]
    cats=[("Giao diện & trải nghiệm (UI/UX)","Trang website",f"={T_TOT}",f"={T_DONE}",f"={T_PROG}",typ_open("UI/UX"),typ_ch("UI/UX")),
          ("Chức năng nghiệp vụ","Tính năng theo yêu cầu",f"={F_TOT}",f"={F_DONE}",f"={F_PROG}",typ_open("Functional"),typ_ch("Functional")),
          ("Tương thích & Responsive","Desktop / Mobile / Tablet","","","",typ_open("Responsive"),typ_ch("Responsive")),
          ("Bảo mật","Test case Security, phân quyền, dữ liệu","","","",typ_open("Security"),typ_ch("Security")),
          ("Dữ liệu & tích hợp (API/DB)","Luồng dữ liệu, API bên thứ ba","","","","",""),
          ("Hiệu năng","Tốc độ tải, chịu tải","","","","",""),
          ("SEO & chuẩn nội dung","Meta, sitemap, robots, URL","","","","",""),
          ("Khả năng truy cập (Accessibility)","Chuẩn WCAG cơ bản","","","","",""),
          ("Vận hành & bàn giao","Tài liệu, tài khoản, hướng dẫn","","","","","")]
    for i,c in enumerate(cats):
        r=line(list(c)+["",""],i,inputs=(8,),nfmts={2:INT,3:INT,4:PCT,5:INT,6:INT},left=(0,1,8),h=26)
        if c[5]=="":
            for cc in ("C","D","E","F","G"): n[f"{cc}{r}"].fill=Fill(INPUT)
            n[f"H{r}"]=None
        else:
            n[f"H{r}"]=f'=IF(G{r}>0,"Không đạt",IF(F{r}>0,"Đạt có điều kiện","Đạt"))'
        n[f"H{r}"].font=F(10,True); n[f"H{r}"].alignment=Al("center")
        if c[5]=="": n[f"H{r}"].fill=Fill(INPUT); dv_kl.add(n[f"H{r}"])
    cat_end=row[0]-1
    gap()

    # ---------- X. KIỂM THỬ PHI CHỨC NĂNG ----------
    sec("X. KẾT QUẢ KIỂM THỬ PHI CHỨC NĂNG (nhập tay — cần bổ sung số đo)")
    head(["Hạng mục","Chỉ số đo","Công cụ đo","Ngưỡng yêu cầu","Kết quả đo","Ngày đo","Đánh giá","Bằng chứng / link","Ghi chú"])
    nfr=[("Hiệu năng","Điểm PageSpeed Desktop"),("Hiệu năng","Điểm PageSpeed Mobile"),("Hiệu năng","LCP (giây)"),
         ("Hiệu năng","Thời gian tải trang chủ (giây)"),("Chịu tải","Số người dùng đồng thời"),
         ("Bảo mật","Quét lỗ hổng (OWASP Top 10)"),("Bảo mật","HTTPS / chứng chỉ SSL"),("Bảo mật","Phân quyền theo vai trò"),
         ("SEO","Sitemap / robots.txt / meta"),("Accessibility","Tương phản & thao tác bàn phím"),
         ("Sao lưu","Cơ chế backup & phục hồi")]
    for i,(a,b) in enumerate(nfr):
        r=line([a,b,"","","","","","",""],i,inputs=(2,3,4,5,7,8),nfmts={5:"dd/mm/yyyy"},left=(0,1,7,8))
        n[f"G{r}"].fill=Fill(INPUT); dv_dg.add(n[f"G{r}"])
    gap()

    # ---------- XI. LỖI TỒN ĐỌNG CHẶN NGHIỆM THU ----------
    sec("XI. LỖI TỒN ĐỌNG ĐANG CHẶN NGHIỆM THU (Critical/High còn mở)")
    note("Danh sách tự động, hiển thị tối đa 30 lỗi ưu tiên. Danh sách đầy đủ xem sheet 05_Lỗi_Tester của file QA. "
         "Không nhập gì vào vùng bảng bên dưới (công thức sẽ tự đổ dữ liệu).",h=26)
    head(["BUG_ID","Trang","Tính năng","Mức độ","Loại lỗi","Ngày phát hiện","Trạng thái","Người xử lý","Hạn xử lý"])
    r=row[0]; cell(n,f"A{r}",BLOCKING_QUERY,F(10),Fill(WHITE),Al("left"))
    for k in range(32):
        for j in range(9):
            n[f"{get_column_letter(1+j)}{r+k}"].border=box
        n.row_dimensions[r+k].height=20
    row[0]=r+32
    gap()

    # ---------- XII. NGOẠI LỆ ĐƯỢC CHẤP NHẬN ----------
    sec("XII. NGOẠI LỆ / LỖI ĐƯỢC CHẤP NHẬN BỎ QUA (WAIVER)")
    note("Mọi lỗi Critical/High không xử lý trước nghiệm thu phải có ngoại lệ được người duyệt của Bên A ký chấp nhận.",h=22)
    head(["STT","BUG_ID / Hạng mục","Mức độ","Lý do chấp nhận","Ảnh hưởng còn lại","Phương án tạm thời","Hạn xử lý dứt điểm","Người duyệt","Ngày duyệt"])
    for i in range(6):
        line([i+1,"","","","","","","",""],i,inputs=(1,2,3,4,5,6,7,8),nfmts={6:"dd/mm/yyyy",8:"dd/mm/yyyy"},left=(1,3,4,5))
    gap()

    # ---------- XIII. RỦI RO & HẠN CHẾ ----------
    sec("XIII. RỦI RO VÀ HẠN CHẾ ĐÃ BIẾT")
    head(["STT","Rủi ro / hạn chế","Khả năng xảy ra","Mức ảnh hưởng","Biện pháp giảm thiểu","Người theo dõi","Hạn xử lý","Trạng thái","Ghi chú"])
    for i in range(5):
        line([i+1,"","","","","","","",""],i,inputs=(1,2,3,4,5,6,7,8),nfmts={6:"dd/mm/yyyy"},left=(1,4,8))
    gap()

    # ---------- XIV. HẠNG MỤC BÀN GIAO ----------
    sec("XIV. HẠNG MỤC BÀN GIAO")
    head(["STT","Hạng mục bàn giao","Định dạng / nơi lưu","Số lượng","Tình trạng","Ngày bàn giao","Người bàn giao","Người nhận","Ghi chú"])
    deliver=["Mã nguồn website (repository)","Cơ sở dữ liệu (dump/backup)","Tài khoản quản trị hệ thống",
             "Tài liệu hướng dẫn sử dụng","Tài liệu kỹ thuật / triển khai","Bộ Test Case (RTM)",
             "Bộ hồ sơ lỗi + bằng chứng","Bộ thiết kế (Figma/ảnh)","Tên miền, hosting, chứng chỉ SSL",
             "Tài khoản dịch vụ bên thứ ba"]
    for i,d in enumerate(deliver):
        r=line([i+1,d,"","","","","","",""],i,inputs=(2,3,5,6,7,8),nfmts={5:"dd/mm/yyyy"},left=(1,2,8))
        n[f"E{r}"].fill=Fill(INPUT); dv_bg.add(n[f"E{r}"])
    gap()

    # ---------- XV. BẢO HÀNH & HỖ TRỢ ----------
    sec("XV. BẢO HÀNH VÀ HỖ TRỢ SAU NGHIỆM THU")
    kv([("Thời gian bảo hành",""),("Từ ngày","","dd/mm/yyyy"),("Đến ngày","","dd/mm/yyyy")])
    kv([("Phạm vi bảo hành",""),("Kênh tiếp nhận lỗi",""),("Thời gian phản hồi (SLA)","")])
    kv([("Thời gian khắc phục Critical",""),("Thời gian khắc phục High",""),("Chi phí ngoài bảo hành","")])
    gap()

    # ---------- XVI. KẾT LUẬN NGHIỆM THU ----------
    sec("XVI. KẾT LUẬN NGHIỆM THU")
    head(["Nội dung","Kết quả","","Diễn giải","","","","",""])
    auto=(f'=IF(COUNTIF(E{crit_start}:E{crit_end},"Chưa đạt")>0,"KHÔNG ĐẠT — còn "&COUNTIF(E{crit_start}:E{crit_end},"Chưa đạt")&" tiêu chí chưa đạt",'
          f'IF(COUNTIF(E{crit_start}:E{crit_end},"Chưa chốt ngưỡng")+COUNTBLANK(E{crit_start}:E{crit_end})>0,'
          f'"CHƯA ĐỦ CĂN CỨ — còn tiêu chí chưa có ngưỡng hoặc chưa đánh giá","ĐẠT — toàn bộ tiêu chí bắt buộc đã thỏa mãn"))')
    r=row[0]
    cell(n,f"A{r}","Đề xuất tự động (theo tiêu chí mục V)",F(10,True),Fill(WHITE),Al("left"),border=box)
    merge(n,f"B{r}:I{r}",auto,font=F(11,True,RED),fill=Fill(RED_BG),al=Al("left"))
    for j in range(1,9): n[f"{get_column_letter(1+j)}{r}"].border=box
    n.row_dimensions[r].height=28; row[0]+=1
    r=row[0]
    cell(n,f"A{r}","Số hạng mục Đạt / Không đạt (mục IX)",F(10,True),Fill(WHITE),Al("left"),border=box)
    merge(n,f"B{r}:I{r}",f'="Đạt: "&COUNTIF(H{cat_start}:H{cat_end},"Đạt")&"  |  Đạt có điều kiện: "&COUNTIF(H{cat_start}:H{cat_end},"Đạt có điều kiện")&"  |  Không đạt: "&COUNTIF(H{cat_start}:H{cat_end},"Không đạt")&"  |  Chưa kết luận: "&COUNTBLANK(H{cat_start}:H{cat_end})',
          font=F(10.5),fill=Fill(WHITE),al=Al("left"))
    for j in range(1,9): n[f"{get_column_letter(1+j)}{r}"].border=box
    n.row_dimensions[r].height=24; row[0]+=1
    r=row[0]
    cell(n,"A"+str(r),"KẾT LUẬN CHÍNH THỨC (người duyệt chọn)",F(10.5,True,WHITE),Fill(NAVY),Al("left"),border=box)
    merge(n,f"B{r}:D{r}","",font=F(12,True),fill=Fill(INPUT),al=Al("center")); dv_kl.add(n[f"B{r}"])
    cell(n,f"E{r}","Ngày kết luận",F(10,True,WHITE),Fill(BLUE),Al("center"),border=box)
    merge(n,f"F{r}:G{r}","",font=F(10.5),fill=Fill(INPUT),al=Al("center")); n[f"F{r}"].number_format="dd/mm/yyyy"
    cell(n,f"H{r}","Hiệu lực từ",F(10,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(n,f"I{r}","",F(10.5),Fill(INPUT),Al("center"),border=box,nfmt="dd/mm/yyyy")
    for j in range(1,9): n[f"{get_column_letter(1+j)}{r}"].border=box
    n.row_dimensions[r].height=30; row[0]+=1
    r=row[0]
    cell(n,f"A{r}","Lý do / điều kiện kèm theo",F(10,True),Fill(WHITE),Al("left"),border=box)
    merge(n,f"B{r}:I{r}","",font=F(10),fill=Fill(INPUT),al=Al("left",wrap=True))
    for j in range(1,9): n[f"{get_column_letter(1+j)}{r}"].border=box
    n.row_dimensions[r].height=46; row[0]+=1
    gap()

    # ---------- XVII. VIỆC CẦN LÀM ----------
    sec("XVII. VIỆC CẦN LÀM ĐỂ HOÀN TẤT NGHIỆM THU")
    head(["STT","Việc cần làm","Đầu ra / điều kiện xong","Người phụ trách","Hạn hoàn thành","Mức ưu tiên","Trạng thái","Bằng chứng","Ghi chú"])
    todos=["Dev sửa dứt điểm toàn bộ lỗi Critical/High đang mở",
           "QA retest và chuyển trạng thái lỗi sang Verified/Closed",
           "Bổ sung link bằng chứng (ảnh/video) cho các mục còn thiếu",
           "Chốt danh sách DS/YC thuộc phạm vi nghiệm thu giai đoạn này",
           "Chốt ngưỡng tiêu chí nghiệm thu tại mục V với Bên A",
           "Đo và điền kết quả kiểm thử phi chức năng (mục X)",
           "Hoàn tất bàn giao tài liệu và tài khoản (mục XIV)",
           "Người duyệt ký kết luận nghiệm thu (mục XVI, XVIII)"]
    for i,t2 in enumerate(todos):
        line([i+1,t2,"","","","","","",""],i,inputs=(2,3,4,5,6,7,8),nfmts={4:"dd/mm/yyyy"},left=(1,2,7,8))
    gap()

    # ---------- XVIII. XÁC NHẬN CỦA CÁC BÊN ----------
    sec("XVIII. XÁC NHẬN CỦA CÁC BÊN")
    note("Biên bản được lập thành 02 bản có giá trị pháp lý như nhau, mỗi bên giữ 01 bản.",h=20)
    head(["Vai trò","Họ và tên","Chức vụ","Ý kiến / ghi chú","","Ngày ký","Chữ ký","",""],h=26)
    signs=["ĐẠI DIỆN BÊN A (Chủ đầu tư)","ĐẠI DIỆN BÊN B (Đơn vị thực hiện)","QUẢN LÝ DỰ ÁN","QA / TESTER","NGƯỜI DUYỆT NGHIỆM THU"]
    for i,s in enumerate(signs):
        r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
        cell(n,f"A{r}",s,F(10,True),bg,Al("left",wrap=True),border=box)
        for cc in ("B","C"): cell(n,f"{cc}{r}","",F(10),Fill(INPUT),Al("center"),border=box)
        merge(n,f"D{r}:E{r}","",font=F(10),fill=Fill(INPUT),al=Al("left",wrap=True))
        for cc in ("D","E"): n[f"{cc}{r}"].border=box
        cell(n,f"F{r}","",F(10),Fill(INPUT),Al("center"),border=box,nfmt="dd/mm/yyyy")
        merge(n,f"G{r}:I{r}","",font=F(10),fill=Fill(WHITE),al=Al("center"))
        for cc in ("G","H","I"): n[f"{cc}{r}"].border=box
        n.row_dimensions[r].height=56; row[0]+=1
    gap()

    # ---------- PHỤ LỤC ----------
    sec("PHỤ LỤC KÈM THEO")
    head(["Mã","Tên phụ lục","Nguồn / đường dẫn","Đính kèm?","Ghi chú","","","",""])
    apps=[("PL-01","Danh sách trang (DS) và trạng thái kiểm thử","File QA — sheet 02_UI_UX"),
          ("PL-02","Danh sách tính năng (YC) và trạng thái kiểm thử","File QA — sheet 03_Tính_năng"),
          ("PL-03","Danh sách lỗi đầy đủ kèm bằng chứng","File QA — sheet 05_Lỗi_Tester"),
          ("PL-04","Ma trận truy vết yêu cầu — Test Case (RTM)","File QA — sheet 06_RTM_Dev_Test"),
          ("PL-05","Nhật ký công việc kiểm thử theo ngày","File QA — sheet 01_Cong_viec_ngay"),
          ("PL-06","Báo cáo tiến độ theo tháng / tuần","File này — sheet 01, 02"),
          ("PL-07","Kết quả đo hiệu năng và bảo mật","Cần bổ sung"),
          ("PL-08","Biên bản họp chốt phạm vi và tiêu chí","Cần bổ sung")]
    for i,(a,b,c) in enumerate(apps):
        r=line([a,b,c,"","","","","",""],i,inputs=(3,4),left=(0,1,2,4))
    n.sheet_view.showGridLines=False
    n.freeze_panes="A3"
    return n

n=build_acceptance()
t=build_targets()
for ws,clr in [(g,NAVY),(m,BLUE),(w,"2E7D32"),(n,"8B5E00")]: ws.sheet_properties.tabColor=clr
wb.active=wb.sheetnames.index("01_Báo_cáo_Tháng")
out="/tmp/BaoCao_BGD_TimViec123.xlsx"; wb.save(out); print("SAVED",out); print(wb.sheetnames)
