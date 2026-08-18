# -*- coding: utf-8 -*-
"""Báo cáo tiến độ Website TimViec123 cho BGD — đồng bộ live qua IMPORTRANGE.
Nguồn lỗi: 05_Lỗi_Tester. Database: file 'Hồ sơ'. Tháng & Tuần dùng chung layout."""
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
zc["A1"]="URL file làm việc"; zc["B1"]=WORK_URL; zc["A2"]="Năm"; zc["B2"]=2026
zc["A3"]="URL file Database"; zc["B3"]=DB_URL
for rr in (1,2,3): zc[f"A{rr}"].font=F(bold=True)
zc.column_dimensions["A"].width=20; zc.column_dimensions["B"].width=70
def src(name,u,rng):
    ws=wb.create_sheet(name); ws["A1"]=f'=IMPORTRANGE({u},"{rng}")'; ws.sheet_state="hidden"; return ws
src("z_Trang",UW(),"02_UI_UX!A3:N3")
src("z_TinhNang",UW(),"03_Tính_năng!A3:N3")
src("z_TrangFull",UW(),"02_UI_UX!A4:R68")       # header row1 -> QUERY: B tên, E trạng thái, I tổng lỗi, P ngày test
src("z_TNFull",UW(),"03_Tính_năng!A4:T92")      # QUERY: B tên, J trạng thái, K tổng lỗi, R ngày test
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
# period-filtered (dùng ô $N$4/$O$4 CỦA CHÍNH SHEET)
LN="$N$4"; LO="$O$4"
def foundp(): return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO})'
def sevp(x):  return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO},{L},"{x}")'
def devp(x):  return f'=COUNTIFS({Q},">="&{LN},{Q},"<="&{LO},{M},"*{x}*")'
def cvdone(): return f'=COUNTIFS(\'z_CV\'!$B:$B,">="&{LN},\'z_CV\'!$B:$B,"<="&{LO},\'z_CV\'!$O:$O,"Hoàn thành")'
def dbp(d):   return f'=COUNTIFS({d},">="&{LN},{d},"<="&{LO})'
def detail_query(rng,namec,bugc,datec,statc,fb):
    S1=f'select {namec}, {bugc}, {datec} where {datec} >= date \''
    S2=f'\' and {datec} <= date \''
    S3=(f'\' and ({statc} = \'Hoàn thành\' or {statc} = \'Kiểm tra lại\') order by {datec} desc '
        f"label {namec} '', {bugc} '', {datec} '' format {datec} 'dd/mm/yyyy', {bugc} '#,##0'")
    return ('=IFERROR(QUERY('+rng+',"'+S1+'"&TEXT($N$4,"yyyy-mm-dd")&"'+S2+
            '"&TEXT($O$4,"yyyy-mm-dd")&"'+S3+'",1),"'+fb+'")')

# ================= build_period (Tháng / Tuần dùng chung) =================
def build_period(sheetname, mode):
    s=wb.create_sheet(sheetname)
    for col,wd in {"A":34,"B":14,"C":13,"D":12,"E":12,"F":30}.items(): s.column_dimensions[col].width=wd
    s.column_dimensions["N"].hidden=True; s.column_dimensions["O"].hidden=True
    P = "tháng" if mode=="month" else "tuần"
    band(s,"A1:F1",f"BÁO CÁO TIẾN ĐỘ THEO {P.upper()}",big=True)
    merge(s,"A2:F2","Ô nền vàng là mục tiêu nhập tay theo kỳ. Số tổng quan là thời điểm hiện tại; chỉ số \""+("trong "+P)+"\" lọc theo kỳ đã chọn.",
          font=F(10,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
    if mode=="month":
        cell(s,"A4","Tháng báo cáo",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"C4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"D4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"E4","Ngày chốt",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"F4","=TODAY()",F(11,True),Fill(INPUT),Al("center"),border=box); s["F4"].number_format="dd/mm/yyyy"
        s["N4"]="=DATE(D4,B4,1)"; s["O4"]="=EOMONTH(N4,0)"
        dv=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); s.add_data_validation(dv); dv.add(s["B4"])
        selrow=5
    else:
        cell(s,"A4","Tháng",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"B4",8,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"C4","Tuần",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"D4",2,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        cell(s,"E4","Năm",F(11,True,WHITE),Fill(BLUE),Al("center"),border=box); cell(s,"F4",2026,F(12,True,NAVY),Fill(INPUT),Al("center"),border=box)
        s["N4"]="=DATE(F4,B4,(D4-1)*7+1)"; s["O4"]="=MIN(EOMONTH(DATE(F4,B4,1),0),DATE(F4,B4,D4*7))"
        d1=DataValidation(type="list",formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); s.add_data_validation(d1); d1.add(s["B4"])
        d2=DataValidation(type="list",formula1='"1,2,3,4,5"'); s.add_data_validation(d2); d2.add(s["D4"])
        cell(s,"A5","Từ ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); cell(s,"B5","=N4",F(10.5,True),Fill(WHITE),Al("center"),border=box); s["B5"].number_format="dd/mm/yyyy"
        cell(s,"C5","Đến ngày",F(10.5,True),Fill(LBLUE),Al("center"),border=box); merge(s,"D5:E5","=O4",font=F(10.5,True),fill=Fill(WHITE),al=Al("center")); s["D5"].number_format="dd/mm/yyyy"
        cell(s,"F5",f'=IF(ISERROR({T_TOT}),"⚠ Chưa kết nối","✔ Đã kết nối")',F(10,True,GREEN),Fill(WHITE),Al("center"),border=box)
        selrow=6
    # ---- KPI table ----
    def kpih(r):
        for i,h in enumerate(["KPI / CHỈ SỐ","Mục tiêu kỳ","Kết quả","Tỷ lệ đạt","Đạt KPI?","Ghi chú"]):
            c=get_column_letter(1+i); cell(s,f"{c}{r}",h,F(10.5,True,WHITE),Fill(NAVY),Al("center" if i else "left"),border=box)
        s.row_dimensions[r].height=22
    def sect(r,txt): merge(s,f"A{r}:F{r}",txt,font=F(11,True,NAVY),fill=Fill(LBLUE),al=Al("left"),border=True); s.row_dimensions[r].height=19
    def kpi(r,name,result,note,zebra=False):
        bg=Fill(ZEBRA) if zebra else Fill(WHITE)
        cell(s,f"A{r}",name,F(10.5),bg,Al("left",wrap=True),border=box)
        cell(s,f"B{r}",None,F(10.5,True,NAVY),Fill(INPUT),Al("center"),border=box)   # mục tiêu nhập tay
        cell(s,f"C{r}",result,F(11,True),bg,Al("center"),border=box,nfmt=INT)
        cell(s,f"D{r}",f'=IFERROR(C{r}/B{r},"")',F(10.5),bg,Al("center"),border=box,nfmt=PCT)
        cell(s,f"E{r}",f'=IF(ISNUMBER(B{r}),IF(C{r}>=B{r},"Đạt","Chưa đạt"),"Theo dõi")',F(10.5,True),bg,Al("center"),border=box)
        cell(s,f"F{r}",note,F(9.5,color=GREY),bg,Al("left",wrap=True),border=box)
    r=selrow+1; kpih(r); r+=1
    sect(r,"I. KIỂM TRA WEBSITE"); r+=1
    kpi(r,"1. Trang hoàn thành",f"={T_DONE}",f"Hoàn thành / tổng {'{}'.format('')}trang"); r+=1
    kpi(r,"2. Tính năng hoàn thành",f"={F_DONE}","Hoàn thành / 88 tính năng",zebra=True); r+=1
    kpi(r,f"3. Công việc hoàn thành trong {P}",cvdone(),"Nhật ký công việc theo ngày"); r+=1
    kpi(r,f"4. Lỗi phát hiện trong {P}",foundp(),"Theo Ngày phát hiện (05_Lỗi_Tester)",zebra=True); r+=1
    sect(r,"II. HOÀN THIỆN WEBSITE"); r+=1
    kpi(r,"1. Lỗi đã xử lý (lũy kế)",RESOLVED,"Fixed + Verified + Closed"); r+=1
    kpi(r,"2. Số Trang đã xử lý",f"={T_TOT}-{T_UN}","Đã kiểm thử = Tổng − Chưa kiểm tra",zebra=True); r+=1
    kpi(r,"3. Số Tính năng đã xử lý",f"={F_TOT}-{F_UN}","Đã kiểm thử = Tổng − Chưa kiểm tra"); r+=1
    sect(r,f"III. DATABASE (trong {P})"); r+=1
    kpi(r,"1. DB Doanh nghiệp đã thêm",dbp(DBDN_D),"Theo Ngày Đăng/Tạo"); r+=1
    kpi(r,"2. DB HR đã thêm",dbp(DBHR_D),"Theo Ngày Đăng/Tạo",zebra=True); r+=1
    kpi(r,"3. DB Ứng viên đã thêm",f"=COUNTA({DBUV})","Theo tổng (chưa có cột ngày nhập)"); r+=1
    kpi(r,"4. Tin tuyển dụng đã đăng",dbp(DBTIN_D),"Theo Ngày viết",zebra=True); r+=1
    r+=1
    # ---- TỔNG QUAN HIỆN TẠI ----
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
    # ---- TÌNH TRẠNG XỬ LÝ LỖI ----
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
    # ---- PHÂN BỐ LỖI TRONG KỲ ----
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
    # ---- THANH TIẾN ĐỘ ----
    merge(s,f"A{r}:F{r}","THANH TIẾN ĐỘ HOÀN THÀNH",font=F(11,True,WHITE),fill=Fill(BLUE),al=Al("left"),border=True); r+=1
    cell(s,f"A{r}","Trang",F(10.5,True),Fill(WHITE),Al("left"),border=box)
    merge(s,f"B{r}:F{r}",f'=REPT("█",ROUND({T_PROG}*30,0))&" "&TEXT({T_PROG},"0.0%")',font=F(11,True,BLUE),fill=Fill(WHITE),al=Al("left"))
    for cc in "BCDEF": s[f"{cc}{r}"].border=box
    r+=1
    cell(s,f"A{r}","Tính năng",F(10.5,True),Fill(ZEBRA),Al("left"),border=box)
    merge(s,f"B{r}:F{r}",f'=REPT("█",ROUND({F_PROG}*30,0))&" "&TEXT({F_PROG},"0.0%")',font=F(11,True,GREEN),fill=Fill(ZEBRA),al=Al("left"))
    for cc in "BCDEF": s[f"{cc}{r}"].border=box
    r+=2
    # ---- CHI TIẾT theo Trang / Tính năng (test trong kỳ) ----
    band(s,f"A{r}:F{r}",f"CHI TIẾT: TRANG & TÍNH NĂNG ĐÃ TEST TRONG {P.upper()}"); r+=1
    merge(s,f"A{r}:F{r}","Liệt kê các trang/tính năng có Ngày test gần nhất trong kỳ (trạng thái Hoàn thành hoặc Kiểm tra lại).",
          font=F(9.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left")); r+=1
    # Trang block
    cell(s,f"A{r}","TRANG đã test trong kỳ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
    cell(s,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"C{r}","Ngày test/retest",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box); r+=1
    cell(s,f"A{r}",detail_query("'z_TrangFull'!$A$1:$R$65","B","I","P","E","(không có trang test trong kỳ)"),F(10),Fill(WHITE),Al("left"))
    for rr2 in range(r,r+64):
        for cc in "ABC": s[f"{cc}{rr2}"].border=box
        s[f"C{rr2}"].number_format="dd/mm/yyyy"; s[f"B{rr2}"].number_format=INT
    r+=65
    cell(s,f"A{r}","TÍNH NĂNG đã test trong kỳ",F(10.5,True,WHITE),Fill(BLUE),Al("left"),border=box)
    cell(s,f"B{r}","Số lỗi",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box)
    cell(s,f"C{r}","Ngày test/retest",F(10.5,True,WHITE),Fill(BLUE),Al("center"),border=box); r+=1
    cell(s,f"A{r}",detail_query("'z_TNFull'!$A$1:$T$89","B","K","R","J","(không có tính năng test trong kỳ)"),F(10),Fill(WHITE),Al("left"))
    for rr2 in range(r,r+90):
        for cc in "ABC": s[f"{cc}{rr2}"].border=box
        s[f"C{rr2}"].number_format="dd/mm/yyyy"; s[f"B{rr2}"].number_format=INT
    s.sheet_view.showGridLines=False
    return s

# ============ 00_Hướng_dẫn ============
g=wb.create_sheet("00_Hướng_dẫn")
for col,w in {"A":26,"B":30,"C":28,"D":20,"E":20,"F":16}.items(): g.column_dimensions[col].width=w
band(g,"A1:F1","BÁO CÁO TIẾN ĐỘ WEBSITE TIMVIEC123 — DÀNH CHO BAN GIÁM ĐỐC",big=True)
merge(g,"A2:F2","Số liệu đồng bộ live từ file làm việc QA và file Database. BGD chỉ xem, chọn kỳ ở sheet Tháng/Tuần.",
      font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left")); g.row_dimensions[2].height=20
merge(g,"A4:B4","NGƯỜI DÙNG",font=F(11,True,WHITE),fill=Fill(BLUE)); merge(g,"C4:F4","CÁCH SỬ DỤNG",font=F(11,True,WHITE),fill=Fill(BLUE))
r=5
for who,how in [("Ban Giám đốc","Xem tiến độ trang/tính năng, tình trạng lỗi, database và nghiệm thu. Chọn Tháng/Tuần để lọc kỳ."),
                ("Tester","Cập nhật dữ liệu tại file làm việc; nhập Mục tiêu kỳ; báo cáo tự đồng bộ."),
                ("Dev","Cập nhật xử lý lỗi tại 06_RTM_Dev_Test của file làm việc.")]:
    merge(g,f"A{r}:B{r}",who,font=F(11,True),fill=Fill(LBLUE),al=Al("left")); merge(g,f"C{r}:F{r}",how,font=F(10.5),al=Al("left",wrap=True)); g.row_dimensions[r].height=30; r+=1
r+=1; band(g,f"A{r}:F{r}","NGUYÊN TẮC TÍNH SỐ"); r+=1
for i,p in enumerate([
 "Tiến độ trang = số trang Hoàn thành / tổng số trang (hiện 63). Tính năng / 88.",
 "Số lỗi lấy toàn bộ từ 05_Lỗi_Tester. \"Đang mở\" = Open + Chưa gửi Dev; \"Đã xử lý\" = Fixed + Verified + Closed.",
 "Mục tiêu kỳ do BGD/Tester tự nhập (ô nền vàng), không cố định.",
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

# build month & week (identical layout)
m=build_period("01_Báo_cáo_Tháng","month")
w=build_period("02_Báo_cáo_Tuần","week")

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
merge(n,f"C{r}:H{r}","Chưa nghiệm thu",font=F(10.5,italic=True,color=GREY),fill=Fill(WHITE),al=Al("left"))
for cc in "CDEFGH": n[f"{cc}{r}"].border=box
r+=1
merge(n,f"A{r}:H{r}","Điều kiện khuyến nghị: không còn Critical/High chưa được phê duyệt; phạm vi và bằng chứng đầy đủ; ngoại lệ phải có người duyệt và lý do rõ ràng.",
      font=F(9.5,color=GREY),fill=Fill(ZEBRA),al=Al("left",wrap=True)); n.row_dimensions[r].height=34
n.sheet_view.showGridLines=False

for ws,clr in [(g,NAVY),(m,BLUE),(w,"2E7D32"),(n,"8B5E00")]: ws.sheet_properties.tabColor=clr
wb.active=wb.sheetnames.index("01_Báo_cáo_Tháng")
out="/tmp/BaoCao_BGD_TimViec123.xlsx"; wb.save(out); print("SAVED",out); print(wb.sheetnames)
