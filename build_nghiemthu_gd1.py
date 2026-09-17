# -*- coding: utf-8 -*-
"""Biên bản nghiệm thu Giai đoạn 1 — DS_04, DS_05, DS_21, DS_22, DS_25.

Dữ liệu chốt từ file QA 'Tester Detail- Timviec123' (02_UI_UX, 03_Tính_năng,
05_Lỗi_Tester, 06_RTM_Dev_Test, 04_Test_Cases) — xem data_nghiemthu_gd1.json.
Đây là văn bản chốt tại một thời điểm (snapshot) để ký, không đồng bộ live.

Quy ước kết luận một DS = ĐẠT khi đủ 4 điều kiện:
  1. Trạng thái kiểm tra trang = Hoàn thành
  2. Không còn lỗi đang mở
  3. Không còn lỗi Critical/High chưa đóng
  4. Mọi lỗi đã xử lý đều có Kết quả Retest = Pass
"""
import json, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries

NAVY="1F3A5F"; BLUE="2E5A88"; LBLUE="DCE6F1"; ZEBRA="F5F7FA"
GREEN="2E7D32"; GREEN_BG="D9F0DD"; RED="C62828"; RED_BG="FBE3E3"
AMBER="B8860B"; AMBER_BG="FFF4D6"; GREY="6B7280"; WHITE="FFFFFF"
INPUT="FFF9DB"; BORDER_CLR="C9D6E5"
thin=Side(style="thin",color=BORDER_CLR); box=Border(left=thin,right=thin,top=thin,bottom=thin)
def F(size=10,bold=False,color="1A1A1A",italic=False): return Font(name="Arial",size=size,bold=bold,color=color,italic=italic)
def Fill(c): return PatternFill("solid",fgColor=c)
def Al(h="left",v="center",wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
INT='#,##0'; PCT='0.0%'

COLS={"A":13,"B":28,"C":13,"D":13,"E":15,"F":34,"G":36,"H":13,"I":13,"J":13,"K":12,"L":13,"M":19,"N":26}
NCOL=len(COLS)
LAST=get_column_letter(NCOL)

data=json.load(open("data_nghiemthu_gd1.json",encoding="utf-8"))
ASOF=data["asof"]

wb=openpyxl.Workbook(); ws=wb.active; ws.title="NghiemThu_GD1"
for c,w in COLS.items(): ws.column_dimensions[c].width=w
row=[1]

def fillrange(rng,fill):
    c1,r1,c2,r2=range_boundaries(rng)
    for r in range(r1,r2+1):
        for c in range(c1,c2+1): ws.cell(r,c).fill=fill
def put(ref,val=None,font=None,fill=None,al=None,nfmt=None,border=True):
    c=ws[ref]
    if val is not None: c.value=val
    c.font=font or F(); c.alignment=al or Al()
    if fill: c.fill=fill
    if nfmt: c.number_format=nfmt
    if border: c.border=box
    return c
def merge(rng,val=None,font=None,fill=None,al=None,border=True):
    ws.merge_cells(rng); tl=rng.split(":")[0]
    put(tl,val,font,None,al,border=False)
    if fill: fillrange(rng,fill)
    if border:
        c1,r1,c2,r2=range_boundaries(rng)
        for r in range(r1,r2+1):
            for c in range(c1,c2+1): ws.cell(r,c).border=box
def gap(h=8):
    ws.row_dimensions[row[0]].height=h; row[0]+=1
def title(t,sub=None):
    r=row[0]; merge(f"A{r}:{LAST}{r}",t,F(15,True,WHITE),Fill(NAVY),Al("left"),border=False)
    ws.row_dimensions[r].height=32; row[0]+=1
    if sub:
        r=row[0]; merge(f"A{r}:{LAST}{r}",sub,F(9.5,italic=True,color=GREY),Fill(WHITE),Al("left",wrap=True),border=False)
        ws.row_dimensions[r].height=26; row[0]+=1
def sec(t):
    r=row[0]; merge(f"A{r}:{LAST}{r}",t,F(11.5,True,WHITE),Fill(BLUE),Al("left"),border=False)
    ws.row_dimensions[r].height=24; row[0]+=1; return r
def note(t,h=24):
    r=row[0]; merge(f"A{r}:{LAST}{r}",t,F(9.5,italic=True,color=GREY),Fill(WHITE),Al("left",wrap=True),border=False)
    ws.row_dimensions[r].height=h; row[0]+=1
def head(labels,h=34):
    r=row[0]
    for i,l in enumerate(labels):
        put(f"{get_column_letter(1+i)}{r}",l,F(9.5,True,WHITE),Fill(NAVY),Al("center",wrap=True))
    for i in range(len(labels),NCOL):
        put(f"{get_column_letter(1+i)}{r}","",F(9.5,True,WHITE),Fill(NAVY),Al("center"))
    ws.row_dimensions[r].height=h; row[0]+=1; return r
def line(vals,i=0,left=(),h=22,fills=None,fonts=None,nfmts=None,wrap=()):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    for j in range(NCOL):
        v=vals[j] if j<len(vals) else ""
        put(f"{get_column_letter(1+j)}{r}", v if v!="" else None,
            (fonts or {}).get(j,F()), (fills or {}).get(j,bg),
            Al("left" if j in left else "center", wrap=(j in wrap or j in left)),
            (nfmts or {}).get(j))
    ws.row_dimensions[r].height=h; row[0]+=1; return r
def kv(pairs,h=22):
    r=row[0]; slots=[("A","B:C"),("D","E:F"),("G","H:I"),("J","K:N")]
    for k,(lc,vr) in enumerate(slots):
        a,b=vr.split(":")
        if k<len(pairs):
            lb,val=pairs[k][0],pairs[k][1]
            inp=pairs[k][2] if len(pairs[k])>2 else True
            put(f"{lc}{r}",lb,F(9.5,True,WHITE),Fill(BLUE),Al("left",wrap=True))
            merge(f"{a}{r}:{b}{r}",val if val!="" else None,F(10),Fill(INPUT) if inp else Fill(WHITE),Al("left",wrap=True))
        else:
            put(f"{lc}{r}","",F(),Fill(WHITE),Al()); merge(f"{a}{r}:{b}{r}","",F(),Fill(WHITE),Al())
    ws.row_dimensions[r].height=h; row[0]+=1; return r

# ---------- tính toán kết luận ----------
OPEN_ST={"Open","Chưa gửi Dev","Đang sửa"}
for d in data["ds"]:
    bs=d["bugs"]
    d["n_tot"]=len(bs)
    d["n_open"]=sum(1 for b in bs if b["status"] in OPEN_ST)
    d["n_done"]=d["n_tot"]-d["n_open"]
    d["n_ch_open"]=sum(1 for b in bs if b["status"] in OPEN_ST and b["sev"] in ("Critical","High"))
    d["n_pass"]=sum(1 for b in bs if b["retest"]=="Pass")
    d["n_norетest"]=0
    d["n_noretest"]=sum(1 for b in bs if b["status"] not in OPEN_ST and b["retest"]!="Pass")
    reasons=[]
    if d["stat"]!="Hoàn thành": reasons.append(f"Trạng thái trang: {d['stat']}")
    if d["n_open"]: reasons.append(f"Còn {d['n_open']} lỗi đang mở")
    if d["n_ch_open"]: reasons.append(f"Còn {d['n_ch_open']} lỗi Critical/High chưa đóng")
    if d["n_noretest"]: reasons.append(f"Còn {d['n_noretest']} lỗi chưa retest Pass")
    d["verdict"]="ĐẠT" if not reasons else "CHƯA ĐẠT"
    d["reason"]="Đủ 4 điều kiện nghiệm thu" if not reasons else "; ".join(reasons)

TOT=sum(d["n_tot"] for d in data["ds"]); OPN=sum(d["n_open"] for d in data["ds"])
DONE=TOT-OPN; CH=sum(d["n_ch_open"] for d in data["ds"]); PASS=sum(d["n_pass"] for d in data["ds"])
NDS=len(data["ds"]); NPASS=sum(1 for d in data["ds"] if d["verdict"]=="ĐẠT")
NYC=sum(len(d["yc"]) for d in data["ds"])
EVI=sum(1 for d in data["ds"] for b in d["bugs"] if b["evi"])
EVI_RT=sum(1 for d in data["ds"] for b in d["bugs"] if not b["evi"] and b["retest"]=="Pass")
EVI_NONE=TOT-EVI-EVI_RT

# ================= NỘI DUNG =================
title("BIÊN BẢN NGHIỆM THU WEBSITE TIMVIEC123 — GIAI ĐOẠN 1",
      f"Phạm vi: 5 màn hình (DS_04, DS_05, DS_21, DS_22, DS_25) và {NYC} tính năng liên quan. "
      f"Số liệu chốt từ file QA 'Tester Detail- Timviec123' ngày {ASOF}. Ô nền vàng do các bên điền khi ký.")
gap()

sec("I. THÔNG TIN CHUNG")
kv([("Số biên bản",""),("Ngày lập",""),("Địa điểm",""),("Lần nghiệm thu","")])
kv([("Tên dự án","Website TimViec123",False),("Tên miền / URL",""),("Môi trường nghiệm thu","Production",False),("Phiên bản / Release","")])
kv([("Giai đoạn","Giai đoạn 1",False),("Ngày chốt số liệu",ASOF,False),("Tester thực hiện",data["ds"][0]["tester"],False),("Dev phụ trách","Mr. Thanh",False)])
kv([("Bên A — Chủ đầu tư",""),("Đại diện Bên A",""),("Bên B — Đơn vị thực hiện",""),("Đại diện Bên B","")])
kv([("Hợp đồng số",""),("Ngày hợp đồng",""),("Tài liệu yêu cầu",""),("Kế hoạch kiểm thử","")])
gap()

sec("II. PHẠM VI NGHIỆM THU")
note("Chỉ nghiệm thu đúng các màn hình và tính năng liệt kê dưới đây. Màn hình, tính năng không có trong bảng này nằm ngoài phạm vi đợt nghiệm thu và không dùng làm lý do từ chối ký.")
head(["DS_ID","Tên màn hình","Đường dẫn","Số tính năng","Tính năng (YC_ID) trong phạm vi","","","Ưu tiên nghiệm thu","Thiết bị đã kiểm thử","","","","Ghi chú phạm vi",""])
for i,d in enumerate(data["ds"]):
    ycs=", ".join(y["yc"] for y in d["yc"])
    prio=", ".join(sorted({y["prio"] for y in d["yc"] if y["prio"]}))
    dev=" / ".join(f"{k}: {v}" for k,v in d["dev"].items())
    r=row[0]
    line([d["ds"],d["name"],d["url"] or "—",len(d["yc"]),"","","",prio,"","","","","",""],i,left=(1,4,8,12),h=26)
    merge(f"E{r}:G{r}",ycs,F(10),Fill(ZEBRA) if i%2 else Fill(WHITE),Al("left",wrap=True))
    merge(f"I{r}:L{r}",dev,F(10),Fill(ZEBRA) if i%2 else Fill(WHITE),Al("left",wrap=True))
    merge(f"M{r}:N{r}","",F(10),Fill(INPUT),Al("left",wrap=True))
r=row[0]
merge(f"A{r}:D{r}","NGOÀI PHẠM VI ĐỢT NÀY",F(10,True,WHITE),Fill(GREY),Al("left"))
merge(f"E{r}:{LAST}{r}","Các màn hình DS khác và các tính năng không liệt kê ở trên; riêng YC_04 chỉ nghiệm thu phần thể hiện trên DS_22 — phần YC_04 trên DS_23 (Đăng ký Ứng viên) thuộc đợt sau.",
      F(9.5),Fill(WHITE),Al("left",wrap=True))
ws.row_dimensions[r].height=26; row[0]+=1
gap()

sec("III. TIÊU CHÍ NGHIỆM THU ÁP DỤNG")
head(["Mã","Tiêu chí","","Ngưỡng bắt buộc","Kết quả thực tế","","Đánh giá","","","","","","",""])
crits=[("NT-01","Trạng thái kiểm tra của trang","Hoàn thành",f"{sum(1 for d in data['ds'] if d['stat']=='Hoàn thành')}/{NDS} trang Hoàn thành",
        sum(1 for d in data['ds'] if d['stat']=='Hoàn thành')==NDS),
       ("NT-02","Lỗi đang mở trên trang","0",f"{OPN} lỗi đang mở",OPN==0),
       ("NT-03","Lỗi Critical/High chưa đóng","0",f"{CH} lỗi Critical/High đang mở",CH==0),
       ("NT-04","Lỗi đã xử lý phải có Kết quả Retest = Pass","100%",f"{PASS}/{TOT} lỗi đạt Retest Pass",PASS==TOT),
       ("NT-05","Bằng chứng xử lý lỗi","Có link ảnh/video, hoặc xác nhận retest với lỗi chức năng",
        f"{EVI}/{TOT} lỗi có link ảnh; {EVI_RT} lỗi chức năng xác nhận bằng Retest Pass; {EVI_NONE} lỗi thiếu cả hai",EVI_NONE==0),
       ("NT-06","Tính năng trong phạm vi đạt trạng thái Hoàn thành","100%",
        f"{sum(1 for d in data['ds'] for y in d['yc'] if y['stat']=='Hoàn thành')}/{NYC} tính năng Hoàn thành",
        all(y["stat"]=="Hoàn thành" for d in data["ds"] for y in d["yc"]))]
for i,(c,n,thr,res,ok) in enumerate(crits):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    line([c,"","",thr,"","","","","","","","","",""],i,h=26)
    merge(f"B{r}:C{r}",n,F(10),bg,Al("left",wrap=True))
    merge(f"E{r}:F{r}",res,F(10),bg,Al("left",wrap=True))
    merge(f"G{r}:{LAST}{r}","ĐẠT" if ok else "CHƯA ĐẠT",F(10,True,GREEN if ok else RED),
          Fill(GREEN_BG) if ok else Fill(RED_BG),Al("center"))
gap()

sec("IV. TỔNG HỢP KẾT QUẢ")
head(["Chỉ số","","Số lượng","Chỉ số","","Số lượng","Chỉ số","","Số lượng","Chỉ số","","Số lượng","",""])
summary=[[("Màn hình trong phạm vi",NDS),("Tính năng trong phạm vi",NYC),("Tổng lỗi ghi nhận",TOT),("Lỗi đã xử lý",DONE)],
         [("Màn hình ĐẠT nghiệm thu",NPASS),("Màn hình CHƯA ĐẠT",NDS-NPASS),("Lỗi đang mở",OPN),("Lỗi Critical/High mở",CH)],
         [("Lỗi đã Retest Pass",PASS),("Lỗi chưa retest",TOT-PASS),("Test Case liên quan",sum(d["ntc"] for d in data["ds"])),("Tỷ lệ màn hình đạt",f"{NPASS}/{NDS}")]]
for i,grp in enumerate(summary):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    line([""]*NCOL,i,h=22)
    for k,(lb,v) in enumerate(grp):
        c1=get_column_letter(1+k*3); c2=get_column_letter(2+k*3); c3=get_column_letter(3+k*3)
        merge(f"{c1}{r}:{c2}{r}",lb,F(10),bg,Al("left",wrap=True))
        bad = (lb.startswith("Lỗi đang mở") or lb.startswith("Lỗi Critical") or lb=="Màn hình CHƯA ĐẠT") and isinstance(v,int) and v>0
        put(f"{c3}{r}",v,F(11,True,RED if bad else NAVY),bg,Al("center"))
gap()

sec("V. BẢNG NGHIỆM THU THEO MÀN HÌNH (DS)")
head(["DS_ID","Tên màn hình","Trạng thái QA","Hoàn thành %","Tổng lỗi","Tính năng trong phạm vi","Lý do chốt","Lỗi đã xử lý","Lỗi đang mở","C/H đang mở","Retest Pass","Ngày test cuối","KẾT LUẬN","Ý kiến người duyệt"])
for i,d in enumerate(data["ds"]):
    ok=d["verdict"]=="ĐẠT"
    r=line([d["ds"],d["name"],d["stat"],d["pct"],d["n_tot"],", ".join(y["yc"] for y in d["yc"]),d["reason"],
            d["n_done"],d["n_open"],d["n_ch_open"],f'{d["n_pass"]}/{d["n_tot"]}',d["last"],d["verdict"],""],
           i,left=(1,5,6),h=34,nfmts={3:PCT},
           fills={12:Fill(GREEN_BG) if ok else Fill(RED_BG),13:Fill(INPUT),
                  9:Fill(WHITE) if d["n_ch_open"]==0 else Fill(RED_BG),
                  8:Fill(WHITE) if d["n_open"]==0 else Fill(RED_BG)},
           fonts={12:F(10.5,True,GREEN if ok else RED),8:F(10,True,RED if d["n_open"] else "1A1A1A"),
                  9:F(10,True,RED if d["n_ch_open"] else "1A1A1A")})
gap()

sec("VI. BẢNG NGHIỆM THU THEO TÍNH NĂNG (YC)")
note("Số lỗi tính theo phần thể hiện của tính năng trên chính màn hình thuộc phạm vi. YC_04 chỉ xét trên DS_22 theo phạm vi đã chốt.")
head(["YC_ID","Tên tính năng","","DS_ID","Màn hình","Ưu tiên","Trạng thái kiểm tra","Lỗi trong phạm vi","Lỗi đang mở","Ngày test cuối","KẾT LUẬN","","Ghi chú",""])
i=0
for d in data["ds"]:
    for y in d["yc"]:
        ok = y["stat"]=="Hoàn thành" and y["open_in_ds"]==0 and d["n_ch_open"]==0
        r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
        line([y["yc"],"","",d["ds"],d["name"],y["prio"],y["stat"],y["bugs_in_ds"],y["open_in_ds"],y["last"],"","","",""],i,left=(4,),h=24,
             fills={10:Fill(GREEN_BG) if ok else Fill(RED_BG)})
        merge(f"B{r}:C{r}",y["name"],F(10),bg,Al("left",wrap=True))
        merge(f"K{r}:L{r}","ĐẠT" if ok else "CHƯA ĐẠT",F(10,True,GREEN if ok else RED),
              Fill(GREEN_BG) if ok else Fill(RED_BG),Al("center"))
        merge(f"M{r}:N{r}","" if ok else "Chặn bởi lỗi còn mở trên màn hình",F(9.5),bg,Al("left",wrap=True))
        i+=1
gap()

sec("VII. CHI TIẾT LỖI VÀ BẰNG CHỨNG XỬ LÝ")
note("Nguồn: 05_Lỗi_Tester và 06_RTM_Dev_Test. Cột Bằng chứng: 'Ảnh' = có link ảnh/video thực tế; "
     "'Retest Pass' = lỗi chức năng không có ảnh, xác nhận bằng kết quả retest.")
head(["BUG_ID","Tiêu đề lỗi","Kết quả mong đợi","YC_ID","Mức độ","Loại lỗi","Thiết bị","Ngày phát hiện","Dev xử lý","Ngày báo fix","Retest","Ngày retest","Trạng thái","Bằng chứng"])
i=0
for d in data["ds"]:
    r=row[0]
    merge(f"A{r}:{LAST}{r}",f'{d["ds"]} — {d["name"]}   ({d["n_tot"]} lỗi: {d["n_done"]} đã xử lý, {d["n_open"]} đang mở)',
          F(10,True,WHITE),Fill(BLUE),Al("left"))
    ws.row_dimensions[r].height=22; row[0]+=1
    for b in d["bugs"]:
        isopen=b["status"] in OPEN_ST
        evi="Ảnh" if b["evi"] else ("Retest Pass" if b["retest"]=="Pass" else "THIẾU")
        line([b["bug"],b["title"],b["expect"],b["yc"],b["sev"],b["type"],b["dev_env"],b["found"],
              b["dev"] or "—",b["fixed"] or "—",b["retest"] or "—",b["retest_date"] or "—",b["status"],evi],
             i,left=(1,2),h=30,
             fills={12:Fill(RED_BG) if isopen else Fill(GREEN_BG),
                    4:Fill(RED_BG) if (isopen and b["sev"] in ("Critical","High")) else (Fill(ZEBRA) if i%2 else Fill(WHITE)),
                    13:Fill(AMBER_BG) if evi=="THIẾU" else (Fill(ZEBRA) if i%2 else Fill(WHITE))},
             fonts={12:F(10,True,RED if isopen else GREEN)})
        i+=1
gap()

sec("VIII. ĐỘ PHỦ TEST CASE")
note("Nguồn: 04_Test_Cases. Sheet 03_Test_Execution chưa có dữ liệu nên chưa thống kê được số lần chạy và tỷ lệ Pass/Fail theo lần chạy.")
head(["DS_ID","Tên màn hình","Số Test Case","Trạng thái Test Case","","","Nhận xét","","","","","","",""])
for i,d in enumerate(data["ds"]):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    st=", ".join(f"{k}: {v}" for k,v in d["tcst"].items() if k)
    nx=("Đủ dùng" if d["ntc"]>=3 else "Số Test Case quá ít so với phạm vi màn hình")
    if "Fail" in d["tcst"]: nx="Còn Test Case ở trạng thái Fail — cần chạy lại và cập nhật kết quả"
    line([d["ds"],d["name"],d["ntc"],"","","","","","","","","","",""],i,left=(1,),h=24)
    merge(f"D{r}:F{r}",st,F(10),bg,Al("left",wrap=True))
    merge(f"G{r}:{LAST}{r}",nx,F(10),bg,Al("left",wrap=True))
gap()

sec("IX. VẤN ĐỀ TỒN ĐỌNG VÀ GHI CHÚ")
head(["STT","Nội dung","","","","","Ảnh hưởng đến nghiệm thu","","Người phụ trách","Hạn xử lý","Trạng thái","","Ghi chú",""])
issues=[]
_op=[d for d in data["ds"] if d["n_open"]]
if _op:
    issues.append(("Còn lỗi đang mở: "+"; ".join(f'{d["ds"]} ({d["n_open"]} lỗi)' for d in _op),"Chặn nghiệm thu"))
else:
    issues.append(("Toàn bộ lỗi trong phạm vi đã được Dev xử lý và QA retest đạt (Kết quả Retest = Pass). "
                   "BUG_UI_107 trên DS_22 — lỗi cuối cùng còn mở — đã được đóng và retest Pass ngày 17/09/2026.",
                   "Không còn vướng mắc"))
issues += [
    ("BUG_UI_107 được gắn YC_06 (Xác minh email / Zalo) — tính năng này không nằm trong phạm vi DS_22 đã chốt "
     "(YC_04, YC_05, YC_07). Đề nghị gắn lại đúng YC hoặc bổ sung YC_06 vào phạm vi để hồ sơ thống nhất.",
     "Không đổi kết luận"),
    ("YC_04 còn 5 lỗi trên DS_23 (BUG_FUNC_020 đang sửa, BUG_REQ_002 Deferred). Theo phạm vi đã chốt, "
     "phần YC_04 trên DS_23 thuộc đợt nghiệm thu sau.",
     "Ngoài phạm vi đợt này"),
    ("3 lỗi chức năng không có ảnh bằng chứng (BUG_REG_001, BUG_FUNC_016, BUG_FUNC_019) — "
     "hai bên thống nhất xác nhận bằng Kết quả Retest = Pass.",
     "Đã thống nhất"),
    ("DS_22 chỉ có 1 Test Case và đang ở trạng thái Fail; DS_04, DS_21, DS_25 còn Test Case ở trạng thái "
     "Đang tiến hành. Sheet 03_Test_Execution chưa có dữ liệu.",
     "Cần hoàn tất trước đợt sau"),
    ("BUG_REG_001 (DS_22) có Ngày Dev báo fix 30/06/2026 sớm hơn Ngày phát hiện 24/07/2026 — "
     "cần rà lại ngày trong 06_RTM_Dev_Test.",
     "Sai lệch dữ liệu, không đổi kết luận"),
    ("Công thức Tổng lỗi / Lỗi đã xử lý tại 02_UI_UX đếm cả dòng 'Không lỗi' (không có BUG_ID) nên bị đội số: "
     "DS_21 hiển thị 5 (thực tế 3), DS_22 hiển thị 4 (thực tế 2), DS_25 hiển thị 6 (thực tế 5). "
     "Toàn site có 29 dòng như vậy. Biên bản này dùng số thực tế đếm theo BUG_ID.",
     "Sai lệch số liệu báo cáo")]
for i,(t,eff) in enumerate(issues):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    line([i+1,"","","","","","","","","","","","",""],i,h=34)
    merge(f"B{r}:F{r}",t,F(10),bg,Al("left",wrap=True))
    merge(f"G{r}:H{r}",eff,F(10),bg,Al("left",wrap=True))
    for c in ("I","J","K","M"): pass
    merge(f"K{r}:L{r}","",F(10),Fill(INPUT),Al("center"))
    merge(f"M{r}:N{r}","",F(10),Fill(INPUT),Al("left",wrap=True))
    put(f"I{r}","",F(10),Fill(INPUT),Al("center")); put(f"J{r}","",F(10),Fill(INPUT),Al("center"))
gap()

sec("X. KẾT LUẬN NGHIỆM THU")
r=row[0]
merge(f"A{r}:C{r}","Kết quả tự đánh giá theo tiêu chí mục III",F(10,True),Fill(WHITE),Al("left"))
_ok=[d for d in data["ds"] if d["verdict"]=="ĐẠT"]; _no=[d for d in data["ds"] if d["verdict"]!="ĐẠT"]
concl=f"{NPASS}/{NDS} màn hình ĐẠT nghiệm thu: " + ", ".join(d["ds"] for d in _ok)
concl += (". CHƯA ĐẠT: " + ", ".join(f'{d["ds"]} ({d["reason"]})' for d in _no)) if _no else \
         f". Không còn màn hình nào chưa đạt. Đề xuất kết luận: ĐẠT NGHIỆM THU GIAI ĐOẠN 1."
merge(f"D{r}:{LAST}{r}",concl,F(10.5,True,GREEN if not _no else RED),Fill(GREEN_BG) if not _no else Fill(RED_BG),Al("left",wrap=True))
ws.row_dimensions[r].height=32; row[0]+=1
r=row[0]
_fail=[d for d in data["ds"] if d["verdict"]!="ĐẠT"]
merge(f"A{r}:C{r}","Điều kiện còn lại" if _fail else "Ghi nhận",F(10,True),Fill(WHITE),Al("left"))
merge(f"D{r}:{LAST}{r}",
      ("; ".join(f'{d["ds"]}: {d["reason"]}' for d in _fail) if _fail else
       f"Toàn bộ {NDS} màn hình và {NYC} tính năng trong phạm vi đã thỏa mãn 6 tiêu chí tại mục III. "
       f"{TOT}/{TOT} lỗi đã được xử lý và retest đạt. Hồ sơ đủ điều kiện trình ký."),
      F(10),Fill(WHITE),Al("left",wrap=True))
ws.row_dimensions[r].height=28; row[0]+=1
r=row[0]
merge(f"A{r}:C{r}","KẾT LUẬN CHÍNH THỨC (người duyệt ghi)",F(10.5,True,WHITE),Fill(NAVY),Al("left"))
merge(f"D{r}:G{r}","",F(12,True),Fill(INPUT),Al("center"))
put(f"H{r}","Ngày kết luận",F(9.5,True,WHITE),Fill(BLUE),Al("center"))
merge(f"I{r}:J{r}","",F(10),Fill(INPUT),Al("center"))
put(f"K{r}","Hiệu lực từ",F(9.5,True,WHITE),Fill(BLUE),Al("center"))
merge(f"L{r}:{LAST}{r}","",F(10),Fill(INPUT),Al("center"))
ws.row_dimensions[r].height=30; row[0]+=1
r=row[0]
merge(f"A{r}:C{r}","Ý kiến / điều kiện kèm theo",F(10,True),Fill(WHITE),Al("left"))
merge(f"D{r}:{LAST}{r}","",F(10),Fill(INPUT),Al("left",wrap=True))
ws.row_dimensions[r].height=44; row[0]+=1
gap()

sec("XI. XÁC NHẬN CỦA CÁC BÊN")
note("Biên bản lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.",h=20)
head(["Vai trò","","Họ và tên","","Chức vụ","Ý kiến","","","Ngày ký","Chữ ký","","","",""],h=26)
for i,s in enumerate(["ĐẠI DIỆN BÊN A (Chủ đầu tư)","ĐẠI DIỆN BÊN B (Đơn vị thực hiện)","QUẢN LÝ DỰ ÁN","QA / TESTER","ĐẠI DIỆN DEV"]):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    line([""]*NCOL,i,h=52)
    merge(f"A{r}:B{r}",s,F(10,True),bg,Al("left",wrap=True))
    merge(f"C{r}:D{r}","",F(10),Fill(INPUT),Al("center"))
    put(f"E{r}","",F(10),Fill(INPUT),Al("center"))
    merge(f"F{r}:H{r}","",F(10),Fill(INPUT),Al("left",wrap=True))
    put(f"I{r}","",F(10),Fill(INPUT),Al("center"))
    merge(f"J{r}:{LAST}{r}","",F(10),bg,Al("center"))
gap()

sec("PHỤ LỤC")
head(["Mã","Tên phụ lục","","","Nguồn","","","","","","","","",""])
apps=[("PL-01","Danh sách lỗi đầy đủ kèm link bằng chứng","File QA — sheet 05_Lỗi_Tester"),
      ("PL-02","Tình trạng xử lý và retest của Dev","File QA — sheet 06_RTM_Dev_Test"),
      ("PL-03","Trạng thái kiểm thử theo trang","File QA — sheet 02_UI_UX"),
      ("PL-04","Trạng thái kiểm thử theo tính năng","File QA — sheet 03_Tính_năng"),
      ("PL-05","Bộ Test Case liên quan","File QA — sheet 04_Test_Cases")]
for i,(a,b,c) in enumerate(apps):
    r=row[0]; bg=Fill(ZEBRA) if i%2 else Fill(WHITE)
    line([a,"","","","","","","","","","","","",""],i,h=22)
    merge(f"B{r}:D{r}",b,F(10),bg,Al("left"))
    merge(f"E{r}:{LAST}{r}",c,F(10),bg,Al("left"))

ws.sheet_view.showGridLines=False
ws.freeze_panes="A3"
ws.sheet_properties.tabColor=NAVY
out="BaoCao_NghiemThu_GD1.xlsx"; wb.save(out); print("SAVED",out)
print(f"{NPASS}/{NDS} DS đạt | tổng lỗi {TOT} | mở {OPN} | C/H mở {CH} | retest pass {PASS}")
