# -*- coding: utf-8 -*-
"""Trích dữ liệu nghiệm thu GĐ2 (Nhà tuyển dụng) — lần 2.

    python3 extract_nghiemthu_gd2.py <bien_ban_lan1.xlsx> <nguon_QA> <ket_qua_kiem_lai.json> [ngay_chot]

- <bien_ban_lan1.xlsx>: "Biên bản nghiệm thu GD2.xlsx" lập 29/09/2026 (Drive > Report Timviec123 > Tháng 9).
- <nguon_QA>: file QA "Tester Detail- Timviec123" tải về dạng .xlsx, HOẶC thư mục chứa
  loi.json + rtm.json (giá trị sheet 05_Lỗi_Tester / 06_RTM_Dev_Test lấy qua Sheets API).
- <ket_qua_kiem_lai.json>: kết quả kiểm lại trên website (xem ket_qua_kiem_lai_1610.json).

Xuất data_nghiemthu_gd2.json để build_nghiemthu_gd2.py dựng biên bản lần 2.
"""
import sys, os, re, json, datetime
import openpyxl

OPEN_ST = ("Open", "Chưa gửi Dev", "Đang sửa")
DEV_DONE_ST = ("Fixed", "Verified")
NTD_DS = [f"DS_{i}" for i in range(46, 68)]

# Lỗi phát sinh sau 29/09 thuộc phạm vi NTD (không có trong biên bản lần 1).
# "chan": True = Critical/High trong phạm vi -> tính vào điều kiện duyệt.
PHAT_SINH = {
    "BUG_API_003": dict(luong="F7", nhom="Phát sinh - chặn duyệt", chan=True,
                        ly_do="Phân quyền: token NTD đọc được dữ liệu riêng tư của Ứng viên (luồng F7)."),
    "BUG_API_002": dict(luong="F4", nhom="Phát sinh - chặn duyệt", chan=True,
                        ly_do="Link tin của NTD có thể mở ra nội dung tin khác (trang DS_65 thuộc F3/F4)."),
    "Retest_BUG_047": dict(luong="F3, F4", nhom="Phát sinh - chặn duyệt", chan=True,
                           ly_do="Tin hết hạn/đã đóng của NTD hiển thị sai (DS_65, DS_50 thuộc F3)."),
    "BUG_FUNC_045": dict(luong="F3", nhom="Phát sinh - chặn duyệt", chan=True,
                         ly_do="Sửa tin (DS_52) không thêm được chứng chỉ; retest Fail, đang sửa."),
    "BUG_FUNC_050": dict(luong="F3", nhom="Phát sinh - chặn duyệt", chan=True,
                         ly_do="Sửa tin (DS_52) lỗi 500 khi lưu; Dev báo đã sửa, chưa retest."),
    "BUG_FUNC_018": dict(luong="F2", nhom="Phát sinh - chặn duyệt", chan=True,
                         ly_do="Đăng tin đầu tiên: chọn Kỹ năng yêu cầu (F2); Dev báo đã sửa, chưa retest."),
    "BUG_UI_077": dict(luong="F2", nhom="Phát sinh - không chặn", chan=False,
                       ly_do="Đăng tin - icon/gợi ý bước Yêu cầu; retest Fail 02/10."),
    "BUG_UI_078": dict(luong="F3", nhom="Phát sinh - không chặn", chan=False,
                       ly_do="Đăng tin - icon cảnh báo xoá; retest Fail 02/10."),
    "BUG_UI_080": dict(luong="F3", nhom="Phát sinh - không chặn", chan=False,
                       ly_do="Sửa tin - vị trí chữ, định dạng thời gian; Dev báo đã sửa."),
    "BUG_LO_016": dict(luong="F3", nhom="Phát sinh - không chặn", chan=False,
                       ly_do="Sửa tin - bố cục lệch Figma (High); Dev báo đã sửa."),
    "BUG_API_004": dict(luong="", nhom="Phát sinh - toàn hệ thống", chan=False,
                        ly_do="Thiếu API làm mới token; ảnh hưởng mọi vai trò, theo dõi riêng."),
}
NHOM_CHAN = ("Cần sửa trước khi duyệt", "QA - cần xử lý")


def s(x):
    if x is None:
        return ""
    if isinstance(x, (datetime.datetime, datetime.date)):
        return x.strftime("%d/%m/%Y")
    return str(x).replace("\xa0", " ").strip()


def sheet_rows(ws, hdr_row=1):
    hdr = [s(c.value) for c in ws[hdr_row]]
    out = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, len(hdr) + 1)]
        if not any(v not in (None, "") for v in vals):
            continue
        out.append({h: s(v) for h, v in zip(hdr, vals) if h})
    return out


def load_qa(src):
    """Trả về (loi: {BUG_ID: dòng 05}, rtm: {BUG_ID: dòng RTM}) — header ở dòng 4."""
    if os.path.isdir(src):
        def rows(fn):
            v = json.load(open(os.path.join(src, fn), encoding="utf-8"))["values"]
            h = [s(x) for x in v[3]]
            return [dict(zip(h, [s(x) for x in r] + [""] * (len(h) - len(r)))) for r in v[4:]]
        L, R = rows("loi.json"), rows("rtm.json")
    else:
        wb = openpyxl.load_workbook(src, data_only=True)
        L, R = sheet_rows(wb["05_Lỗi_Tester"], 4), sheet_rows(wb["06_RTM_Dev_Test"], 4)
    loi = {x["BUG_ID"]: x for x in L if x.get("BUG_ID")}
    rtm = {}
    for x in R:
        if x.get("BUG_ID"):
            rtm.setdefault(x["BUG_ID"], x)
    return loi, rtm


def danh_gia(item):
    """Đánh giá lần 2 của một mã lỗi theo thứ tự ưu tiên:
    kết quả kiểm lại 16/10 > Retest Pass trong RTM sau 29/09 > trạng thái QA."""
    kl = item.get("kiem_lai") or {}
    kq = kl.get("ket_qua", "")
    if kq == "Pass":
        return "ĐÃ ĐẠT"
    if kq == "Fail":
        return "KHÔNG ĐẠT"
    if item["retest_rtm"] == "Pass" and item.get("retest_rtm_moi"):
        return "ĐÃ ĐẠT"
    st = item["tt_qa"]
    if st in DEV_DONE_ST:
        return "CHỜ KIỂM LẠI"
    if st in ("Closed", "Deferred"):
        return st.upper()
    return "CÒN MỞ"


def is_after(d, ref=datetime.date(2026, 9, 29)):
    try:
        return datetime.datetime.strptime(d, "%d/%m/%Y").date() > ref
    except ValueError:
        return False


def main(lan1, qa_src, kl_path, asof):
    wb1 = openpyxl.load_workbook(lan1, data_only=True)
    loi, rtm = load_qa(qa_src)
    kl = json.load(open(kl_path, encoding="utf-8"))
    kl_map = {x["ma"]: x for x in kl["items"]}

    def qa_info(ma):
        l, r = loi.get(ma, {}), rtm.get(ma, {})
        return dict(tt_qa=l.get("Trạng thái lỗi hiện tại") or r.get("Trạng thái lỗi", ""),
                    retest_rtm=r.get("Kết quả Retest", ""), ngay_retest_rtm=r.get("Ngày Retest", ""),
                    retest_rtm_moi=is_after(r.get("Ngày Retest", "")),
                    ngay_fix=r.get("Ngày Dev báo fix", ""), dev=r.get("Dev phụ trách", ""),
                    xu_ly_dev=r.get("Phản hồi/Cách xử lý Dev", ""), ngay_gui=r.get("Ngày gửi lỗi", ""))

    # ---- lỗi lần 1 + phát sinh ----
    items = []
    for x in sheet_rows(wb1["Danh sách lỗi"]):
        it = dict(ma=x["Mã"], nhom=x["Nhóm"], chan=x["Nhóm"] in NHOM_CHAN, loai=x["Loại"], muc_do=x["Mức độ"],
                  ds=x["DS_ID"], trang=x["Trang"], yc=x["YC_ID"], tinh_nang=x["Tính năng"], tieu_de=x["Tiêu đề"],
                  mong_doi=x["Kết quả mong đợi"], thiet_bi=x["Thiết bị"], kich_ban=x["Kịch bản"],
                  lan1=x["Kiểm lại 29/09"] or "Ghi nhận 28–29/09", tt_lan1=x["Trạng thái"], phat_sinh=False)
        it.update(qa_info(it["ma"]))
        it["kiem_lai"] = kl_map.get(it["ma"])
        it["danh_gia"] = danh_gia(it)
        items.append(it)
    for ma, meta in PHAT_SINH.items():
        l = loi.get(ma)
        if not l:
            continue
        it = dict(ma=ma, nhom=meta["nhom"], chan=meta["chan"], loai=l.get("Loại lỗi", ""), muc_do=l.get("Mức độ", ""),
                  ds=l.get("DS_ID", ""), trang=l.get("Tên trang", ""), yc=l.get("YC_ID", ""),
                  tinh_nang=l.get("Tính năng", ""), tieu_de=l.get("Tiêu đề lỗi", ""),
                  mong_doi=l.get("Kết quả mong đợi", ""), thiet_bi=l.get("Thiết bị", ""), kich_ban="",
                  lan1="Chưa có ở lần 1 (phát sinh " + l.get("Ngày phát hiện", "") + ")", tt_lan1="",
                  phat_sinh=True, ly_do=meta["ly_do"], luong=meta["luong"])
        it.update(qa_info(ma))
        it["kiem_lai"] = kl_map.get(ma)
        it["danh_gia"] = danh_gia(it)
        items.append(it)
    by_ma = {it["ma"]: it for it in items}

    def tinh_trang_ma(codes):
        cs = [c.strip() for c in re.split(r"[,\n]", codes or "") if c.strip()]
        return {c: by_ma[c]["danh_gia"] for c in cs if c in by_ma}

    # mã lỗi QA "cần xử lý" theo trang (lần 1) -> dùng để xét luồng đi qua trang đó
    qa_theo_trang = {x["DS_ID"]: x["Mã lỗi QA cần xử lý"] for x in sheet_rows(wb1["Giao diện UI-UX theo trang"])}

    # ---- luồng ----
    luong = []
    for x in sheet_rows(wb1["Kết quả theo luồng"]):
        ds_list = [d.strip() for d in x["Trang liên quan (DS_ID)"].split(",") if d.strip()]
        codes = [x["Mã lỗi / yêu cầu của luồng"]] + [qa_theo_trang.get(d, "") for d in ds_list]
        codes += [it["ma"] for it in items if it["phat_sinh"] and x["Luồng"] in it.get("luong", "")]
        st = tinh_trang_ma(", ".join(codes))
        chan = [c for c, v in st.items() if by_ma[c]["chan"] and v != "ĐÃ ĐẠT"]
        luong.append(dict(ma=x["Luồng"], ten=x["Tên luồng"], so_kb=x["Số kịch bản"], dat=x["Đạt"], loi=x["Lỗi"],
                          ghi_nhan=x["Ghi nhận"], trang=x["Trang liên quan (DS_ID)"], ma_loi=x["Mã lỗi / yêu cầu của luồng"],
                          kl1=x["Kết luận"], ma_tt=st, chan_con=chan))

    # ---- tính năng (YC) ----
    yc = []
    for x in sheet_rows(wb1["Tính năng theo yêu cầu"]):
        st = tinh_trang_ma(", ".join([x["Mã lỗi / yêu cầu đợt này"], x["Mã lỗi QA liên quan"]]))
        yc.append(dict(ma=x["Mã yêu cầu"], ten=x["Tên yêu cầu"], trang=x["Trang"], uu_tien=x["Ưu tiên"],
                       tt_qa1=x["Trạng thái QA"], so_kb=x["Số kịch bản"], dat=x["Đạt"], loi=x["Lỗi"],
                       ghi_nhan=x["Ghi nhận"], kl1=x["Kết luận"], ma_tt=st))

    # ---- trang ----
    trang = []
    for x in sheet_rows(wb1["Giao diện UI-UX theo trang"]):
        st = tinh_trang_ma(", ".join([x["Mã lỗi QA cần xử lý"], x["Mã lỗi QA lệch nhỏ"], x["Mã lỗi / yêu cầu đợt này"]]))
        trang.append(dict(ds=x["DS_ID"], ten=x["Tên trang"], tong=x["Tổng lỗi QA"], da_xl=x["Đã xử lý"],
                          con_mo=x["Còn mở"], kl1=x["Kết luận"], ma_tt=st,
                          tt_qa_moi=sum(1 for l in loi.values() if l.get("DS_ID") == x["DS_ID"]
                                        and l.get("Trạng thái lỗi hiện tại") in OPEN_ST)))

    # ---- kịch bản ----
    kb_kl = {x["ma"]: x for x in kl.get("kich_ban", [])}
    kich_ban = []
    for x in sheet_rows(wb1["Kịch bản chi tiết"]):
        st = tinh_trang_ma(x["Mã lỗi / yêu cầu"])
        kich_ban.append(dict(ma=x["Mã"], luong=x["Luồng"], ten=x["Kịch bản"], uu_tien=x["Ưu tiên"],
                             thiet_bi=x["Thiết bị"], ds=x["Trang (DS_ID)"], yc=x["Yêu cầu (YC_ID)"],
                             kq1=x["Kết quả kiểm thử"], kl1=x["Kết luận"], ma_loi=x["Mã lỗi / yêu cầu"],
                             ma_tt=st, kiem_lai=kb_kl.get(x["Mã"])))

    # ---- lỗi Verified toàn site (Dev báo sửa, chờ tester kiểm) ----
    verified = []
    for ma, r in rtm.items():
        if r.get("Trạng thái lỗi") != "Verified":
            continue
        l = loi.get(ma, {})
        verified.append(dict(ma=ma, ds=r.get("DS_ID", ""), trang=r.get("Tên trang", ""), muc_do=r.get("Mức độ", ""),
                             thiet_bi=r.get("Thiết bị", ""), tieu_de=l.get("Tiêu đề lỗi", "") or r.get("Mô tả hiện tượng", ""),
                             hien_tuong=l.get("Mô tả hiện tượng", "") or r.get("Mô tả hiện tượng", ""),
                             retest_rtm=r.get("Kết quả Retest", ""), ngay_retest_rtm=r.get("Ngày Retest", ""),
                             ngay_fix=r.get("Ngày Dev báo fix", ""),
                             ntd=any(d in r.get("DS_ID", "") for d in NTD_DS) or ma in PHAT_SINH,
                             kiem_lai=kl_map.get(ma)))

    out = dict(asof=asof, kiem_lai_meta={k: v for k, v in kl.items() if k not in ("items", "kich_ban")},
               items=items, luong=luong, yc=yc, trang=trang, kich_ban=kich_ban, verified=verified)
    dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_nghiemthu_gd2.json")
    json.dump(out, open(dst, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    import collections
    print("Lỗi lần 1 + phát sinh:", len(items), collections.Counter(i["danh_gia"] for i in items))
    print("Chặn còn lại:", [i["ma"] for i in items if i["chan"] and i["danh_gia"] != "ĐÃ ĐẠT"])
    print("Verified:", len(verified), collections.Counter((v["kiem_lai"] or {}).get("ket_qua", "Chưa kiểm") for v in verified))


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3],
         sys.argv[4] if len(sys.argv) > 4 else datetime.date.today().strftime("%d/%m/%Y"))
