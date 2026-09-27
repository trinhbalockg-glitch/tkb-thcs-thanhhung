# -*- coding: utf-8 -*-
"""
HỆ THỐNG XẾP THỜI KHÓA BIỂU - THCS THẠNH HƯNG
Refactored: 2026-09-26

Thay đổi chính so với bản cũ:
  - Gộp tất cả các hàm tiện ích bị định nghĩa trùng lặp (bo_dau,
    chuan_hoa_ten_mon, lay_so_tiet_chuan_cho_mon, ...) thành một bản duy nhất.
  - Bỏ khối chương trình CLI dùng đường dẫn cứng (D:\\2026-2027\\...) vì nó
    không dùng được trên máy chủ Streamlit và là nguyên nhân gây lỗi biến
    ngoài phạm vi (NameError) + lỗi cú pháp (return ngoài hàm).
  - Viết lại logic CP-SAT (đọc dữ liệu -> tạo model -> giải -> xuất kết quả)
    thành MỘT hàm lõi dùng chung: _giai_tkb_core(). Hai hàm mỏng bên trên nó
    (giai_tkb_dataframe, giai_tkb_excel) chỉ khác nhau ở phần xuất kết quả,
    nên không còn hai bản logic xếp lịch khác nhau có thể lệch nhau.
  - Gộp hàm mục tiêu (objective) vốn bị viết hai lần thành một bản duy nhất.
  - Thay vì ghi file .xlsx ra ổ đĩa, nút tải xuống trong Streamlit tạo file
    trong bộ nhớ (BytesIO) - phù hợp khi chạy trên máy chủ / Streamlit Cloud.
"""

import io
import re
import unicodedata
import zipfile

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from ortools.sat.python import cp_model


class TkbError(Exception):
    """Lỗi nghiệp vụ khi xếp TKB (dữ liệu không hợp lệ, không tìm được nghiệm...)."""
    pass


# ============================================================
# 0. HÀM TIỆN ÍCH TIỀN XỬ LÝ DỮ LIỆU (mỗi hàm chỉ định nghĩa 1 lần)
# ============================================================

def bo_dau(text):
    """Bỏ dấu tiếng Việt và chuẩn hóa chuỗi."""
    if text is None:
        return ""
    s = str(text).strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.replace("Đ", "D").replace("đ", "d")


def chuan_hoa_ten_mon(ten_mon):
    """Chuẩn hoá tên môn học để so khớp (bỏ dấu, viết hoa, chỉ giữ A-Z0-9)."""
    if ten_mon is None:
        return ""
    s = bo_dau(ten_mon).upper()
    return re.sub(r"[^A-Z0-9]", "", s)


def chuan_hoa_ten_gv(gv):
    if gv is None:
        return ""
    return " ".join(str(gv).strip().split())


def lay_ten_ngan_gv(ten_gv):
    """Rút gọn tên giáo viên hiển thị trên TKB (lấy 2 từ cuối)."""
    if not ten_gv:
        return ""
    tu = str(ten_gv).strip().split()
    return " ".join(tu) if len(tu) <= 2 else " ".join(tu[-2:])


def chuan_hoa_ten_lop(raw):
    """Chuẩn hóa chuỗi tên lớp về dạng '6/1'."""
    if raw is None:
        return None
    s = str(raw).strip()
    for w in ("Lớp", "lớp", "Lop", "lop"):
        s = s.replace(w, "")
    s = s.strip()
    m = re.match(r"^(\d{1,2})\s*[/\-.]\s*(\d{1,2})$", s)
    return f"{int(m.group(1))}/{int(m.group(2))}" if m else None


def _parse_so_tiet(val):
    if val is None or val == "" or isinstance(val, bool):
        return 0
    if isinstance(val, (int, float)):
        return int(val) if val > 0 else 0
    s = str(val).strip()
    if not s or s in "-–—":
        return 0
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 0


_TU_KHOA_TINHOC = {"TINHOC", "TIN", "INFORMATICS", "CNTT"}
_TU_KHOA_THEDUC = {"THEDUC", "GDTC", "THEDUCTHETHAO", "PE", "THECHAT"}
_TU_KHOA_GDDP = {"GDDP", "GIAODUCDIAPHUONG", "DIAPHUONG"}


def la_mon_tin_hoc(mon):
    n = chuan_hoa_ten_mon(mon)
    return n in _TU_KHOA_TINHOC or "TINHOC" in n


def la_mon_the_duc(mon):
    n = chuan_hoa_ten_mon(mon)
    return n in _TU_KHOA_THEDUC or "THEDUC" in n or "GDTC" in n or "THECHAT" in n


def la_mon_gddp(mon):
    n = chuan_hoa_ten_mon(mon)
    return n in _TU_KHOA_GDDP or "GIAODUCDIAPHUONG" in n or "DIAPHUONG" in n


def la_mon_cong_nghe(mon):
    """Bắt mọi kiểu viết tên môn Công nghệ: 'C-nghệ', 'CN', 'Cong nghe'..."""
    if not mon:
        return False
    n = chuan_hoa_ten_mon(mon)
    if not n or la_mon_tin_hoc(mon):
        return False
    return n in ("CONGNGHE", "CNGHE", "CN") or n.startswith("CONGNGHE") or n.startswith("CNGHE")


def lay_so_tiet_chuan_cho_mon(ten_mon, khoi_lop, so_tiet_excel=0):
    """Số tiết chuẩn/tuần theo phân phối chương trình cho từng môn."""
    norm = chuan_hoa_ten_mon(ten_mon)
    if "NGUVAN" in norm or "VAN" in norm or "TOAN" in norm or "KHTN" in norm or "KHOAHOCTUNHIEN" in norm:
        return 4
    if "TIENGANH" in norm or "ANH" in norm or "LICHSU" in norm or "DIALI" in norm or "LS-ĐL" in norm or "SU" in norm or "DIA" in norm:
        return 3
    if "SINHHOC" in norm or "SINH" in norm:
        return 2
    if "VATLY" in norm or "VLI" in norm or "LY" in norm or "HOAHOC" in norm or "HOA" in norm:
        return 1
    if la_mon_cong_nghe(ten_mon):
        return 2 if str(khoi_lop) == "9" else 1
    return so_tiet_excel if 0 < so_tiet_excel <= 4 else 1


# ============================================================
# 1. ĐỌC FILE PHÂN CÔNG TỪ UPLOAD (STREAMLIT)
# ============================================================

def _go_bo_hinh_anh_trong_xlsx(du_lieu_file):
    """
    Trả về nội dung .xlsx sau khi loại bỏ mọi hình ảnh/logo nhúng trong file.

    File Excel phân công thường có logo trường được dán ở đầu sheet. Nếu logo
    đó ở định dạng cũ (WMF/EMF) hoặc bị lỗi, openpyxl sẽ báo lỗi
    "Invalid binary data format: <class 'NoneType'>" ngay khi mở file, trước
    cả khi đọc đến dữ liệu phân công. Vì ứng dụng chỉ cần giá trị các ô, hàm
    này gỡ bỏ phần hình ảnh (media, drawings) khỏi file .xlsx (vốn chỉ là một
    file .zip) trước khi đưa cho openpyxl đọc, để lỗi đó không còn xảy ra.
    """
    zip_goc = zipfile.ZipFile(io.BytesIO(du_lieu_file), "r")
    buffer_moi = io.BytesIO()
    with zipfile.ZipFile(buffer_moi, "w", zipfile.ZIP_DEFLATED) as zip_moi:
        for item in zip_goc.infolist():
            noi_dung = zip_goc.read(item.filename)

            # Bỏ hẳn các phần hình ảnh và bản vẽ (drawings)
            if item.filename.startswith("xl/media/") or item.filename.startswith("xl/drawings/"):
                continue

            # Gỡ thẻ <drawing .../> khỏi từng sheet để không còn tham chiếu treo
            if re.match(r"^xl/worksheets/sheet\d+\.xml$", item.filename):
                noi_dung = re.sub(rb"<drawing[^>]*/>", b"", noi_dung)

            # Gỡ các Relationship trỏ tới drawing trong file .rels của từng sheet
            if re.match(r"^xl/worksheets/_rels/sheet\d+\.xml\.rels$", item.filename):
                noi_dung = re.sub(
                    rb'<Relationship[^>]*Type="[^"]*?/drawing"[^>]*/>', b"", noi_dung
                )

            # Gỡ khai báo Override cho các phần drawing trong Content_Types
            if item.filename == "[Content_Types].xml":
                noi_dung = re.sub(
                    rb'<Override PartName="/xl/drawings/[^"]*"[^>]*/>', b"", noi_dung
                )

            zip_moi.writestr(item, noi_dung)

    buffer_moi.seek(0)
    return buffer_moi.getvalue()


def doc_phan_cong_stream(uploaded_file):
    """Đọc file phân công giảng dạy (.xlsx) upload từ Streamlit."""
    du_lieu_file = uploaded_file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(du_lieu_file), data_only=True)
    except Exception:
        # File có hình ảnh/logo mà openpyxl không đọc được -> thử lại sau khi
        # gỡ bỏ hình ảnh, vì ứng dụng không cần đến chúng.
        try:
            du_lieu_da_xu_ly = _go_bo_hinh_anh_trong_xlsx(du_lieu_file)
            wb = openpyxl.load_workbook(io.BytesIO(du_lieu_da_xu_ly), data_only=True)
        except Exception as loi_lan_2:
            raise TkbError(
                "Không thể đọc file Excel này (có thể do hình ảnh/logo trong file bị lỗi "
                "định dạng). Hãy thử mở file trong Excel, xóa hình ảnh, lưu lại rồi tải lên "
                f"lại. (Chi tiết lỗi: {loi_lan_2})"
            )
    phan_cong_list, gv_chu_nhiem = [], {}

    for ten_sheet in wb.sheetnames:
        ws = wb[ten_sheet]
        if ws.max_row < 2:
            continue

        dong_tieu_de = 1
        for r in range(1, min(ws.max_row, 10) + 1):
            val_b = str(ws.cell(row=r, column=2).value or "")
            if "GIÁO VIÊN" in val_b.upper() or "GIAO VIEN" in bo_dau(val_b).upper():
                dong_tieu_de = r
                break

        for r in range(dong_tieu_de + 1, ws.max_row + 1):
            gv_val = ws.cell(row=r, column=2).value    # Cột B: Giáo viên
            mon_val = ws.cell(row=r, column=3).value   # Cột C: Môn
            lop_val = ws.cell(row=r, column=5).value   # Cột E: Lớp
            st_val = ws.cell(row=r, column=7).value    # Cột G: Số tiết
            cn_val = ws.cell(row=r, column=8).value    # Cột H: Chủ nhiệm lớp

            if not gv_val or not str(gv_val).strip():
                continue
            gv = chuan_hoa_ten_gv(gv_val)

            if cn_val and str(cn_val).strip():
                ds_cn = re.findall(r'(\d{1,2}\s*[/\-.]\s*\d{1,2})', str(cn_val))
                for cn_item in ds_cn:
                    l_chuan = chuan_hoa_ten_lop(cn_item)
                    if l_chuan:
                        gv_chu_nhiem[l_chuan] = gv
                if not ds_cn and "/" in str(cn_val):
                    l_chuan = chuan_hoa_ten_lop(cn_val)
                    if l_chuan:
                        gv_chu_nhiem[l_chuan] = gv

            if not mon_val or not lop_val:
                continue
            mon, so_tiet_excel = str(mon_val).strip(), _parse_so_tiet(st_val)

            raw_lop, ds_lop = str(lop_val).strip(), []
            match_khoi = re.match(r'^(\d{1,2})\s*/\s*([\d,\s\.\-]+)$', raw_lop)
            if match_khoi:
                khoi = match_khoi.group(1)
                for num in re.findall(r'\d+', match_khoi.group(2)):
                    ds_lop.append(f"{khoi}/{num}")
            else:
                for cum in re.findall(r'\d{1,2}\s*[/\-.]\s*\d{1,2}', raw_lop):
                    l_chuan = chuan_hoa_ten_lop(cum)
                    if l_chuan:
                        ds_lop.append(l_chuan)

            for lop_item in ds_lop:
                khoi_lop = lop_item.split("/")[0]
                so_tiet = lay_so_tiet_chuan_cho_mon(mon, khoi_lop, so_tiet_excel)
                for _ in range(so_tiet):
                    phan_cong_list.append({"class": lop_item, "subject": mon, "teacher": gv})

    phan_cong_list = [
        item for item in phan_cong_list
        if not la_mon_tin_hoc(item["subject"]) and not la_mon_the_duc(item["subject"])
        and not (item["class"].startswith("9/") and la_mon_gddp(item["subject"]))
    ]
    if not phan_cong_list:
        raise TkbError("Không đọc được dữ liệu phân công giảng dạy nào từ file đã tải lên.")

    return phan_cong_list, gv_chu_nhiem


# ============================================================
# 2. LÕI THUẬT TOÁN CP-SAT (DÙNG CHUNG CHO MỌI HÌNH THỨC XUẤT KẾT QUẢ)
# ============================================================

CAC_THU = ["HAI", "BA", "TƯ", "NĂM", "SÁU"]
CAC_TIET = [1, 2, 3, 4, 5]
DANH_SACH_LOP_DIEM_LE_MAC_DINH = ["7/5", "8/5", "6/4"]


def _giai_tkb_core(
    pc_buoi,
    cn_buoi,
    khoi_list,
    ten_buoi,
    tiet_chao_co=1,
    lops_diem_le=None,
    bo_chao_co_lops=None,
    thoi_gian=300,
    log=None,
):
    """
    Xây dựng model CP-SAT và giải bài toán xếp TKB cho một buổi (sáng/chiều).

    Trả về dict:
        {solver, x, fixed, gv_mon_lop, danh_sach_lop, cac_thu, cac_tiet}
    Raise TkbError nếu dữ liệu không hợp lệ hoặc không tìm được nghiệm.

    `log`, nếu được truyền vào, là một list — mọi thông báo tiến trình sẽ
    được append vào đó thay vì chỉ in ra console (để Streamlit có thể hiển
    thị lại cho người dùng bằng st.write/st.expander).
    """
    def ghi_log(msg):
        if log is not None:
            log.append(msg)

    if lops_diem_le is None:
        lops_diem_le = list(DANH_SACH_LOP_DIEM_LE_MAC_DINH)
    if bo_chao_co_lops is None:
        bo_chao_co_lops = []

    cac_thu, cac_tiet = CAC_THU, CAC_TIET
    khoi_set = set(map(str, khoi_list))
    pc_buoi = [item for item in pc_buoi if item["class"].split("/")[0] in khoi_set]

    danh_sach_lop = sorted(
        set(item["class"] for item in pc_buoi),
        key=lambda l: (int(l.split("/")[0]), int(l.split("/")[1]))
    )

    mon_lop, gv_mon_lop = {}, {}
    for item in pc_buoi:
        key = (item["class"], item["subject"])
        mon_lop[key] = mon_lop.get(key, 0) + 1
        gv_mon_lop[key] = item["teacher"]

    # ---- 1. KIỂM TRA TẢI DỮ LIỆU ----
    ghi_log(f"KIỂM TRA TẢI DỮ LIỆU - {ten_buoi.upper()}")
    loi_du_lieu = []
    for lop in danh_sach_lop:
        so_tiet_mon = sum(n for (l, mon), n in mon_lop.items() if l == lop)
        co_chao_co = lop not in bo_chao_co_lops
        tong = so_tiet_mon + 1 + int(co_chao_co)
        muc_tieu = 25 if co_chao_co else 24
        ghi_log(f"  {lop:<6} | Môn={so_tiet_mon:2d} | CC={int(co_chao_co)} | SH=1 | TỔNG={tong}/{muc_tieu}")
        if tong != muc_tieu:
            loi_du_lieu.append(f"{lop}: có {tong} tiết nhưng cần {muc_tieu} tiết.")

    for (lop, mon), n in mon_lop.items():
        if "CONGNGHE" not in chuan_hoa_ten_mon(mon):
            continue
        grade = lop.split("/")[0]
        if grade == "8" and n != 1:
            loi_du_lieu.append(f"{lop}: Công nghệ phải đúng 1 tiết, hiện có {n}.")
        if grade == "9" and n != 2:
            loi_du_lieu.append(f"{lop}: Công nghệ phải đúng 2 tiết, hiện có {n}.")

    if loi_du_lieu:
        raise TkbError(
            "Dữ liệu không thể xếp sau khi kiểm tra tải:\n" + "\n".join(f"- {l}" for l in loi_du_lieu)
        )
    ghi_log("✅ TẢI DỮ LIỆU HỢP LỆ.")

    # ---- 2. KHỞI TẠO MODEL & BIẾN ----
    model = cp_model.CpModel()
    x = {}
    for (lop, mon), so_tiet in mon_lop.items():
        for thu in range(5):
            for tiet in cac_tiet:
                x[(lop, mon, thu, tiet)] = model.NewBoolVar(f"x_{lop}_{mon}_{thu}_{tiet}")

    # ---- 3. RÀNG BUỘC ĐỦ TIẾT ----
    for (lop, mon), so_tiet in mon_lop.items():
        model.Add(sum(x[(lop, mon, thu, tiet)] for thu in range(5) for tiet in cac_tiet) == so_tiet)

    # ---- 4. CỐ ĐỊNH CHÀO CỜ VÀ SINH HOẠT LỚP ----
    fixed = []
    for lop in danh_sach_lop:
        gv_cn = cn_buoi.get(lop)
        if not gv_cn:
            ghi_log(f"⚠️ CẢNH BÁO: Không tìm thấy GVCN lớp {lop}.")
            continue
        if lop not in bo_chao_co_lops:
            fixed.append((lop, "Chào cờ", gv_cn, 0, tiet_chao_co))
        fixed.append((lop, "SH Lớp", gv_cn, 4, 5))

    fixed_gv = {}
    for lop, mon, gv, thu, tiet in fixed:
        key = (gv, thu, tiet)
        if key in fixed_gv:
            raise TkbError(f"GVCN bị trùng tiết cố định: GV {gv} ở lớp {fixed_gv[key]} và {lop}")
        fixed_gv[key] = lop

    # ---- 5. MỖI LỚP TỐI ĐA 1 MÔN / Ô ----
    for lop in danh_sach_lop:
        for thu in range(5):
            for tiet in cac_tiet:
                vars_slot = [x[(l, mon, thu, tiet)] for (l, mon) in mon_lop if l == lop]
                co_fixed = any(fl == lop and fthu == thu and ftiet == tiet for (fl, fmon, fgv, fthu, ftiet) in fixed)
                if co_fixed:
                    if vars_slot:
                        model.Add(sum(vars_slot) == 0)
                elif vars_slot:
                    model.Add(sum(vars_slot) <= 1)

    # ---- 6. GIÁO VIÊN KHÔNG TRÙNG TIẾT ----
    danh_sach_gv = sorted(set(gv_mon_lop.values()) | set(cn_buoi.values()))
    for gv in danh_sach_gv:
        for thu in range(5):
            for tiet in cac_tiet:
                vars_gv = [x[(lop, mon, thu, tiet)] for (lop, mon), g in gv_mon_lop.items() if g == gv]
                fixed_count = sum(1 for (fl, fmon, fgv, fthu, ftiet) in fixed if fgv == gv and fthu == thu and ftiet == tiet)
                if fixed_count:
                    if vars_gv:
                        model.Add(sum(vars_gv) == 0)
                elif vars_gv:
                    model.Add(sum(vars_gv) <= 1)

    # ---- 7. KHÔNG CÓ TIẾT LỎNG ----
    for lop in danh_sach_lop:
        for thu in range(5):
            muc_tieu_ngay = 4 if (thu == 0 and lop in bo_chao_co_lops) else 5
            vars_day = [x[(l, mon, thu, tiet)] for (l, mon) in mon_lop if l == lop for tiet in cac_tiet]
            fixed_count = sum(1 for (fl, fmon, fgv, fthu, ftiet) in fixed if fl == lop and fthu == thu)
            model.Add(sum(vars_day) + fixed_count == muc_tieu_ngay)

            for tiet in cac_tiet:
                vars_slot = [x[(lop, mon, thu, tiet)] for (l, mon) in mon_lop if l == lop]
                co_fixed = any(fl == lop and fthu == thu and ftiet == tiet for (fl, fmon, fgv, fthu, ftiet) in fixed)
                if co_fixed:
                    continue
                if thu == 0 and lop in bo_chao_co_lops and tiet == 5:
                    model.Add(sum(vars_slot) == 0)
                else:
                    model.Add(sum(vars_slot) == 1)

    # ---- 8. PHÂN BỐ MÔN TRONG NGÀY (CÔNG NGHỆ 9 ÉP TÁCH NGÀY) ----
    for (lop, mon), so_tiet in mon_lop.items():
        is_cn, is_khoi9 = la_mon_cong_nghe(mon), lop.startswith("9/")
        for thu in range(5):
            vars_day = [x[(lop, mon, thu, tiet)] for tiet in cac_tiet]
            if is_cn and is_khoi9:
                model.Add(sum(vars_day) <= 1)
            elif so_tiet >= 3:
                model.Add(sum(vars_day) <= 2)
            elif so_tiet == 1:
                model.Add(sum(vars_day) <= 1)

    # ---- 9. CÔNG NGHỆ KHỐI 8 & 9 (SỐ TIẾT CỐ ĐỊNH, KHÔNG LIỀN NHAU Ở KHỐI 9) ----
    for (lop, mon), so_tiet in mon_lop.items():
        if not la_mon_cong_nghe(mon):
            continue
        grade = lop.split("/")[0]
        if grade == "8":
            model.Add(sum(x[(lop, mon, thu, tiet)] for thu in range(5) for tiet in cac_tiet) == 1)
        elif grade == "9":
            model.Add(sum(x[(lop, mon, thu, tiet)] for thu in range(5) for tiet in cac_tiet) == 2)
            for thu in range(5):
                for tiet in range(1, 5):
                    model.Add(x[(lop, mon, thu, tiet)] + x[(lop, mon, thu, tiet + 1)] <= 1)

    # ---- 10. RÀNG BUỘC CỨNG: cấm GV có tên chứa "LOC" dạy Công nghệ tiết 1 ----
    # (Áp dụng cho mọi buổi — nếu chỉ muốn áp dụng riêng buổi sáng, bọc thêm
    #  `if ten_buoi == "Buổi sáng":` quanh khối này.)
    for (lop, mon), gv in gv_mon_lop.items():
        if "LOC" in bo_dau(gv).upper() and la_mon_cong_nghe(mon):
            for thu in range(5):
                if (lop, mon, thu, 1) in x:
                    model.Add(x[(lop, mon, thu, 1)] == 0)

    # ---- 11. TIẾT ĐÔI (khuyến khích, không áp dụng cho Công nghệ khối 9) ----
    TRONG_SO_DOI_MON_NHIEU_TIET, TRONG_SO_DOI_MON_KHAC = 5, 1
    pair_vars_co_trong_so = []
    for (lop, mon), so_tiet in mon_lop.items():
        if la_mon_cong_nghe(mon) and lop.startswith("9/"):
            continue
        pairs_subject = []
        for thu in range(5):
            for tiet in range(1, 5):
                a, b = x[(lop, mon, thu, tiet)], x[(lop, mon, thu, tiet + 1)]
                pair = model.NewBoolVar(f"pair_{lop}_{mon}_{thu}_{tiet}")
                model.Add(pair <= a)
                model.Add(pair <= b)
                model.Add(pair >= a + b - 1)
                pairs_subject.append(pair)
        if so_tiet == 4:
            model.Add(sum(pairs_subject) >= 1)
        trong_so = TRONG_SO_DOI_MON_NHIEU_TIET if so_tiet >= 3 else TRONG_SO_DOI_MON_KHAC
        for pair in pairs_subject:
            pair_vars_co_trong_so.append((pair, trong_so))

    # ---- 12. GIÁO VIÊN DI CHUYỂN ĐIỂM LẺ: không dạy 2 tiết liền ở 2 điểm khác nhau ----
    # LƯU Ý: phải tính cả các tiết CỐ ĐỊNH (Chào cờ, SH Lớp) chứ không chỉ các
    # tiết dạy môn — vì đó chính xác là nơi GVCN bị "buộc" vào cơ sở của lớp
    # chủ nhiệm mình, và trước đây bị bỏ sót nên không phát hiện được xung đột.
    set_gvcn = set(cn_buoi.values())

    for gv in sorted(set(gv_mon_lop.values()) | set_gvcn):
        mon_cua_gv = [(l, m) for (l, m), g in gv_mon_lop.items() if g == gv]
        fixed_cua_gv = [(fl, fthu, ftiet) for (fl, fmon, fgv, fthu, ftiet) in fixed if fgv == gv]

        for thu in range(5):
            for tiet in range(1, 5):
                le_k = sum(x[(l, m, thu, tiet)] for (l, m) in mon_cua_gv if l in lops_diem_le and (l, m, thu, tiet) in x)
                le_k += sum(1 for (fl, fthu, ftiet) in fixed_cua_gv if fthu == thu and ftiet == tiet and fl in lops_diem_le)

                chinh_k = sum(x[(l, m, thu, tiet)] for (l, m) in mon_cua_gv if l not in lops_diem_le and (l, m, thu, tiet) in x)
                chinh_k += sum(1 for (fl, fthu, ftiet) in fixed_cua_gv if fthu == thu and ftiet == tiet and fl not in lops_diem_le)

                le_k1 = sum(x[(l, m, thu, tiet + 1)] for (l, m) in mon_cua_gv if l in lops_diem_le and (l, m, thu, tiet + 1) in x)
                le_k1 += sum(1 for (fl, fthu, ftiet) in fixed_cua_gv if fthu == thu and ftiet == tiet + 1 and fl in lops_diem_le)

                chinh_k1 = sum(x[(l, m, thu, tiet + 1)] for (l, m) in mon_cua_gv if l not in lops_diem_le and (l, m, thu, tiet + 1) in x)
                chinh_k1 += sum(1 for (fl, fthu, ftiet) in fixed_cua_gv if fthu == thu and ftiet == tiet + 1 and fl not in lops_diem_le)

                model.Add(le_k + chinh_k1 <= 1)
                model.Add(chinh_k + le_k1 <= 1)

    # ---- 13. RÀNG BUỘC & THƯỞNG CHO GVCN / GVBM ----
    gv_cn_day_vars, gv_bm_day_vars = [], []
    list_ngay_1_tiet, pair_vars_lien_tiep_gv = [], []
    list_khoang_trong_gv = []  # phạt "tiết lỏng": có tiết trước và sau nhưng trống ở giữa

    for gv in sorted(set(gv_mon_lop.values()) | set_gvcn):
        is_gvcn = gv in set_gvcn
        mon_cua_gv = [(l, mon) for (l, mon), g in gv_mon_lop.items() if g == gv]
        ten_gv_kd = bo_dau(gv)

        if is_gvcn:
            for thu_ep in (0, 4):  # Thứ Hai và Thứ Sáu
                vars_ep = [x[(l, mon, thu_ep, tiet)] for (l, mon) in mon_cua_gv for tiet in cac_tiet if (l, mon, thu_ep, tiet) in x]
                fixed_ep = sum(1 for (fl, fmon, fgv, fthu, ftiet) in fixed if fgv == gv and fthu == thu_ep)
                if vars_ep or fixed_ep > 0:
                    model.Add(sum(vars_ep) + fixed_ep >= 1)

        for thu in range(5):
            vars_day_gv = [x[(l, mon, thu, tiet)] for (l, mon) in mon_cua_gv for tiet in cac_tiet if (l, mon, thu, tiet) in x]
            fixed_day_count = sum(1 for (fl, fmon, fgv, fthu, ftiet) in fixed if fgv == gv and fthu == thu)

            co_day_ngay = model.NewBoolVar(f"gv_day_{ten_gv_kd}_{thu}")
            tong_tiet_ngay = sum(vars_day_gv) + fixed_day_count
            model.Add(tong_tiet_ngay >= 1).OnlyEnforceIf(co_day_ngay)
            model.Add(tong_tiet_ngay == 0).OnlyEnforceIf(co_day_ngay.Not())

            if is_gvcn:
                if thu in (0, 4):
                    gv_cn_day_vars.append(co_day_ngay)
            else:
                gv_bm_day_vars.append(co_day_ngay)

            is_1_tiet = model.NewBoolVar(f"gv_1tiet_{ten_gv_kd}_{thu}")
            model.Add(tong_tiet_ngay == 1).OnlyEnforceIf(is_1_tiet)
            model.Add(tong_tiet_ngay != 1).OnlyEnforceIf(is_1_tiet.Not())
            list_ngay_1_tiet.append(is_1_tiet)

            # --- "Có tiết ở tiết k trong ngày" — tính một lần, dùng lại cho cả
            #     thưởng tiết đôi lẫn phạt tiết lỏng bên dưới ---
            co_tiet_gv = {}
            for tiet in cac_tiet:
                vars_k = [x[(l, mon, thu, tiet)] for (l, mon) in mon_cua_gv if (l, mon, thu, tiet) in x]
                fixed_k = 1 if any(fgv == gv and fthu == thu and ftiet == tiet for (fl, fmon, fgv, fthu, ftiet) in fixed) else 0
                co_tiet_var = model.NewBoolVar(f"gv_co_tiet_{ten_gv_kd}_{thu}_{tiet}")
                model.Add(co_tiet_var == sum(vars_k) + fixed_k)
                co_tiet_gv[tiet] = co_tiet_var

            # --- Thưởng 2 tiết liền kề (giữ nguyên như trước) ---
            for tiet in range(1, 5):
                pair_gv = model.NewBoolVar(f"pair_gv_{ten_gv_kd}_{thu}_{tiet}")
                model.Add(pair_gv <= co_tiet_gv[tiet])
                model.Add(pair_gv <= co_tiet_gv[tiet + 1])
                model.Add(pair_gv >= co_tiet_gv[tiet] + co_tiet_gv[tiet + 1] - 1)
                pair_vars_lien_tiep_gv.append(pair_gv)

            # --- Phạt "tiết lỏng": có tiết TRƯỚC và tiết SAU nhưng tiết này trống
            #     (đây là trường hợp dạy tiết 1 rồi lại dạy tiết 4/5, ví dụ bạn nêu) ---
            for tiet in (2, 3, 4):
                co_truoc = model.NewBoolVar(f"gv_truoc_{ten_gv_kd}_{thu}_{tiet}")
                model.AddMaxEquality(co_truoc, [co_tiet_gv[t] for t in cac_tiet if t < tiet])
                co_sau = model.NewBoolVar(f"gv_sau_{ten_gv_kd}_{thu}_{tiet}")
                model.AddMaxEquality(co_sau, [co_tiet_gv[t] for t in cac_tiet if t > tiet])
                trong_tiet = model.NewBoolVar(f"gv_trong_{ten_gv_kd}_{thu}_{tiet}")
                model.Add(trong_tiet == 1 - co_tiet_gv[tiet])
                khoang_trong = model.NewBoolVar(f"gv_khoangtrong_{ten_gv_kd}_{thu}_{tiet}")
                model.AddMinEquality(khoang_trong, [co_truoc, co_sau, trong_tiet])
                list_khoang_trong_gv.append(khoang_trong)

    # ---- 14. HÀM MỤC TIÊU (một bản duy nhất, gộp mọi trọng số) ----
    TRONG_SO_PHAT_KHOANG_TRONG = 20  # mức "vừa phải" theo yêu cầu — xem ghi chú bên dưới nếu cần chỉnh

    diem_tiet_doi_mon = sum(trong_so * pair for pair, trong_so in pair_vars_co_trong_so)
    diem_gvcn = sum(10 * v for v in gv_cn_day_vars)
    diem_gvbm = sum(3 * v for v in gv_bm_day_vars)
    diem_tiet_lien_nhau_gv = sum(4 * v for v in pair_vars_lien_tiep_gv)
    phat_ngay_1_tiet = sum(15 * v for v in list_ngay_1_tiet)
    phat_khoang_trong_gv = sum(TRONG_SO_PHAT_KHOANG_TRONG * v for v in list_khoang_trong_gv)

    model.Maximize(
        diem_tiet_doi_mon
        + diem_gvcn
        + diem_gvbm
        + diem_tiet_lien_nhau_gv
        - phat_ngay_1_tiet
        - phat_khoang_trong_gv
    )

    # ---- 15. GIẢI MODEL ----
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = thoi_gian
    solver.parameters.num_search_workers = 8
    solver.parameters.cp_model_presolve = True
    solver.parameters.linearization_level = 2

    ghi_log(f"ĐANG TÌM NGHIỆM TKB {ten_buoi.upper()}...")
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise TkbError("Không tìm được nghiệm. Cần kiểm tra lại phân công hoặc ràng buộc.")

    return {
        "solver": solver, "x": x, "fixed": fixed, "gv_mon_lop": gv_mon_lop,
        "danh_sach_lop": danh_sach_lop, "cac_thu": cac_thu, "cac_tiet": cac_tiet,
    }


def _gia_tri_o(ket_qua, lop, thu, tiet):
    """Nội dung một ô trong bảng TKB: 'MÔN-GV' hoặc rỗng."""
    for (fl, fmon, fgv, fthu, ftiet) in ket_qua["fixed"]:
        if fl == lop and fthu == thu and ftiet == tiet:
            return f"{fmon}-{lay_ten_ngan_gv(fgv)}"
    solver, x = ket_qua["solver"], ket_qua["x"]
    for (l, mon), gv in ket_qua["gv_mon_lop"].items():
        if l == lop and solver.Value(x[(l, mon, thu, tiet)]) == 1:
            return f"{mon}-{lay_ten_ngan_gv(gv)}"
    return ""


# ============================================================
# 3. HAI CÁCH XUẤT KẾT QUẢ, DÙNG CHUNG LÕI Ở TRÊN
# ============================================================

def giai_tkb(pc_buoi, cn_buoi, khoi_list, ten_buoi, tiet_chao_co=1,
             lops_diem_le=None, bo_chao_co_lops=None, thoi_gian=300, log=None):
    """Giải TKB và trả về (DataFrame, danh_sach_lop) — dùng để hiển thị trong Streamlit."""
    ket_qua = _giai_tkb_core(
        pc_buoi, cn_buoi, khoi_list, ten_buoi, tiet_chao_co,
        lops_diem_le, bo_chao_co_lops, thoi_gian, log=log,
    )
    danh_sach_lop, cac_thu, cac_tiet = ket_qua["danh_sach_lop"], ket_qua["cac_thu"], ket_qua["cac_tiet"]

    rows_data = []
    for thu_idx, ten_thu in enumerate(cac_thu):
        for tiet in cac_tiet:
            row_dict = {"Thứ": ten_thu, "Tiết": tiet}
            for lop in danh_sach_lop:
                gia_tri = _gia_tri_o(ket_qua, lop, thu_idx, tiet)
                row_dict[f"Lớp {lop}"] = f"{gia_tri}" if gia_tri else ""
            rows_data.append(row_dict)

    return pd.DataFrame(rows_data), danh_sach_lop


def giai_tkb_va_xuat_excel_bytes(pc_buoi, cn_buoi, khoi_list, ten_buoi, tieu_de_ky_hoc,
                                  tiet_chao_co=1, lops_diem_le=None, bo_chao_co_lops=None,
                                  thoi_gian=300, log=None):
    """Giải TKB và trả về nội dung file .xlsx dưới dạng bytes (để tải xuống từ Streamlit)."""
    ket_qua = _giai_tkb_core(
        pc_buoi, cn_buoi, khoi_list, ten_buoi, tiet_chao_co,
        lops_diem_le, bo_chao_co_lops, thoi_gian, log=log,
    )
    danh_sach_lop, cac_thu, cac_tiet = ket_qua["danh_sach_lop"], ket_qua["cac_thu"], ket_qua["cac_tiet"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.cell(row=1, column=1, value=tieu_de_ky_hoc)
    ws.cell(row=2, column=1, value=f"KHỐI {', '.join(map(str, khoi_list))} {ten_buoi.upper()}")

    headers = ["Thứ", "Tiết"] + [f"Lớp {lop}" for lop in danh_sach_lop]
    for c, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=header)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row = 4
    for thu_idx, ten_thu in enumerate(cac_thu):
        start_row = row
        for tiet in cac_tiet:
            ws.cell(row=row, column=2, value=tiet).alignment = Alignment(horizontal="center", vertical="center")
            for c, lop in enumerate(danh_sach_lop, 3):
                cell = ws.cell(row=row, column=c, value=_gia_tri_o(ket_qua, lop, thu_idx, tiet))
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            row += 1
        ws.merge_cells(start_row=start_row, end_row=row - 1, start_column=1, end_column=1)
        ws.cell(row=start_row, column=1, value=ten_thu).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 7
    for c in range(3, 3 + len(danh_sach_lop)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 22

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# 4. GIAO DIỆN STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Hệ thống Xếp Thời Khóa Biểu - THCS Thạnh Hưng",
    page_icon="🏫",
    layout="wide"
)

import os  # dùng cho kiểm tra logo, để gần nơi dùng cho dễ theo dõi

if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)
elif os.path.exists("logo.jpg"):
    st.sidebar.image("logo.jpg", use_container_width=True)

col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=120)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=120)
    else:
        st.write("🏫")

with col_title:
    st.title("HỆ THỐNG XẾP THỜI KHÓA BIỂU")
    st.subheader("TRƯỜNG THCS THẠNH HƯNG")

st.markdown("---")

uploaded_file = st.sidebar.file_uploader("📂 Tải lên File Excel Phân Công (PCGDHKI.xlsx)", type=["xlsx"])
thoi_gian_giai = st.sidebar.slider("⏱️ Thời gian giải tối đa mỗi buổi (giây)", 30, 600, 300, step=30)

# Cấu hình riêng cho từng buổi — dùng chung cho cả nút "xếp cả hai buổi" lẫn
# hai nút "chỉ xếp lại buổi sáng / chiều", để không phải lặp lại các tham số
# này ở nhiều nơi (đây chính là kiểu trùng lặp đã gây lỗi trước đây).
CAU_HINH_BUOI = {
    "sang": {
        "nhan": "🌅 Buổi sáng (Khối 6 & 9)",
        "tien_to_lop": ("6/", "9/"),
        "khoi_list": [6, 9],
        "ten_buoi": "Buổi sáng",
        "tiet_chao_co": 1,
        "lops_diem_le": ["6/4"],
        "bo_chao_co_lops": ["6/4"],
        "file_excel": "TKB_Sang.xlsx",
    },
    "chieu": {
        "nhan": "🌇 Buổi chiều (Khối 7 & 8)",
        "tien_to_lop": ("7/", "8/"),
        "khoi_list": [7, 8],
        "ten_buoi": "Buổi chiều",
        "tiet_chao_co": 5,
        "lops_diem_le": ["7/5", "8/5"],
        "bo_chao_co_lops": ["7/5"],
        "file_excel": "TKB_Chieu.xlsx",
    },
}
TIEU_DE_KY_HOC = "THỜI KHÓA BIỂU HỌC KỲ I NĂM HỌC 2026-2027"


def _xep_va_luu_mot_buoi(key, phan_cong_list, gv_chu_nhiem, thoi_gian_giai, log):
    """Giải TKB cho một buổi (sang/chieu) và lưu kết quả vào st.session_state."""
    cfg = CAU_HINH_BUOI[key]
    pc_buoi = [i for i in phan_cong_list if i["class"].startswith(cfg["tien_to_lop"])]
    cn_buoi = {l: g for l, g in gv_chu_nhiem.items() if l.startswith(cfg["tien_to_lop"])}

    df, _ = giai_tkb(
        pc_buoi, cn_buoi, cfg["khoi_list"], cfg["ten_buoi"],
        tiet_chao_co=cfg["tiet_chao_co"], lops_diem_le=cfg["lops_diem_le"],
        bo_chao_co_lops=cfg["bo_chao_co_lops"], thoi_gian=thoi_gian_giai, log=log,
    )
    excel_bytes = giai_tkb_va_xuat_excel_bytes(
        pc_buoi, cn_buoi, cfg["khoi_list"], cfg["ten_buoi"], TIEU_DE_KY_HOC,
        tiet_chao_co=cfg["tiet_chao_co"], lops_diem_le=cfg["lops_diem_le"],
        bo_chao_co_lops=cfg["bo_chao_co_lops"], thoi_gian=thoi_gian_giai,
    )
    st.session_state[f"df_{key}"] = df
    st.session_state[f"excel_{key}"] = excel_bytes


if "df_sang" not in st.session_state:
    st.session_state.df_sang = None
if "df_chieu" not in st.session_state:
    st.session_state.df_chieu = None
if "excel_sang" not in st.session_state:
    st.session_state.excel_sang = None
if "excel_chieu" not in st.session_state:
    st.session_state.excel_chieu = None

if uploaded_file is not None:
    try:
        phan_cong_list, gv_chu_nhiem = doc_phan_cong_stream(uploaded_file)
        st.sidebar.success(f"✓ Đã đọc {len(phan_cong_list)} tiết phân công.")

        tab_sang, tab_chieu = st.tabs([
            "🌅 Thời Khóa Biểu BUỔI SÁNG (Khối 6 & 9)",
            "🌇 Thời Khóa Biểu BUỔI CHIỀU (Khối 7 & 8)",
        ])

        st.sidebar.markdown("### Xếp thời khóa biểu")
        nut_ca_hai = st.sidebar.button("🚀 Xếp CẢ HAI buổi")
        col_sang, col_chieu = st.sidebar.columns(2)
        nut_chi_sang = col_sang.button("🌅 Chỉ buổi sáng")
        nut_chi_chieu = col_chieu.button("🌇 Chỉ buổi chiều")
        st.sidebar.caption(
            "Dùng 'Chỉ buổi sáng/chiều' khi giữa học kỳ chỉ có thay đổi ở một "
            "buổi, để không phải giải lại buổi còn lại (tiết kiệm thời gian)."
        )

        if nut_ca_hai or nut_chi_sang or nut_chi_chieu:
            nhat_ky = []
            try:
                if nut_ca_hai or nut_chi_sang:
                    with st.spinner("⚡ Đang xếp TKB buổi sáng (Khối 6, 9)..."):
                        _xep_va_luu_mot_buoi("sang", phan_cong_list, gv_chu_nhiem, thoi_gian_giai, nhat_ky)

                if nut_ca_hai or nut_chi_chieu:
                    with st.spinner("⚡ Đang xếp TKB buổi chiều (Khối 7, 8)..."):
                        _xep_va_luu_mot_buoi("chieu", phan_cong_list, gv_chu_nhiem, thoi_gian_giai, nhat_ky)

                st.sidebar.success("🎉 Đã xếp TKB thành công!")

            except TkbError as e:
                st.sidebar.error(f"❌ {e}")
                if nhat_ky:
                    with st.expander("Xem chi tiết log"):
                        st.text("\n".join(nhat_ky))

        with tab_sang:
            st.subheader("📋 Bảng Thời Khóa Biểu Buổi Sáng")
            if st.session_state.df_sang is not None:
                st.dataframe(st.session_state.df_sang, use_container_width=True, height=500)
                st.download_button(
                    "⬇️ Tải Excel Buổi Sáng", data=st.session_state.excel_sang,
                    file_name="TKB_Sang.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("👈 Bấm nút 'BẮT ĐẦU XẾP THỜI KHÓA BIỂU' ở thanh bên trái để chạy thuật toán.")

        with tab_chieu:
            st.subheader("📋 Bảng Thời Khóa Biểu Buổi Chiều")
            if st.session_state.df_chieu is not None:
                st.dataframe(st.session_state.df_chieu, use_container_width=True, height=500)
                st.download_button(
                    "⬇️ Tải Excel Buổi Chiều", data=st.session_state.excel_chieu,
                    file_name="TKB_Chieu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("👈 Bấm nút 'BẮT ĐẦU XẾP THỜI KHÓA BIỂU' ở thanh bên trái để chạy thuật toán.")

    except TkbError as e:
        st.error(f"❌ {e}")
    except Exception as e:
        st.error(f"❌ Có lỗi khi đọc file hoặc chạy thuật toán: {e}")
else:
    st.warning("👈 Vui lòng tải file Excel Phân công giảng dạy ở thanh bên trái để bắt đầu!")
