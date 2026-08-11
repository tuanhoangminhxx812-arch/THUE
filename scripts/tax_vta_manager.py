# -*- coding: utf-8 -*-
"""
Module quản lý và cập nhật file Excel Master TAX_VTA_YYYY_MM.xlsx.
Chép dữ liệu từ các file đầu vào vào từng sheet tương ứng, bảo toàn toàn bộ công thức Excel.
Định dạng số chuẩn (native int/float) và lọc TRONGTHANG / CUOITHANG chuẩn cho Sheet GCS và Nhom TC.
"""
import io
import os
import shutil
import pandas as pd
import openpyxl
from openpyxl.utils.datetime import from_excel
from pathlib import Path
from datetime import datetime
from utils import load_config, resolve_config_filenames, resolve_file, get_project_root, get_input_dir, get_output_dir


def get_template_file() -> Path:
    """Trả về đường dẫn file mẫu TAX_VTA_2026_07.xlsx gốc."""
    base_dir = get_project_root()
    p1 = base_dir / "ĐẦU VÀO" / "2026-07" / "TAX_VTA_2026_07.xlsx"
    if p1.exists():
        return p1
    p2 = base_dir / "THAM KHẢO" / "TAX_VTA_2026_07.xlsx"
    if p2.exists():
        return p2
    for f in (base_dir / "ĐẦU VÀO").rglob("TAX_VTA*.xlsx"):
        return f
    raise FileNotFoundError("Không tìm thấy file mẫu TAX_VTA_2026_07.xlsx!")


def get_month_tax_vta_file(report_month: str) -> Path:
    """Trả về đường dẫn file TAX_VTA của tháng trong thư mục ĐẦU VÀO/YYYY-MM/."""
    m_clean = report_month.replace("-", "_")
    p = get_project_root() / "ĐẦU VÀO" / report_month / f"TAX_VTA_{m_clean}.xlsx"
    if not p.exists():
        alt = get_project_root() / "ĐẦU VÀO" / report_month / f"TAX_VTA_{m_clean}_latest.xlsx"
        if alt.exists():
            return alt
    return p


def clean_num_val(val):
    """Chuyển đổi chuỗi số (kể cả chứa khoảng trắng, dấu phẩy, dấu chấm phân cách, \xa0) thành float/int chuẩn."""
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        val_str = val.strip()
        if not val_str:
            return None
        if "/" in val_str or ":" in val_str or "Tài khoản" in val_str:
            return val
        if val_str.count("-") > 1 or (len(val_str) >= 10 and val_str.startswith("0") and "-" in val_str):
            return val

        s_clean = val_str.replace(" ", "").replace("\xa0", "")
        if s_clean.count(".") > 1:
            s_clean = s_clean.replace(".", "")
        else:
            s_clean = s_clean.replace(",", "")

        check_str = s_clean[1:] if s_clean.startswith("-") else s_clean
        if check_str.isdigit():
            if len(s_clean) >= 10 and s_clean.startswith("0"):
                return val
            try:
                return int(s_clean)
            except ValueError:
                return val
        try:
            return float(s_clean)
        except ValueError:
            return val
    return val


def set_cell_val(ws, r, c, val):
    """Ghi giá trị an toàn vào cell (tránh lỗi ô gộp ReadOnly) và định dạng số chuẩn."""
    cell = ws.cell(row=r, column=c)
    if not isinstance(cell, openpyxl.cell.cell.MergedCell):
        cell.value = clean_num_val(val)


def clear_range(ws, min_row, min_col, max_row, max_col):
    """Xóa dữ liệu trong vùng chỉ định để chép dữ liệu mới không bị dính nét cũ."""
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            set_cell_val(ws, r, c, None)


def populate_sheet_gcs(wb_master, config: dict, report_month: str):
    """
    Sheet GCS: Chép dữ liệu A1:O29 từ BC_HDon_01_THopTheoNgayGCS chép đè vào A1:O29 của Sheet GCS.
    Lọc TRONGTHANG / CUOITHANG ở cột P (col 16) CHỈ ĐỐI VỚI DÒNG SỐ LIỆU CHI TIẾT (STT là số: 1, 2, 3...).
    Bỏ qua các dòng tiêu đề và dòng Tổng số để công thức SUMIF không bị cộng trùng.
    """
    try:
        fname = resolve_file(config["input_files"]["bchdon"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["GCS"]

    # Copy A1:O29 trực tiếp với định dạng số chuẩn
    for r in range(1, 30):
        for c in range(1, 16):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    # Cập nhật cột P (16) CHỈ cho các dòng số liệu chi tiết (STT là số)
    target_ym = report_month
    for r in range(1, 30):
        stt_val = ws_src.cell(r, 1).value
        stt_str = str(stt_val or "").strip()

        # Kiểm tra xem có phải dòng số liệu chi tiết không (STT là số)
        is_detail_row = stt_str.isdigit()

        if is_detail_row:
            date_val = ws_src.cell(r, 5).value or ws_src.cell(r, 6).value
            p_val = "TRONGTHANG"
            if isinstance(date_val, (datetime, pd.Timestamp)):
                if date_val.strftime("%Y-%m") > target_ym:
                    p_val = "CUOITHANG"
            elif isinstance(date_val, (int, float)):
                try:
                    dt = from_excel(date_val)
                    if dt.strftime("%Y-%m") > target_ym:
                        p_val = "CUOITHANG"
                except Exception:
                    pass
            elif isinstance(date_val, str) and ("2026-08" in date_val or "08/2026" in date_val):
                p_val = "CUOITHANG"

            set_cell_val(ws_dst, r, 16, p_val)
        else:
            # Bỏ qua dòng tiêu đề, "Loại hoá đơn" và dòng "Tổng số"
            set_cell_val(ws_dst, r, 16, None)

    wb_src.close()


def populate_sheet_0903(wb_master, config: dict):
    """
    Sheet 0903: Xóa dữ liệu cũ, lọc các TK 51111, 51113, 5114 từ GL00903_T7.xlsx.
    Chép dữ liệu từ cột A -> U bắt đầu từ ô A1. Bổ sung công thức cột V: =LEFT(R{r},6).
    """
    try:
        fname = resolve_file(config["input_files"]["gl0903"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["0903"]

    target_tks = ("51111", "51113", "5114", "51111000000", "51113000000", "51140000000")
    filtered_rows = []

    for r in range(7, ws_src.max_row + 1):
        row_vals = [ws_src.cell(r, c).value for c in range(1, 22)]
        tk_val = str(ws_src.cell(r, 6).value or "").strip()
        row_str = " ".join([str(v or "") for v in row_vals])
        if any(row_str.startswith(tk) or tk in row_str or tk_val.startswith(tk) for tk in target_tks):
            filtered_rows.append(row_vals)

    max_dst_r = max(ws_dst.max_row, 6 + len(filtered_rows) + 5)
    clear_range(ws_dst, 1, 1, max_dst_r, 22)

    # Chép header 6 dòng đầu A1:U6
    for r in range(1, 7):
        for c in range(1, 22):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    # Chép dữ liệu đã lọc từ dòng 7 trở đi
    curr_row = 7
    for row_vals in filtered_rows:
        for c_idx in range(min(21, len(row_vals))):
            set_cell_val(ws_dst, curr_row, c_idx + 1, row_vals[c_idx])
        set_cell_val(ws_dst, curr_row, 22, f"=LEFT(R{curr_row},6)")
        curr_row += 1

    wb_src.close()


def populate_sheet_4a(wb_master, config: dict):
    """
    Sheet 4A: Chép dữ liệu từ rptKDDN4A.xlsx (dòng A2:J2) đè vào A2:J2 của Sheet 4A.
    """
    try:
        fname = resolve_file(config["input_files"]["kd4a"], config)
    except FileNotFoundError:
        return

    engine = "xlrd" if str(fname).endswith(".xls") else "openpyxl"
    df = pd.read_excel(fname, engine=engine, header=None)

    ws_dst = wb_master["4A"]
    if len(df) >= 2:
        for c in range(min(10, df.shape[1])):
            val = df.iloc[1, c]
            set_cell_val(ws_dst, 2, c + 1, val if pd.notna(val) else None)


def populate_sheet_ta35(wb_master, config: dict):
    """
    Sheet TA35 (T35): Xóa dữ liệu cũ cột A->S (cols 1..19), chép dữ liệu từ TA_035_3331.xlsx (cột A->S bắt đầu ô A1).
    Cột T (20) & U (21): điền công thức =LEFT(G{r},6) và =LEN(F{r}) cho các dòng dữ liệu từ dòng 19 trở đi.
    BẢO TOÀN NGUYÊN VẸN 100% CÁC CỘT TỪ V TRỞ ĐI (CỘT 22+) VÀ CÁC BẢNG TỔNG HỢP CÓ SẴN CỦA ANH.
    """
    try:
        fname = resolve_file(config["input_files"]["ta035"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["TA35"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)

    # CHỈ XÓA TỪ CỘT A (1) ĐẾN CỘT S (19)
    clear_range(ws_dst, 1, 1, max_dst_r, 19)

    # Copy A1:S{src_max_r} từ file nguồn
    for r in range(1, src_max_r + 1):
        for c in range(1, 20):  # Cột 1 đến 19 (A đến S)
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    # Điền công thức cột T (20) và U (21) cho các dòng chi tiết từ dòng 19 trở đi (KHÔNG ĐỤNG ĐẾN CỘT V TRỞ ĐI)
    for r in range(19, src_max_r + 1):
        stt_val = str(ws_src.cell(r, 1).value or "").strip()
        g_val = str(ws_src.cell(r, 7).value or "").strip()
        k_val = str(ws_src.cell(r, 11).value or "").strip()

        if stt_val == "Tổng cộng":
            set_cell_val(ws_dst, r, 20, f"=LEFT(K{r},6)" if k_val else f"=LEFT(G{r},6)")
        elif stt_val or g_val:
            set_cell_val(ws_dst, r, 20, f"=LEFT(G{r},6)")
            set_cell_val(ws_dst, r, 21, f"=LEN(F{r})")

    # Xóa công thức dư cột T & U ở các dòng bên dưới src_max_r
    for r in range(src_max_r + 1, max_dst_r + 1):
        set_cell_val(ws_dst, r, 20, None)
        set_cell_val(ws_dst, r, 21, None)

    wb_src.close()


def populate_sheet_333111(wb_master, config: dict):
    """
    Sheet 333111: Xóa dữ liệu cũ cột A->H, chép từ TA_030_3331.xlsx (cột A->H bắt đầu ô A1).
    Điền công thức cột I: =LEN(A{r}), J: =LEFT(D{r},6), K: =IF(LEFT(A{r},11)="Tài khoản: ",MID(A{r},12,6),K{r-1}).
    """
    try:
        fname = resolve_file(config["input_files"]["ta030_3331"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["333111"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)
    clear_range(ws_dst, 1, 1, max_dst_r, 11)

    for r in range(1, src_max_r + 1):
        for c in range(1, 9):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    for r in range(11, src_max_r + 1):
        set_cell_val(ws_dst, r, 9, f"=LEN(A{r})")
        set_cell_val(ws_dst, r, 10, f"=LEFT(D{r},6)")
        set_cell_val(ws_dst, r, 11, f'=IF(LEFT(A{r},11)="Tài khoản: ",MID(A{r},12,6),K{r-1})')

    wb_src.close()


def populate_sheet_33895(wb_master, config: dict):
    """
    Sheet 33895: Xóa dữ liệu cũ cột A->I, chép từ GL038_33895.xlsx (cột A->I bắt đầu ô A1).
    Phân loại thông minh Cột J (10) cho tất cả các dòng (gồm Nợ giải trừ và Có treo) ngay cả khi diễn giải không chứa tiền tố.
    Bảo toàn 100% công thức M2, M3, N2, N3.
    """
    try:
        fname = resolve_file(config["input_files"]["gl038"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["33895"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)

    # Clear A -> K
    for r in range(1, max_dst_r + 1):
        for c in range(1, 12):
            cell = ws_dst.cell(row=r, column=c)
            if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                cell.value = None

    # Copy A -> I
    for r in range(1, src_max_r + 1):
        for c in range(1, 10):
            val = ws_src.cell(r, c).value
            set_cell_val(ws_dst, r, c, val)

    # Gán mã phân loại chuẩn cho Cột J (col 10) - Bỏ qua các dòng tổng cộng / dư cuối kỳ
    for r in range(15, src_max_r + 1):
        col1 = str(ws_dst.cell(r, 1).value or "").strip()
        col2 = str(ws_dst.cell(r, 2).value or "").strip()
        col5 = str(ws_dst.cell(r, 5).value or "").strip()

        # Bỏ qua các dòng tổng cộng, dư cuối kỳ, dư đầu kỳ, dòng người lập
        is_summary = any(kw in col1.lower() or kw in col2.lower() or kw in col5.lower() for kw in ["cộng", "tổng", "dư cuối", "dư đầu", "người lập", "kế toán"])
        if is_summary:
            continue

        desc1 = col5.upper()
        desc2 = str(ws_dst.cell(r + 1, 5).value or "").strip().upper()
        combined = desc1 + " " + desc2

        no_val = clean_num_val(ws_dst.cell(r, 6).value) or 0
        co_val = clean_num_val(ws_dst.cell(r, 7).value) or 0

        # Chỉ xử lý dòng giao dịch chi tiết thực sự
        if no_val == 0 and co_val == 0 and not ("DIEN" in combined or "CSPK" in combined or "1388" in combined or "5118" in combined or "4500" in combined):
            continue

        code = None
        if "DIEN01" in combined:
            code = "DIEN01"
        elif "CSPK02" in combined:
            code = "CSPK02"
        elif "DIEN00" in combined:
            code = "DIEN00"
        elif "1388" in combined or "ĐIỀU ĐỘNG" in combined:
            code = "1388ĐĐ"
        elif "5118" in combined:
            code = "5118QT"
        elif "4500" in combined:
            code = "4500TT"
        elif no_val > 0:
            if "AR1647" in combined or "AR1782" in combined:
                code = "DIEN01"
            elif "AR1649" in combined or "AR1784" in combined:
                code = "CSPK02"
            elif "AR1648" in combined or "AR1783" in combined:
                code = "DIEN00"
            else:
                if no_val > 1_000_000_000:
                    code = "DIEN01"
                elif no_val > 1000:
                    code = "CSPK02"

        if code:
            set_cell_val(ws_dst, r, 10, code)
        else:
            set_cell_val(ws_dst, r, 10, f"=LEFT(E{r+1},6)")

        set_cell_val(ws_dst, r, 11, f"=IF(H{r}=$K$9,G{r},0)")

    # Đảm bảo các công thức bảng tổng hợp M2, M3, N2, N3
    ws_dst["L1"].value = None
    ws_dst["M1"].value = "Nợ_giải trừ"
    ws_dst["N1"].value = "Có_treo"

    ws_dst["L2"].value = "DIEN01"
    ws_dst["M2"].value = "=SUMIF($J:$J,$L2,F:F)"
    ws_dst["N2"].value = "=SUMIF($J:$J,$L2,G:G)"

    ws_dst["L3"].value = "CSPK02"
    ws_dst["M3"].value = "=SUMIF($J:$J,$L3,F:F)"
    ws_dst["N3"].value = "=SUMIF($J:$J,$L3,G:G)"

    ws_dst["M4"].value = "=SUM(M2:M3)"
    ws_dst["N4"].value = "=SUM(N2:N3)"

    wb_src.close()


def populate_sheet_ta36(wb_master, config: dict):
    """
    Sheet TA36: Xóa dữ liệu cũ cột A->S, chép từ TA_036_1331.xlsx (cột A->S bắt đầu ô A1).
    Điền công thức cột T: =IF(OR(LEN(F{r})=11,LEN(F{r})=15),"Đ","SAI"), U: =ROUND(H{r}*I{r}%,0)-J{r}, V: =IF(Q{r}=$V$12,13313,13311).
    """
    try:
        fname = resolve_file(config["input_files"]["ta036"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["TA36"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)
    # Xóa dữ liệu cột A->S (1->19)
    clear_range(ws_dst, 1, 1, max_dst_r, 19)

    for r in range(1, src_max_r + 1):
        for c in range(1, 20):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    # Đảm bảo ô V12 có giá trị "XDCB"
    ws_dst["V12"].value = "XDCB"

    # Điền lại công thức cho Cột T (20), U (21), V (22) từ hàng 14 đến hết
    for r in range(14, src_max_r + 1):
        set_cell_val(ws_dst, r, 20, f'=IF(OR(LEN(F{r})=11,LEN(F{r})=15),"Đ","SAI")')
        set_cell_val(ws_dst, r, 21, f"=ROUND(H{r}*I{r}%,0)-J{r}")
        set_cell_val(ws_dst, r, 22, f"=IF(Q{r}=$V$12,13313,13311)")

    # Xóa bớt công thức dư thừa ở các dòng sau src_max_r cho T, U, V
    for r in range(src_max_r + 1, max_dst_r + 1):
        set_cell_val(ws_dst, r, 20, None)
        set_cell_val(ws_dst, r, 21, None)
        set_cell_val(ws_dst, r, 22, None)

    # Cập nhật Bảng thống kê theo Người lập W1:AA15
    users = [
        "TUAN2HM.HCM.VTA", "LINH3NTT.HCM.VTA", "THANH5DT.HCM.VTA",
        "HIENNTK.HCM.VTA", "QUYENNN.HCM.VTA", "HUENT.HCM.VTA",
        "THAOPTT.HCM.VTA", "CHAULTN.HCM.VTA", "THUTT.HCM.VTA"
    ]

    for idx, u in enumerate(users, start=1):
        ws_dst.cell(idx, 23).value = u
        ws_dst.cell(idx, 24).value = f"=SUMIF($R:$R,$W{idx},$H:$H)"
        ws_dst.cell(idx, 25).value = f"=SUMIF($R:$R,$W{idx},$J:$J)"
        ws_dst.cell(idx, 26).value = f"=SUMIF('13311'!$G:$G,'TA36'!$W{idx},'13311'!$E:$E)"
        ws_dst.cell(idx, 27).value = f"=Y{idx}-Z{idx}"

    ws_dst.cell(10, 24).value = "=SUM(X1:X9)"
    ws_dst.cell(10, 25).value = "=SUM(Y1:Y9)"
    ws_dst.cell(10, 26).value = "=SUM(Z1:Z9)"
    ws_dst.cell(10, 27).value = "=SUM(AA1:AA9)"

    ws_dst.cell(11, 24).value = "TAX36"
    ws_dst.cell(11, 25).value = "TAX36"
    ws_dst.cell(11, 26).value = "TA030"

    ws_dst.cell(12, 24).value = "trước thuế"
    ws_dst.cell(12, 25).value = "tiền thuế"

    ws_dst.cell(13, 23).value = "SXKD"
    ws_dst.cell(13, 24).value = "=X10-X14"
    ws_dst.cell(13, 25).value = "=Y10-Y14"

    ws_dst.cell(14, 23).value = "XDCB"
    ws_dst.cell(14, 24).value = "=SUMIF($Q:$Q,$V$12,$H:$H)"
    ws_dst.cell(14, 25).value = "=SUMIF($Q:$Q,$V$12,$J:$J)"

    ws_dst.cell(15, 24).value = "=SUM(X13:X14)"
    ws_dst.cell(15, 25).value = "=SUM(Y13:Y14)"

    wb_src.close()


def populate_sheet_13311(wb_master, config: dict):
    """
    Sheet 13311: Xóa dữ liệu cũ cột A->H, chép từ TA_030_1331.xlsx (cột A->H bắt đầu ô A1).
    Điền công thức cột I: =IF(LEN(A{r})=5,E{r},0), J: =IF(LEFT(A{r},11)="Tài khoản: ",MID(A{r},12,5),J{r-1}).
    """
    try:
        fname = resolve_file(config["input_files"]["ta030_1331"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["13311"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)
    clear_range(ws_dst, 1, 1, max_dst_r, 10)

    for r in range(1, src_max_r + 1):
        for c in range(1, 9):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    for r in range(11, src_max_r + 1):
        set_cell_val(ws_dst, r, 9, f"=IF(LEN(A{r})=5,E{r},0)")
        set_cell_val(ws_dst, r, 10, f'=IF(LEFT(A{r},11)="Tài khoản: ",MID(A{r},12,5),J{r-1})')

    wb_src.close()


def populate_sheet_nhomtc(wb_master, config: dict, report_month: str):
    """
    Sheet Nhom TC: Xóa dữ liệu cũ cột A->J, chép từ BC_San luong nhom Toan Cau.xlsx (cột A->J bắt đầu ô A1).
    Lọc TRONGTHANG / CUOITHANG ở cột L (col 12) CHỈ đối với các dòng số liệu chi tiết (bỏ qua dòng tiêu đề và dòng Tổng Cộng).
    """
    try:
        fname = resolve_file(config["input_files"]["nhomtc"], config)
    except FileNotFoundError:
        return

    wb_src = openpyxl.load_workbook(fname, data_only=True)
    ws_src = wb_src.active
    ws_dst = wb_master["Nhom TC"]

    src_max_r = ws_src.max_row
    max_dst_r = max(ws_dst.max_row, src_max_r + 5)
    clear_range(ws_dst, 1, 1, max_dst_r, 12)

    for r in range(1, src_max_r + 1):
        for c in range(1, 11):
            set_cell_val(ws_dst, r, c, ws_src.cell(r, c).value)

    target_ym = report_month
    for r in range(1, src_max_r + 1):
        stt_val = ws_src.cell(r, 1).value
        d_val = ws_src.cell(r, 6).value

        # Dòng số liệu chi tiết: có Ngày phát hành ở cột 6 (dạng datetime, serial number hoặc str ngày tháng)
        is_detail_row = (d_val is not None) and (isinstance(d_val, (datetime, pd.Timestamp, int, float)) or "2026" in str(d_val)) and ("Tổng" not in str(stt_val or ""))

        if is_detail_row:
            p_val = "TRONGTHANG"
            if isinstance(d_val, (datetime, pd.Timestamp)):
                if d_val.strftime("%Y-%m") > target_ym:
                    p_val = "CUOITHANG"
            elif isinstance(d_val, (int, float)):
                try:
                    dt = from_excel(d_val)
                    if dt.strftime("%Y-%m") > target_ym:
                        p_val = "CUOITHANG"
                except Exception:
                    pass
            elif isinstance(d_val, str) and ("2026-08" in d_val or "08/2026" in d_val):
                p_val = "CUOITHANG"

            set_cell_val(ws_dst, r, 11, f"=I{r}")
            set_cell_val(ws_dst, r, 12, p_val)
        else:
            set_cell_val(ws_dst, r, 11, None)
            set_cell_val(ws_dst, r, 12, None)

    wb_src.close()


def get_prev_month_gcs_values(report_month: str):
    """
    Tính toán động số liệu GCS Cuối tháng của tháng T-1 để chuyển sang tháng T (dòng GCS T{m-1} sang T{m} HT).
    F28: Doanh số DIEN00 CUOITHANG
    F29: Doanh số DIEN01 CUOITHANG (TienDien CUOITHANG - DIEN00 CUOITHANG)
    F30: Doanh số CSPK02 CUOITHANG (TienCSPK CUOITHANG)
    """
    try:
        parts = report_month.split("-")
        y = int(parts[0])
        m = int(parts[1])
        m_prev = 12 if m == 1 else m - 1
        y_prev = y - 1 if m == 1 else y
        prev_month_str = f"{y_prev:04d}-{m_prev:02d}"

        cfg = load_config()
        cfg["report_month"] = prev_month_str
        cfg = resolve_config_filenames(cfg)

        from processors.bchdon import process_bchdon
        from processors.nhomtc import process_nhomtc

        bchdon_res = process_bchdon(cfg)
        nhomtc_res = process_nhomtc(cfg)

        b_sum = bchdon_res.get("summary", None)
        n_sum = nhomtc_res.get("summary", None)

        cuoi_b = b_sum[b_sum["THOIGIAN"] == "CUOITHANG"] if b_sum is not None and not b_sum.empty else None
        cuoi_n = n_sum[n_sum["THOIGIAN"] == "CUOITHANG"] if n_sum is not None and not n_sum.empty else None

        dien00 = float(cuoi_n["TongTien"].values[0]) if cuoi_n is not None and not cuoi_n.empty else 0.0
        gcs_dien = float(cuoi_b["TienDien"].values[0]) if cuoi_b is not None and not cuoi_b.empty else 0.0
        gcs_cspk = float(cuoi_b["TienCSPK"].values[0]) if cuoi_b is not None and not cuoi_b.empty else 0.0

        dien01 = gcs_dien - dien00 if gcs_dien > 0 else 0.0
        return dien00, dien01, gcs_cspk
    except Exception:
        return 0.0, 0.0, 0.0


def populate_sheet_taxvta(wb_master, report_month: str):
    """
    Sheet TAXVTA: Cập nhật tiêu đề xanh và các công thức/số liệu liên kết GCS tháng trước (GCS T{m-1} sang T{m} HT)
    cho các tháng từ tháng 6 trở đi. Liên kết nội bộ từ Sheet 33895 để tránh lỗi #REF! do file ngoài [1].
    """
    if "TAXVTA" not in wb_master.sheetnames:
        return
    ws = wb_master["TAXVTA"]

    try:
        parts = report_month.split("-")
        y = int(parts[0])
        m = int(parts[1])
        m_prev = 12 if m == 1 else m - 1
        y_prev = y - 1 if m == 1 else y
        yy_str = str(y)[-2:]
        ws["F26"].value = f"GCS T{m_prev}-{y_prev} sang T{m}/{yy_str} HT"
    except Exception:
        pass

    dien00, dien01, cspk02 = get_prev_month_gcs_values(report_month)
    if dien00 > 0 or dien01 > 0 or cspk02 > 0:
        ws["F28"].value = dien00
        ws["F29"].value = dien01
        ws["F30"].value = cspk02

    ws["G29"].value = "='33895'!M2"
    ws["G30"].value = "='33895'!M3"


def update_tax_vta_file(report_month: str) -> Path:
    """
    Hàm chính: Đọc phôi TAX_VTA_2026_07.xlsx, nạp 9 file đầu vào của report_month vào các sheet,
    lưu kết quả đè vào file TAX_VTA_YYYY_MM.xlsx trong thư mục tháng và thư mục ĐẦU RA.
    """
    config = load_config()
    config["report_month"] = report_month
    config = resolve_config_filenames(config)

    template_file = get_template_file()
    m_clean = report_month.replace("-", "_")
    out_file = get_project_root() / "ĐẦU VÀO" / report_month / f"TAX_VTA_{m_clean}.xlsx"

    out_file.parent.mkdir(parents=True, exist_ok=True)

    with open(template_file, "rb") as f:
        in_mem = io.BytesIO(f.read())

    wb_master = openpyxl.load_workbook(in_mem, data_only=False)

    populate_sheet_gcs(wb_master, config, report_month)
    populate_sheet_0903(wb_master, config)
    populate_sheet_4a(wb_master, config)
    populate_sheet_ta35(wb_master, config)
    populate_sheet_333111(wb_master, config)
    populate_sheet_33895(wb_master, config)
    populate_sheet_ta36(wb_master, config)
    populate_sheet_13311(wb_master, config)
    populate_sheet_nhomtc(wb_master, config, report_month)
    populate_sheet_taxvta(wb_master, report_month)

    final_saved_file = out_file
    try:
        wb_master.save(out_file)
    except PermissionError:
        alt_file = out_file.parent / f"TAX_VTA_{m_clean}_latest.xlsx"
        wb_master.save(alt_file)
        final_saved_file = alt_file

    output_dir = get_output_dir()
    output_dir.mkdir(parents=True, exist_ok=True)
    out_file_output_dir = output_dir / f"TAX_VTA_{m_clean}.xlsx"
    try:
        shutil.copy2(final_saved_file, out_file_output_dir)
    except Exception:
        pass

    return final_saved_file


def ensure_tax_vta_file(report_month: str) -> Path:
    """
    Đảm bảo file TAX_VTA của tháng tồn tại. Nếu chưa có, tự động tạo và cập nhật.
    """
    target_file = get_month_tax_vta_file(report_month)
    if not target_file.exists():
        return update_tax_vta_file(report_month)
    return target_file


def get_tax_vta_summary_data(report_month: str) -> dict:
    """
    Trích xuất và tính toán dữ liệu tổng hợp từ các sheet của file master TAX_VTA_YYYY_MM.xlsx.
    Bảo đảm tính chính xác 100% cho Streamlit UI ngay cả khi Excel chưa lưu cached values.
    """
    file_path = ensure_tax_vta_file(report_month)
    
    with open(file_path, "rb") as f:
        in_mem = io.BytesIO(f.read())
    wb = openpyxl.load_workbook(in_mem, data_only=False)

    # 1. Tính Mua Vào từ TA36
    ws_ta36 = wb["TA36"]
    ds_13311, thue_13311 = 0.0, 0.0
    ds_13313, thue_13313 = 0.0, 0.0
    v12_val = str(ws_ta36.cell(12, 22).value or "").strip()

    for r in range(14, ws_ta36.max_row + 1):
        stt = str(ws_ta36.cell(r, 1).value or "").strip()
        if not stt or stt.startswith("Tổng") or stt.startswith("Cộng") or stt.startswith("1.") or stt.startswith("2."):
            continue
        try:
            ds = float(ws_ta36.cell(r, 8).value or 0)
            thue = float(ws_ta36.cell(r, 10).value or 0)
        except (ValueError, TypeError):
            continue

        q_val = str(ws_ta36.cell(r, 17).value or "").strip()
        if q_val and v12_val and q_val == v12_val:
            ds_13313 += ds
            thue_13313 += thue
        else:
            ds_13311 += ds
            thue_13311 += thue

    ds_mua_vao_total = ds_13311 + ds_13313
    thue_mua_vao_total = thue_13311 + thue_13313

    # 2. Tính Bán Ra GCS (Chỉ lấy các dòng chi tiết mà cột P == "TRONGTHANG" hoặc "CUOITHANG")
    ws_gcs = wb["GCS"]
    ds_ban_trong_thang, thue_ban_trong_thang = 0.0, 0.0
    ds_ban_cuoi_thang, thue_ban_cuoi_thang = 0.0, 0.0

    for r in range(1, 30):
        stt_str = str(ws_gcs.cell(r, 1).value or "").strip()
        if not stt_str.isdigit():
            continue

        p_val = str(ws_gcs.cell(r, 16).value or "").strip()
        try:
            ds = float(ws_gcs.cell(r, 9).value or 0)
            thue = float(ws_gcs.cell(r, 10).value or 0)
        except (ValueError, TypeError):
            continue

        if p_val == "TRONGTHANG":
            ds_ban_trong_thang += ds
            thue_ban_trong_thang += thue
        elif p_val == "CUOITHANG":
            ds_ban_cuoi_thang += ds
            thue_ban_cuoi_thang += thue

    ds_ban_gcs_total = ds_ban_trong_thang + ds_ban_cuoi_thang
    thue_ban_gcs_total = thue_ban_trong_thang + thue_ban_cuoi_thang

    # 3. Tính phân loại TA35
    ws_ta35 = wb["TA35"]
    codes = ["DIEN00", "DIEN01", "CSPK02", "1388ĐĐ", "5118QT", "4500TT"]
    ta35_dict = {c: {"doanh_so": 0.0, "thue": 0.0} for c in codes}

    for r in range(17, ws_ta35.max_row + 1):
        stt = str(ws_ta35.cell(r, 1).value or "").strip()
        g_val = str(ws_ta35.cell(r, 7).value or "").strip()
        k_val = str(ws_ta35.cell(r, 11).value or "").strip()
        code_str = g_val[:6] if len(g_val) >= 6 else k_val[:6]

        try:
            ds = float(ws_ta35.cell(r, 8).value or 0)
            thue = float(ws_ta35.cell(r, 10).value or 0)
        except (ValueError, TypeError):
            continue

        if stt and not stt.startswith("Tổng") and not stt.startswith("Cộng") and not stt.startswith("1.") and not stt.startswith("2.") and not stt.startswith("3.") and not stt.startswith("4.") and not stt.startswith("5."):
            for target_code in codes:
                if code_str.startswith(target_code):
                    ta35_dict[target_code]["doanh_so"] += ds
                    ta35_dict[target_code]["thue"] += thue
                    break

    ta35_items = [{"code": c, "doanh_so": ta35_dict[c]["doanh_so"], "thue": ta35_dict[c]["thue"]} for c in codes]
    ta35_total_ds = sum(item["doanh_so"] for item in ta35_items)
    ta35_total_thue = sum(item["thue"] for item in ta35_items)

    thue_phai_nop = ta35_total_thue - thue_mua_vao_total

    wb.close()

    return {
        "report_month": report_month,
        "file_path": str(file_path),
        "mua_vao": {
            "13311": {"doanh_so": ds_13311, "thue": thue_13311},
            "13313": {"doanh_so": ds_13313, "thue": thue_13313},
            "total": {"doanh_so": ds_mua_vao_total, "thue": thue_mua_vao_total},
        },
        "ban_ra": {
            "trong_thang": {"doanh_so": ds_ban_trong_thang, "thue": thue_ban_trong_thang},
            "cuoi_thang": {"doanh_so": ds_ban_cuoi_thang, "thue": thue_ban_cuoi_thang},
            "total_gcs": {"doanh_so": ds_ban_gcs_total, "thue": thue_ban_gcs_total},
            "ta35_items": ta35_items,
            "ta35_total": {"doanh_so": ta35_total_ds, "thue": ta35_total_thue},
        },
        "summary_cards": {
            "thue_dau_ra": ta35_total_thue,
            "thue_dau_vao": thue_mua_vao_total,
            "thue_phai_nop": thue_phai_nop,
        }
    }
