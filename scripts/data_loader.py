# -*- coding: utf-8 -*-
"""
Module đọc và parse dữ liệu từ các file Excel đầu vào.
Mỗi file có cấu trúc header riêng, module này chuẩn hóa chúng.
"""
import pandas as pd
from utils import read_excel_file, read_excel_sheets, clean_string, safe_float, get_input_dir


def load_ta035(config: dict) -> pd.DataFrame:
    """
    Đọc TA_035_TK3331 (Bảng kê HĐ bán ra).
    Header: row 13 (STT, Ký hiệu HĐ, Số HĐ, Ngày phát hành, ...)
    Dữ liệu bắt đầu từ các section (1., 2., 3., 4.)
    """
    fname = config["input_files"]["ta035"]
    df = read_excel_file(fname, config, header=None)

    # Tìm header row (row chứa "STT")
    header_row = None
    for i in range(min(20, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "STT" in vals or "1" in vals:
            # Row 13 là header chính, row 15 là số thứ tự cột [1] [2] ...
            if any(v.startswith("[") or v == "1" for v in vals):
                header_row = i
                break

    # Lấy tên cột từ row 13
    col_names_row = None
    for i in range(min(20, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "STT" in vals and "Mặt hàng" in ' '.join(vals):
            col_names_row = i
            break

    if col_names_row is None:
        col_names_row = 13

    # Đặt tên cột
    columns = [
        "STT", "KyHieuHD", "SoHD", "NgayPhatHanh",
        "TenNguoiMua", "MSTNguoiMua", "MatHang",
        "DoanhSoChuaThue", "ThueSuat", "ThueGTGT", "GhiChu",
        "NguonCT", "SoCT", "NgayLapCT", "ThoiHanTT",
        "SoTienDaTT", "LoaiHinhKD", "NguoiLapCT", "TrangThai"
    ]

    # Lọc các dòng dữ liệu (bỏ header, section header, tổng cộng, footer)
    data_rows = []
    data_started = False
    for i in range(len(df)):
        first_val = clean_string(df.iloc[i, 0])

        # Bỏ qua header rows
        if i <= col_names_row + 2:
            continue

        # Dòng section header (bắt đầu bằng "1.", "2.", "3.", "4." + text mô tả)
        if first_val.startswith(("1.", "2.", "3.", "4.")) and len(first_val) > 3:
            data_started = True
            continue

        # Dòng "Tổng cộng"
        if "tổng cộng" in first_val.lower():
            continue

        # Dòng footer/ký tên
        if "người nộp thuế" in first_val.lower() or "đại diện" in first_val.lower():
            break

        # Dòng trống
        if not first_val:
            continue

        # Kiểm tra xem STT có phải là số
        try:
            int(first_val)
            data_rows.append(i)
        except ValueError:
            continue

    if not data_rows:
        return pd.DataFrame(columns=columns)

    result = df.iloc[data_rows].copy()
    result.columns = columns[:len(result.columns)]
    result = result.reset_index(drop=True)

    # Chuyển đổi kiểu dữ liệu
    for col in ["DoanhSoChuaThue", "ThueGTGT", "SoTienDaTT"]:
        if col in result.columns:
            result[col] = result[col].apply(safe_float)

    return result


def load_ta030(config: dict, tk_key: str) -> pd.DataFrame:
    """
    Đọc TA_030 (Sổ chi tiết theo TK thuế).
    tk_key: "ta030_3331" hoặc "ta030_1331"
    Header: row 8-9 (Ngày CT, Số CT, Nguồn CT, Diễn giải, PS Nợ, PS Có)
    Dữ liệu phân theo marker "Tài khoản: XXXXX"
    """
    fname = config["input_files"][tk_key]
    df = read_excel_file(fname, config, header=None)

    columns = ["NgayCT", "SoCT", "NguonCT", "DienGiai", "PSNo", "PSCo", "NguoiHT", "TrangThai"]

    # Parse dữ liệu, giữ thông tin TK từ marker rows
    records = []
    current_tk = ""
    du_dau_ky = {}

    for i in range(len(df)):
        first_val = clean_string(df.iloc[i, 0])
        col3_val = clean_string(df.iloc[i, 3]) if df.shape[1] > 3 else ""

        # Marker row: "Tài khoản: 33311100000" hoặc "Tài khoản: 13311000000"
        if first_val.startswith("Tài khoản:") or (first_val.startswith("Tài khoản:") is False and "Tài khoản:" in first_val):
            # Trích xuất mã TK
            tk_text = first_val.replace("Tài khoản:", "").strip()
            if tk_text:
                # Lấy 5 ký tự đầu
                tk_num = ''.join(c for c in tk_text if c.isdigit())[:5]
                current_tk = f"TK{tk_num}"

            # Kiểm tra dư đầu kỳ trên cùng dòng
            if "Dư đầu kỳ" in col3_val or "dư đầu kỳ" in col3_val.lower():
                du_dau = safe_float(df.iloc[i, 4])
                du_dau_co = safe_float(df.iloc[i, 5])
                du_dau_ky[current_tk] = {"no": du_dau, "co": du_dau_co}
            continue

        # Dòng "Dư đầu kỳ" riêng
        if "dư đầu kỳ" in col3_val.lower():
            du_dau = safe_float(df.iloc[i, 4])
            du_dau_co = safe_float(df.iloc[i, 5])
            du_dau_ky[current_tk] = {"no": du_dau, "co": du_dau_co}
            continue

        # Bỏ qua header, footer, dòng tổng
        if not current_tk:
            continue
        if "cộng phát sinh" in col3_val.lower() or "dư cuối kỳ" in col3_val.lower():
            # Lưu thông tin tổng
            records.append({
                "TaiKhoan": current_tk,
                "NgayCT": None,
                "SoCT": "",
                "NguonCT": "",
                "DienGiai": col3_val,
                "PSNo": safe_float(df.iloc[i, 4]),
                "PSCo": safe_float(df.iloc[i, 5]),
                "NguoiHT": "",
                "TrangThai": "",
                "LoaiDong": "TONG" if "cộng" in col3_val.lower() else "DUOICK"
            })
            continue
        if "TP. Hồ Chí Minh" in first_val or "EVN_TA_030" in first_val:
            break

        # Dòng dữ liệu bình thường
        ngay = df.iloc[i, 0]
        if pd.isna(ngay):
            continue

        # Kiểm tra xem có phải ngày không
        try:
            if isinstance(ngay, str) and ('/' in ngay or '-' in ngay):
                pass  # OK
            elif hasattr(ngay, 'date'):
                pass  # datetime
            else:
                continue
        except Exception:
            continue

        records.append({
            "TaiKhoan": current_tk,
            "NgayCT": ngay,
            "SoCT": clean_string(df.iloc[i, 1]),
            "NguonCT": clean_string(df.iloc[i, 2]),
            "DienGiai": clean_string(df.iloc[i, 3]),
            "PSNo": safe_float(df.iloc[i, 4]),
            "PSCo": safe_float(df.iloc[i, 5]),
            "NguoiHT": clean_string(df.iloc[i, 6]) if df.shape[1] > 6 else "",
            "TrangThai": clean_string(df.iloc[i, 7]) if df.shape[1] > 7 else "",
            "LoaiDong": "DATA"
        })

    result = pd.DataFrame(records)
    return result, du_dau_ky


def load_ta036(config: dict) -> pd.DataFrame:
    """
    Đọc TA_036_TK1331 (Bảng kê HĐ mua vào).
    Header: row 13-15
    """
    fname = config["input_files"]["ta036"]
    df = read_excel_file(fname, config, header=None)

    columns = [
        "STT", "KyHieuHD", "SoHD", "NgayPhatHanh",
        "TenNguoiBan", "MSTNguoiBan", "MatHang",
        "DoanhSoChuaThue", "ThueSuat", "ThueGTGT", "GhiChu",
        "NguonCT", "SoCT", "NgayLapCT", "ThoiHanTT",
        "SoTienDaTT", "LoaiHinhSXKD", "NguoiLapCT", "TrangThai"
    ]

    # Tìm header row
    col_names_row = 13
    for i in range(min(20, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "STT" in vals and any("Mặt hàng" in v for v in vals):
            col_names_row = i
            break

    # Tìm dòng dữ liệu
    data_rows = []
    for i in range(col_names_row + 3, len(df)):
        first_val = clean_string(df.iloc[i, 0])

        # Section headers
        if first_val.startswith(("1.", "2.", "3.", "4.")) and len(first_val) > 3:
            continue

        # Tổng cộng
        if "tổng cộng" in first_val.lower() or "cộng" == first_val.lower():
            continue

        # Footer
        if "người nộp thuế" in first_val.lower() or first_val == "":
            if first_val == "":
                continue
            break

        # Dòng dữ liệu (STT là số)
        try:
            int(first_val)
            data_rows.append(i)
        except ValueError:
            continue

    if not data_rows:
        return pd.DataFrame(columns=columns)

    result = df.iloc[data_rows].copy()
    result.columns = columns[:len(result.columns)]
    result = result.reset_index(drop=True)

    # Chuyển đổi kiểu dữ liệu
    for col in ["DoanhSoChuaThue", "ThueGTGT", "SoTienDaTT"]:
        if col in result.columns:
            result[col] = result[col].apply(safe_float)

    return result


def load_gl038(config: dict) -> pd.DataFrame:
    """
    Đọc GL_038_TK33895 (Sổ chi tiết TK33895).
    Header: row 6-8
    """
    fname = config["input_files"]["gl038"]
    df = read_excel_file(fname, config, header=None)

    records = []
    du_dau_ky = 0.0

    for i in range(len(df)):
        col0 = clean_string(df.iloc[i, 0])
        col4 = clean_string(df.iloc[i, 4]) if df.shape[1] > 4 else ""

        # Dòng dư đầu kỳ
        if "số dư đầu kỳ" in col4.lower() or "dư đầu kỳ" in col4.lower():
            du_dau_ky = safe_float(df.iloc[i, 6])  # PS Có
            continue

        # Dòng tổng
        if "tổng" in col0.lower() or "cộng" in col0.lower():
            continue
        if "tổng dư cuối kỳ" in col0.lower() or "Tổng Dư cuối kỳ" in col0:
            continue

        # Bỏ header/footer
        if "TỔNG CÔNG TY" in col0 or "SỔ CHI TIẾT" in col0:
            continue
        if "Nguồn bút toán" in col0 or "Tài khoản" in col0:
            continue
        if "33895" in col0 and "Phải trả" in col0:
            continue
        if "Khác" in col0.strip() or col0.strip() == "":
            continue
        if "Người lập" in col0 or "(Ký, họ tên)" in col0:
            break
        if "Từ ngày" in col0 or "Tất cả" in col0:
            continue

        # Dòng dữ liệu: Nguồn = GL-Manual, etc.
        if col0.startswith(("GL", "AR", "AP")):
            ngay = df.iloc[i, 1]
            # Diễn giải có thể nằm ở row hiện tại hoặc row kế tiếp
            dien_giai = col4
            if not dien_giai and i + 1 < len(df):
                dien_giai = clean_string(df.iloc[i + 1, 4])

            records.append({
                "NguonBT": col0,
                "NgayCT": ngay,
                "SoCT_PH": clean_string(df.iloc[i, 2]),
                "SoCT_GL": clean_string(df.iloc[i, 3]),
                "DienGiai": dien_giai,
                "PSNo": safe_float(df.iloc[i, 5]),
                "PSCo": safe_float(df.iloc[i, 6]),
                "NguoiLap": clean_string(df.iloc[i, 7]) if df.shape[1] > 7 else "",
                "TrangThai": clean_string(df.iloc[i, 8]) if df.shape[1] > 8 else "",
            })

    result = pd.DataFrame(records)
    return result, du_dau_ky


def load_gl0903(config: dict) -> pd.DataFrame:
    """
    Đọc GL_0903_TK511 (Báo cáo giao dịch TK511 - doanh thu).
    Header: row 5
    """
    fname = config["input_files"]["gl0903"]
    df = read_excel_file(fname, config, header=None)

    columns = [
        "NguonPS", "SoGD", "NgayGD", "ChiNhanh", "TTCP",
        "TaiKhoan", "LoaiHinh", "SanPham", "YeuTo", "DVNB",
        "DPDonVi", "DPEVN", "LoaiTien", "NoNT", "CoNT",
        "NoQD", "CoQD", "NoiDung", "TrangThai", "NguoiTao", "NguoiCN"
    ]

    # Tìm header row
    header_row = 5
    for i in range(min(10, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "Nguồn phát sinh" in vals:
            header_row = i
            break

    # Lọc dòng dữ liệu
    data_rows = []
    for i in range(header_row + 1, len(df)):
        first_val = clean_string(df.iloc[i, 0])
        if first_val.startswith(("AR_", "GL", "AP_")):
            data_rows.append(i)
        elif "tổng cộng" in first_val.lower():
            break

    if not data_rows:
        return pd.DataFrame(columns=columns)

    result = df.iloc[data_rows].copy()
    result.columns = columns[:len(result.columns)]
    result = result.reset_index(drop=True)

    # Chuyển đổi số
    for col in ["CoNT", "NoNT", "CoQD", "NoQD"]:
        if col in result.columns:
            result[col] = result[col].apply(safe_float)

    return result


def load_bchdon(config: dict) -> pd.DataFrame:
    """
    Đọc BC_HDon_01_THopTheoNgayGCS - Sheet BC.
    Header: row 7-8 (trong sheet BC)
    """
    fname = config["input_files"]["bchdon"]
    filepath = get_input_dir(config) / fname
    xls = pd.ExcelFile(filepath, engine="openpyxl")

    # Đọc sheet BC
    sheet_name = "BC" if "BC" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine="openpyxl")

    columns_map = {
        0: "STT", 1: "NgayGhi", 2: "SoSoGCS", 3: "SoHoaDon",
        4: "NgayPhatHanhHD", 5: "DNTT_SH", 6: "DNTT_NgoaiSH", 7: "DNTT_Tong",
        8: "TienDien", 9: "ThueDien", 10: "TongTienDien",
        11: "TienCSPK", 12: "ThueCSPK", 13: "TongCSPK", 14: "TongCong"
    }

    # Tìm header row
    header_row = 7
    for i in range(min(15, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "Stt" in vals or "STT" in vals:
            header_row = i
            break

    # Lọc dòng dữ liệu (là số hoặc có dữ liệu ngày)
    data_rows = []
    for i in range(header_row + 2, len(df)):
        first_val = clean_string(df.iloc[i, 0])
        # Bỏ header loại HĐ
        if "loại hoá đơn" in first_val.lower() or "loại hóa đơn" in first_val.lower():
            continue
        # Bỏ tổng số
        if "tổng số" in first_val.lower():
            continue
        # Bỏ dòng trống
        if not first_val:
            continue
        # Bỏ footer
        if any(kw in first_val.lower() for kw in ["nv lập", "tổ trưởng", "tp kinh doanh", "phạm", "nguyễn", "đặng"]):
            break
        # Dòng dữ liệu (STT là số)
        try:
            int(first_val)
            data_rows.append(i)
        except ValueError:
            continue

    if not data_rows:
        return pd.DataFrame()

    result = df.iloc[data_rows].copy()
    result = result.rename(columns=columns_map)
    result = result.reset_index(drop=True)

    # Chuyển đổi số
    num_cols = ["TienDien", "ThueDien", "TongTienDien", "TienCSPK", "ThueCSPK", "TongCSPK", "TongCong"]
    for col in num_cols:
        if col in result.columns:
            result[col] = result[col].apply(safe_float)

    return result


def load_nhomtc(config: dict) -> pd.DataFrame:
    """
    Đọc BC Sản lượng Nhôm Toàn Cầu - sheet cuối cùng.
    Header: row 7
    """
    fname = config["input_files"]["nhomtc"]
    filepath = get_input_dir(config) / fname
    xls = pd.ExcelFile(filepath, engine="openpyxl")

    # Lấy sheet cuối cùng
    last_sheet = xls.sheet_names[-1]
    df = pd.read_excel(filepath, sheet_name=last_sheet, header=None, engine="openpyxl")

    # Tìm header row
    header_row = 7
    for i in range(min(15, len(df))):
        vals = [clean_string(v) for v in df.iloc[i].values]
        if "STT" in vals:
            header_row = i
            break

    columns_map = {
        0: "STT", 1: "MaKH", 2: "DanhSo", 3: "DiaChi",
        4: "KyThangNam", 5: "NgayPhatHanh", 6: "LoaiHD",
        7: "DienTieuThu", 8: "TienPS", 9: "ThuePS", 10: "TongTien"
    }

    # Lọc dòng dữ liệu
    data_rows = []
    for i in range(header_row + 1, len(df)):
        first_val = clean_string(df.iloc[i, 0])
        col4 = clean_string(df.iloc[i, 4]) if df.shape[1] > 4 else ""

        # Dòng tổng cộng
        if "tổng cộng" in first_val.lower():
            continue
        # Footer
        if not first_val and not col4:
            # Kiểm tra xem dòng có dữ liệu không
            has_data = any(pd.notna(df.iloc[i, j]) and clean_string(df.iloc[i, j]) for j in range(df.shape[1]))
            if not has_data:
                continue
        # Dòng ký tên
        if any(kw in first_val for kw in ["Phạm", "Nguyễn", "Đặng", "Lại"]):
            break

        # Dòng dữ liệu - có thể STT là số hoặc trống (dòng tiếp theo của cùng KH)
        has_amount = pd.notna(df.iloc[i, 10]) if df.shape[1] > 10 else False
        if has_amount or pd.notna(df.iloc[i, 8]):
            data_rows.append(i)

    if not data_rows:
        return pd.DataFrame()

    result = df.iloc[data_rows].copy()
    result = result.rename(columns={k: v for k, v in columns_map.items() if k < result.shape[1]})
    result = result.reset_index(drop=True)

    # Chuyển đổi số
    for col in ["TienPS", "ThuePS", "TongTien"]:
        if col in result.columns:
            result[col] = result[col].apply(safe_float)

    return result


def load_kd4a(config: dict) -> pd.DataFrame:
    """
    Đọc rptKDDN4A (Tổng hợp bán điện).
    Hỗ trợ 2 format:
      - Format báo cáo (11 cột): Row 0=Tiêu đề, Row 1=STT cột, Row 2=Data
      - Format database (10 cột): Row 0=Header (MADONVI,...), Row 1=Data
    """
    fname = config["input_files"]["kd4a"]
    df = read_excel_file(fname, config, header=None)

    if df.empty:
        return {}

    # Phát hiện format dựa trên row 0
    first_val = str(df.iloc[0, 0]).strip().upper() if pd.notna(df.iloc[0, 0]) else ""

    if first_val == "MADONVI":
        # === Format database: Row 0 = Header, Row 1 = Data ===
        if len(df) < 2:
            return {}
        row = df.iloc[1]
        result = {
            "TenDonVi": clean_string(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
            "DNTT": safe_float(row.iloc[2]),           # TONGDTP
            "TongDoanhThu": safe_float(row.iloc[3]),    # TIENDIEN (tổng tiền điện)
            "DoanhThu_DTDL": safe_float(row.iloc[3]),   # TIENDIEN = DT ĐTDL (giống TongDoanhThu)
            "DoanhThu_CSPK": safe_float(row.iloc[4]),   # TIENCSPK
            "GiaBQ": safe_float(row.iloc[9]) if len(row) > 9 else 0,  # GIABQKH
            "TongThue": safe_float(row.iloc[5]),        # THUE_TD
            "Thue_DTDL": safe_float(row.iloc[5]),       # THUE_TD
            "Thue_CSPK": safe_float(row.iloc[6]),       # THUE_CSPK
            "TongCong": safe_float(row.iloc[3]) + safe_float(row.iloc[5]),  # TIENDIEN + THUE_TD
        }
        return result
    else:
        # === Format báo cáo gốc: Row 2 = Data ===
        if len(df) < 3:
            return {}
        row = df.iloc[2]
        result = {
            "TenDonVi": clean_string(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
            "DNTT": safe_float(row.iloc[2]),
            "TongDoanhThu": safe_float(row.iloc[3]),
            "DoanhThu_DTDL": safe_float(row.iloc[4]),
            "DoanhThu_CSPK": safe_float(row.iloc[5]),
            "GiaBQ": safe_float(row.iloc[6]),
            "TongThue": safe_float(row.iloc[7]),
            "Thue_DTDL": safe_float(row.iloc[8]),
            "Thue_CSPK": safe_float(row.iloc[9]),
            "TongCong": safe_float(row.iloc[10]),
        }
        return result

