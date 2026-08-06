# -*- coding: utf-8 -*-
"""
Xử lý TA_036_TK1331 (Bảng kê HĐ mua vào).
- Thêm cột MST_CHECK: kiểm tra MST ≠ 10 ký tự → cảnh báo
- Thêm cột TAIKHOAN: "ĐẦU TƯ XÂY DỰNG" → TK13313, còn lại → TK13311
- Group by TAIKHOAN → SUM(ThueGTGT)
"""
import pandas as pd
import re
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from data_loader import load_ta036
from utils import clean_string


def validate_mst(mst_value: str, config: dict) -> dict:
    """
    Kiểm tra tính hợp lệ của MST.
    Returns dict: {'valid': bool, 'message': str, 'cleaned': str}
    """
    mst = clean_string(mst_value)

    if not mst:
        return {"valid": False, "message": "Trống", "cleaned": ""}

    # Loại bỏ dấu ' ở đầu (từ Excel)
    mst = mst.lstrip("'")

    # Tách phần chính và chi nhánh (nếu có)
    parts = mst.split("-")
    main_part = parts[0].strip()

    # Chỉ lấy số
    digits_only = re.sub(r'[^0-9]', '', main_part)

    valid_length = config.get("mst_rules", {}).get("valid_length", 10)

    if len(digits_only) == valid_length:
        return {"valid": True, "message": "✓ OK", "cleaned": mst}
    elif len(digits_only) < valid_length:
        return {"valid": False, "message": f"⚠ Thiếu ({len(digits_only)} ký tự)", "cleaned": mst}
    else:
        return {"valid": False, "message": f"⚠ Thừa ({len(digits_only)} ký tự)", "cleaned": mst}


def process_ta036(config: dict) -> dict:
    """
    Xử lý file TA_036.
    Returns dict:
      - 'detail': DataFrame chi tiết (có MST_CHECK, TAIKHOAN)
      - 'summary': Group by TAIKHOAN → SUM(ThueGTGT, DoanhSoChuaThue)
      - 'summary_thue_suat': Group by ThueSuat → SUM(DoanhSoChuaThue, ThueGTGT)
      - 'mst_warnings': DataFrame các dòng MST sai
    """
    df = load_ta036(config)

    if df.empty:
        return {
            "detail": df,
            "summary": pd.DataFrame(),
            "summary_thue_suat": pd.DataFrame(),
            "mst_warnings": pd.DataFrame()
        }

    # Thêm cột MST_CHECK
    mst_results = df["MSTNguoiBan"].apply(lambda x: validate_mst(x, config))
    df["MST_CHECK"] = [r["message"] for r in mst_results]
    df["MST_VALID"] = [r["valid"] for r in mst_results]

    # Thêm cột TAIKHOAN dựa trên cột LoaiHinhSXKD
    xdcb_keyword = config.get("ta036_account_rules", {}).get("xdcb_keyword", "ĐẦU TƯ XÂY DỰNG")
    xdcb_account = config.get("ta036_account_rules", {}).get("xdcb_account", "TK13313")
    default_account = config.get("ta036_account_rules", {}).get("default_account", "TK13311")

    def assign_account(row):
        loai_hinh = clean_string(row.get("LoaiHinhSXKD", ""))
        if xdcb_keyword.upper() in loai_hinh.upper():
            return xdcb_account
        return default_account

    df["TAIKHOAN"] = df.apply(assign_account, axis=1)

    # Group by TAIKHOAN
    summary = df.groupby("TAIKHOAN").agg(
        DoanhSoChuaThue=("DoanhSoChuaThue", "sum"),
        ThueGTGT=("ThueGTGT", "sum"),
        SoLuong=("STT", "count")
    ).reset_index()

    # Thêm dòng tổng
    total = pd.DataFrame([{
        "TAIKHOAN": "Tổng cộng",
        "DoanhSoChuaThue": summary["DoanhSoChuaThue"].sum(),
        "ThueGTGT": summary["ThueGTGT"].sum(),
        "SoLuong": summary["SoLuong"].sum()
    }])
    summary = pd.concat([summary, total], ignore_index=True)

    # Group by ThueSuat
    summary_ts = df.groupby("ThueSuat").agg(
        DoanhSoChuaThue=("DoanhSoChuaThue", "sum"),
        ThueGTGT=("ThueGTGT", "sum"),
        SoLuong=("STT", "count")
    ).reset_index()

    total_ts = pd.DataFrame([{
        "ThueSuat": "Tổng",
        "DoanhSoChuaThue": summary_ts["DoanhSoChuaThue"].sum(),
        "ThueGTGT": summary_ts["ThueGTGT"].sum(),
        "SoLuong": summary_ts["SoLuong"].sum()
    }])
    summary_ts = pd.concat([summary_ts, total_ts], ignore_index=True)

    # Danh sách MST cảnh báo
    mst_warnings = df[~df["MST_VALID"]].copy()

    return {
        "detail": df,
        "summary": summary,
        "summary_thue_suat": summary_ts,
        "mst_warnings": mst_warnings
    }
