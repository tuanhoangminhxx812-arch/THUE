# -*- coding: utf-8 -*-
"""
Xử lý TA_030 (Sổ chi tiết theo TK thuế).
- TA_030_TK3331: thêm PHANLOAI, TAIKHOAN, group by PHANLOAI → SUM(PSCo)
- TA_030_TK1331: tách TK13311 và TK13313, thêm TAIKHOAN
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from classifier import Classifier
from data_loader import load_ta030


def process_ta030_3331(config: dict) -> dict:
    """
    Xử lý TA_030_TK3331 (thuế đầu ra).
    Returns dict:
      - 'detail': DataFrame chi tiết
      - 'summary': Group by PHANLOAI + TAIKHOAN → SUM(PSCo)
      - 'du_dau_ky': dict dư đầu kỳ theo TK
    """
    clf = Classifier(config)
    df, du_dau_ky = load_ta030(config, "ta030_3331")

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame(), "du_dau_ky": du_dau_ky}

    # Thêm cột PHANLOAI cho dòng DATA
    df["PHANLOAI"] = df["DienGiai"].apply(clf.classify)

    # Lọc chỉ dòng DATA để group
    data_df = df[df["LoaiDong"] == "DATA"].copy()

    # Group by TAIKHOAN + PHANLOAI
    summary = data_df.groupby(["TaiKhoan", "PHANLOAI"]).agg(
        PSCo=("PSCo", "sum"),
        PSNo=("PSNo", "sum"),
        SoLuong=("DienGiai", "count")
    ).reset_index()

    # Sắp xếp
    code_order = clf.get_all_codes()
    summary["_order"] = summary["PHANLOAI"].apply(
        lambda x: code_order.index(x) if x in code_order else len(code_order)
    )
    summary = summary.sort_values(["TaiKhoan", "_order"]).drop(columns=["_order"]).reset_index(drop=True)

    return {"detail": df, "summary": summary, "du_dau_ky": du_dau_ky}


def process_ta030_1331(config: dict) -> dict:
    """
    Xử lý TA_030_TK1331 (thuế đầu vào).
    Tách thành TK13311 và TK13313.
    Returns dict:
      - 'detail': DataFrame chi tiết
      - 'detail_13311': DataFrame chỉ TK13311
      - 'detail_13313': DataFrame chỉ TK13313
      - 'summary': Tổng PS Nợ/Có theo TK
      - 'du_dau_ky': dict dư đầu kỳ theo TK
    """
    df, du_dau_ky = load_ta030(config, "ta030_1331")

    if df.empty:
        return {
            "detail": df,
            "detail_13311": pd.DataFrame(),
            "detail_13313": pd.DataFrame(),
            "summary": pd.DataFrame(),
            "du_dau_ky": du_dau_ky
        }

    # Tách theo TK
    df_13311 = df[df["TaiKhoan"] == "TK13311"].copy()
    df_13313 = df[df["TaiKhoan"] == "TK13313"].copy()

    # Tính tổng phát sinh cho mỗi TK
    summary_records = []
    for tk, sub_df in [("TK13311", df_13311), ("TK13313", df_13313)]:
        data_rows = sub_df[sub_df["LoaiDong"] == "DATA"]
        tong_rows = sub_df[sub_df["LoaiDong"] == "TONG"]

        ps_no = data_rows["PSNo"].sum() if not data_rows.empty else 0
        ps_co = data_rows["PSCo"].sum() if not data_rows.empty else 0

        # Lấy từ dòng tổng nếu có
        if not tong_rows.empty:
            ps_no = tong_rows.iloc[0]["PSNo"]
            ps_co = tong_rows.iloc[0]["PSCo"]

        du_dk = du_dau_ky.get(tk, {})

        summary_records.append({
            "TaiKhoan": tk,
            "DuDauKy_No": du_dk.get("no", 0),
            "DuDauKy_Co": du_dk.get("co", 0),
            "PSNo": ps_no,
            "PSCo": ps_co,
        })

    summary = pd.DataFrame(summary_records)

    return {
        "detail": df,
        "detail_13311": df_13311,
        "detail_13313": df_13313,
        "summary": summary,
        "du_dau_ky": du_dau_ky
    }
