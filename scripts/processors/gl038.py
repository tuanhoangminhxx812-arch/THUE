# -*- coding: utf-8 -*-
"""
Xử lý GL_038_TK33895 (Sổ chi tiết TK33895 - thuế chưa kê khai).
- Thêm cột PHANLOAI dựa trên diễn giải
- Group by PHANLOAI → SUM(PSCo)
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from classifier import Classifier
from data_loader import load_gl038


def process_gl038(config: dict) -> dict:
    """
    Xử lý file GL_038_TK33895.
    Returns dict:
      - 'detail': DataFrame chi tiết
      - 'summary': Group by PHANLOAI → SUM(PSCo, PSNo)
      - 'du_dau_ky': Số dư đầu kỳ
    """
    clf = Classifier(config)
    df, du_dau_ky = load_gl038(config)

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame(), "du_dau_ky": du_dau_ky}

    # Thêm cột PHANLOAI
    df["PHANLOAI"] = df["DienGiai"].apply(clf.classify)
    df["TAIKHOAN"] = "TK33895"

    # Group by PHANLOAI
    summary = df.groupby("PHANLOAI").agg(
        PSNo=("PSNo", "sum"),
        PSCo=("PSCo", "sum"),
        SoLuong=("DienGiai", "count")
    ).reset_index()

    # Sắp xếp
    code_order = clf.get_all_codes()
    summary["_order"] = summary["PHANLOAI"].apply(
        lambda x: code_order.index(x) if x in code_order else len(code_order)
    )
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    # Tính dư cuối kỳ
    du_cuoi_ky = du_dau_ky + df["PSCo"].sum() - df["PSNo"].sum()

    return {
        "detail": df,
        "summary": summary,
        "du_dau_ky": du_dau_ky,
        "du_cuoi_ky": du_cuoi_ky,
        "tong_ps_no": df["PSNo"].sum(),
        "tong_ps_co": df["PSCo"].sum()
    }
