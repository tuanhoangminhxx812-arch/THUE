# -*- coding: utf-8 -*-
"""
Xử lý GL_0903_TK511 (Báo cáo giao dịch TK511 - doanh thu).
- Thêm cột PHANLOAI dựa trên Nội dung
- Group by PHANLOAI → SUM(Có nguyên tệ)
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from classifier import Classifier
from data_loader import load_gl0903


def process_gl0903(config: dict) -> dict:
    """
    Xử lý file GL_0903_TK511.
    Returns dict:
      - 'detail': DataFrame chi tiết (có PHANLOAI)
      - 'summary': Group by PHANLOAI → SUM(CoNT)
    """
    clf = Classifier(config)
    df = load_gl0903(config)

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame()}

    # Thêm cột PHANLOAI dựa trên cột NoiDung
    df["PHANLOAI"] = df["NoiDung"].apply(clf.classify)

    # Group by PHANLOAI → SUM Có nguyên tệ
    summary = df.groupby("PHANLOAI").agg(
        CoNT=("CoNT", "sum"),
        SoLuong=("NoiDung", "count")
    ).reset_index()

    # Sắp xếp theo thứ tự chuẩn
    code_order = clf.get_all_codes()
    summary["_order"] = summary["PHANLOAI"].apply(
        lambda x: code_order.index(x) if x in code_order else len(code_order)
    )
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    # Thêm dòng tổng
    total = pd.DataFrame([{
        "PHANLOAI": "Tổng cộng",
        "CoNT": summary["CoNT"].sum(),
        "SoLuong": summary["SoLuong"].sum()
    }])
    summary = pd.concat([summary, total], ignore_index=True)

    return {"detail": df, "summary": summary}
