# -*- coding: utf-8 -*-
"""
Xử lý TA_035_TK3331 (Bảng kê HĐ bán ra).
- Thêm cột PHANLOAI dựa trên diễn giải/mặt hàng/ghi chú
- Thêm cột TAIKHOAN (tất cả = TK33311)
- Group by PHANLOAI → SUM(DoanhSoChuaThue, ThueGTGT)
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from classifier import Classifier
from data_loader import load_ta035


def process_ta035(config: dict) -> dict:
    """
    Xử lý file TA_035.
    Returns:
        dict với keys:
        - 'detail': DataFrame chi tiết (có thêm cột PHANLOAI, TAIKHOAN)
        - 'summary': DataFrame group by PHANLOAI
    """
    clf = Classifier(config)
    df = load_ta035(config)

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame()}

    # Thêm cột PHANLOAI: ưu tiên GhiChu, rồi MatHang, rồi diễn giải chung
    def classify_row(row):
        # Thử GhiChu trước
        text = str(row.get("GhiChu", "")) + " " + str(row.get("MatHang", ""))
        # Cũng ghép thêm LoaiHinhKD
        text += " " + str(row.get("LoaiHinhKD", ""))
        return clf.classify(text)

    df["PHANLOAI"] = df.apply(classify_row, axis=1)

    # Thêm cột TAIKHOAN - tất cả = TK33311 (theo yêu cầu anh)
    df["TAIKHOAN"] = "TK33311"

    # Group by PHANLOAI
    summary = df.groupby("PHANLOAI").agg(
        DoanhSoChuaThue=("DoanhSoChuaThue", "sum"),
        ThueGTGT=("ThueGTGT", "sum"),
        SoLuong=("STT", "count")
    ).reset_index()

    # Sắp xếp theo thứ tự chuẩn
    code_order = clf.get_all_codes()
    summary["_order"] = summary["PHANLOAI"].apply(
        lambda x: code_order.index(x) if x in code_order else len(code_order)
    )
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    # Thêm dòng tổng
    total = pd.DataFrame([{
        "PHANLOAI": "Cộng",
        "DoanhSoChuaThue": summary["DoanhSoChuaThue"].sum(),
        "ThueGTGT": summary["ThueGTGT"].sum(),
        "SoLuong": summary["SoLuong"].sum()
    }])
    summary = pd.concat([summary, total], ignore_index=True)

    return {"detail": df, "summary": summary}
