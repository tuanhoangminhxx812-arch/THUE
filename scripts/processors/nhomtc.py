# -*- coding: utf-8 -*-
"""
Xử lý BC Sản lượng Nhôm Toàn Cầu.
- Phân TRONGTHANG / CUOITHANG dựa trên ngày phát hành vs tháng báo cáo
- Group by THOIGIAN → SUM(TongTien)
- Liên kết qua ký hiệu DIEN00
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from data_loader import load_nhomtc
from utils import get_report_month_date, get_report_month_end, parse_date_flexible


def process_nhomtc(config: dict) -> dict:
    """
    Xử lý file Nhôm Toàn Cầu.
    Returns dict:
      - 'detail': DataFrame chi tiết (có THOIGIAN)
      - 'summary': DataFrame TRONGTHANG/CUOITHANG → SUM(TongTien)
      - 'total': Tổng tiền toàn bộ (= DIEN00)
    """
    df = load_nhomtc(config)

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame(), "total": 0}

    report_end = get_report_month_end(config)

    # Phân loại TRONGTHANG/CUOITHANG
    def classify_period(ngay_ph):
        d = parse_date_flexible(ngay_ph)
        if d is None:
            return "TRONGTHANG"
        if d > report_end:
            return "CUOITHANG"
        return "TRONGTHANG"

    df["THOIGIAN"] = df["NgayPhatHanh"].apply(classify_period)
    df["LIENKET"] = "DIEN00"  # Liên kết với các file khác

    # Group by THOIGIAN
    summary = df.groupby("THOIGIAN").agg(
        TongTien=("TongTien", "sum"),
        TienPS=("TienPS", "sum"),
        ThuePS=("ThuePS", "sum"),
        SoLuong=("TongTien", "count")
    ).reset_index()

    # Sắp xếp
    order = {"TRONGTHANG": 0, "CUOITHANG": 1}
    summary["_order"] = summary["THOIGIAN"].map(order).fillna(2)
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    # Dòng tổng
    total_amount = summary["TongTien"].sum()
    total = pd.DataFrame([{
        "THOIGIAN": "TỔNG CỘNG",
        "TongTien": total_amount,
        "TienPS": summary["TienPS"].sum(),
        "ThuePS": summary["ThuePS"].sum(),
        "SoLuong": summary["SoLuong"].sum()
    }])
    summary = pd.concat([summary, total], ignore_index=True)

    return {"detail": df, "summary": summary, "total": total_amount}
