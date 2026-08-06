# -*- coding: utf-8 -*-
"""
Xử lý BC_HDon_01_THopTheoNgayGCS (Báo cáo tổng hợp theo ngày GCS).
- Phân TRONGTHANG / CUOITHANG dựa trên ngày phát hành HĐ vs tháng báo cáo
- Group by THOIGIAN → SUM(TienDien, ThueDien, TongDien, TienCSPK, ThueCSPK, TongCSPK)
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from data_loader import load_bchdon
from utils import get_report_month_date, get_report_month_end, parse_date_flexible


def process_bchdon(config: dict) -> dict:
    """
    Xử lý file BC_HDon_01.
    Returns dict:
      - 'detail': DataFrame chi tiết (có THOIGIAN)
      - 'summary': DataFrame tổng hợp TRONGTHANG/CUOITHANG
    """
    df = load_bchdon(config)

    if df.empty:
        return {"detail": df, "summary": pd.DataFrame()}

    # Lấy tháng báo cáo
    report_end = get_report_month_end(config)

    # Phân loại TRONGTHANG/CUOITHANG dựa trên ngày phát hành HĐ
    def classify_period(ngay_ph):
        d = parse_date_flexible(ngay_ph)
        if d is None:
            return "TRONGTHANG"  # Mặc định
        if d > report_end:
            return "CUOITHANG"
        return "TRONGTHANG"

    df["THOIGIAN"] = df["NgayPhatHanhHD"].apply(classify_period)

    # Group by THOIGIAN
    summary = df.groupby("THOIGIAN").agg(
        TienDien=("TienDien", "sum"),
        ThueDien=("ThueDien", "sum"),
        TongTienDien=("TongTienDien", "sum"),
        TienCSPK=("TienCSPK", "sum"),
        ThueCSPK=("ThueCSPK", "sum"),
        TongCSPK=("TongCSPK", "sum"),
        TongCong=("TongCong", "sum")
    ).reset_index()

    # Đảm bảo thứ tự TRONGTHANG trước CUOITHANG
    order = {"TRONGTHANG": 0, "CUOITHANG": 1}
    summary["_order"] = summary["THOIGIAN"].map(order).fillna(2)
    summary = summary.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    # Thêm dòng tổng cộng
    total = pd.DataFrame([{
        "THOIGIAN": "TỔNG CỘNG",
        "TienDien": summary["TienDien"].sum(),
        "ThueDien": summary["ThueDien"].sum(),
        "TongTienDien": summary["TongTienDien"].sum(),
        "TienCSPK": summary["TienCSPK"].sum(),
        "ThueCSPK": summary["ThueCSPK"].sum(),
        "TongCSPK": summary["TongCSPK"].sum(),
        "TongCong": summary["TongCong"].sum()
    }])
    summary = pd.concat([summary, total], ignore_index=True)

    return {"detail": df, "summary": summary}
