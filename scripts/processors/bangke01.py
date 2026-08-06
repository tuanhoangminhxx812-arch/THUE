# -*- coding: utf-8 -*-
"""
Processor Bảng Kê 01 - Giảm thuế GTGT theo NQ142/TT80.
Lọc hóa đơn thuế suất 8% từ TA036 (mua vào) và TA035 (bán ra).
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils import safe_float


def process_bangke01(all_results: dict, config: dict) -> dict:
    """
    Tạo Bảng Kê 01 từ dữ liệu TA036 và TA035.
    
    Returns dict:
      - 'muavao_detail': DataFrame chi tiết mua vào 8%
      - 'muavao_summary': DataFrame tổng hợp mua vào theo NCC
      - 'banra_detail': DataFrame chi tiết bán ra 8%
      - 'banra_summary': DataFrame tổng hợp bán ra theo loại HHDV
      - 'tong_muavao_chuathue': Tổng doanh số chưa thuế mua vào
      - 'tong_muavao_thue': Tổng thuế GTGT mua vào
      - 'tong_banra_chuathue': Tổng doanh số chưa thuế bán ra
      - 'tong_banra_thue_duocgiam': Tổng thuế được giảm bán ra
    """

    # === PHẦN I: MUA VÀO (TA036) - Thuế suất 8% ===
    ta036 = all_results.get("ta036", {})
    ta036_detail = ta036.get("detail", pd.DataFrame())

    muavao_8 = pd.DataFrame()
    muavao_summary = pd.DataFrame()
    tong_mv_chuathue = 0
    tong_mv_thue = 0

    if not ta036_detail.empty and "ThueSuat" in ta036_detail.columns:
        # Lọc thuế suất 8
        mask = ta036_detail["ThueSuat"].apply(
            lambda x: safe_float(x) == 8
        )
        muavao_8 = ta036_detail[mask].copy()

        if not muavao_8.empty:
            # Group theo tên NCC
            muavao_summary = muavao_8.groupby("TenNguoiBan", as_index=False).agg({
                "DoanhSoChuaThue": "sum",
                "ThueGTGT": "sum"
            }).sort_values("TenNguoiBan")

            # Thêm STT
            muavao_summary.insert(0, "STT", range(1, len(muavao_summary) + 1))
            muavao_summary.columns = ["STT", "Tên hàng hóa, dịch vụ",
                                       "Giá trị HHDV mua vào chưa có thuế GTGT",
                                       "Thuế GTGT của HHDV mua vào"]

            tong_mv_chuathue = muavao_summary["Giá trị HHDV mua vào chưa có thuế GTGT"].sum()
            tong_mv_thue = muavao_summary["Thuế GTGT của HHDV mua vào"].sum()

    # === PHẦN II: BÁN RA (TA035) - Thuế suất 8% ===
    ta035 = all_results.get("ta035", {})
    ta035_detail = ta035.get("detail", pd.DataFrame())

    banra_8 = pd.DataFrame()
    banra_summary = pd.DataFrame()
    tong_br_chuathue = 0
    tong_br_thue_duocgiam = 0

    if not ta035_detail.empty and "ThueSuat" in ta035_detail.columns:
        mask = ta035_detail["ThueSuat"].apply(
            lambda x: safe_float(x) == 8
        )
        banra_8 = ta035_detail[mask].copy()

        if not banra_8.empty:
            # Theo mẫu: DIEN01+CSPK02 gộp thành "Hoạt động kinh doanh điện"
            # Các mã khác (1388ĐĐ, 4500TT, ...) liệt kê chi tiết từng dòng

            # Nhóm mã gộp vào "Hoạt động kinh doanh điện"
            ma_gop = {"DIEN01", "CSPK02"}

            # Xác định mã hàng từ cột MatHang
            dien_mask = banra_8["MatHang"].apply(
                lambda x: str(x).strip().upper() in ma_gop if pd.notna(x) else False
            )

            rows = []

            # Dòng 1: Gộp hoạt động kinh doanh điện
            dien_rows = banra_8[dien_mask]
            if not dien_rows.empty:
                tong_dien = safe_float(dien_rows["DoanhSoChuaThue"].sum())
                rows.append({
                    "Tên hàng hóa, dịch vụ": "Hoạt động kinh doanh điện",
                    "Giá trị HHDV chưa có thuế GTGT": tong_dien,
                    "Thuế suất theo QĐ": 10,
                    "Thuế suất sau giảm": 8,
                    "Thuế GTGT được giảm": int(tong_dien * 0.02),
                })

            # Các dòng còn lại: liệt kê chi tiết
            other_rows = banra_8[~dien_mask]
            for _, row in other_rows.iterrows():
                ds = safe_float(row.get("DoanhSoChuaThue", 0))
                ten = str(row.get("MatHang", "")).strip()
                rows.append({
                    "Tên hàng hóa, dịch vụ": ten,
                    "Giá trị HHDV chưa có thuế GTGT": ds,
                    "Thuế suất theo QĐ": 10,
                    "Thuế suất sau giảm": 8,
                    "Thuế GTGT được giảm": int(ds * 0.02),
                })

            banra_summary = pd.DataFrame(rows)
            banra_summary.insert(0, "STT", range(1, len(banra_summary) + 1))

            tong_br_chuathue = banra_summary["Giá trị HHDV chưa có thuế GTGT"].sum()
            tong_br_thue_duocgiam = banra_summary["Thuế GTGT được giảm"].sum()

    return {
        "muavao_detail": muavao_8,
        "muavao_summary": muavao_summary,
        "banra_detail": banra_8,
        "banra_summary": banra_summary,
        "tong_muavao_chuathue": tong_mv_chuathue,
        "tong_muavao_thue": tong_mv_thue,
        "tong_banra_chuathue": tong_br_chuathue,
        "tong_banra_thue_duocgiam": tong_br_thue_duocgiam,
    }

