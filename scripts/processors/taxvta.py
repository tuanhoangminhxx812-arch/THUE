# -*- coding: utf-8 -*-
"""
Tổng hợp TAXVTA - sheet chính kết hợp dữ liệu từ tất cả processor.
Tính toán cross-check và kiểm tra chênh lệch.
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils import safe_float


def process_taxvta(all_results: dict, config: dict) -> dict:
    """
    Tổng hợp dữ liệu từ tất cả processor thành báo cáo chính.

    Args:
        all_results: dict chứa kết quả từ tất cả processor
        config: cấu hình

    Returns dict:
      - 'thue_dau_vao': Tổng hợp thuế đầu vào
      - 'thue_dau_ra': Tổng hợp thuế đầu ra
      - 'cross_checks': Các kiểm tra chéo
      - 'kiem_do_tk': Kiểm dò theo tài khoản
    """
    # === THUẾ ĐẦU VÀO ===
    ta036 = all_results.get("ta036", {})
    ta036_summary = ta036.get("summary", pd.DataFrame())
    ta036_ts = ta036.get("summary_thue_suat", pd.DataFrame())

    ta030_1331 = all_results.get("ta030_1331", {})
    ta030_1331_summary = ta030_1331.get("summary", pd.DataFrame())

    # Thuế đầu vào theo TK
    thue_dau_vao = {
        "TA036": {},
        "TA030": {},
        "chenh_lech": {}
    }

    for _, row in ta036_summary.iterrows():
        tk = row["TAIKHOAN"]
        if tk != "Tổng cộng":
            thue_dau_vao["TA036"][tk] = {
                "doanh_so": row["DoanhSoChuaThue"],
                "thue": row["ThueGTGT"]
            }

    for _, row in ta030_1331_summary.iterrows():
        tk = row["TaiKhoan"]
        thue_dau_vao["TA030"][tk] = {
            "ps_no": row["PSNo"],
            "ps_co": row["PSCo"]
        }

    # Tính chênh lệch TA036 vs TA030
    for tk in ["TK13311", "TK13313"]:
        ta036_thue = thue_dau_vao["TA036"].get(tk, {}).get("thue", 0)
        ta030_psno = thue_dau_vao["TA030"].get(tk, {}).get("ps_no", 0)
        # Thuế đầu vào: TA036.ThueGTGT should = TA030.PSNo (phát sinh nợ)
        # Nhưng cần tính tổng PS Nợ từ data (không phải dòng tổng, vì có kết chuyển)
        thue_dau_vao["chenh_lech"][tk] = ta036_thue - ta030_psno

    # Thuế đầu vào theo thuế suất (cho tờ khai)
    thue_dv_theo_ts = []
    if not ta036_ts.empty:
        for _, row in ta036_ts.iterrows():
            ts = row["ThueSuat"]
            if str(ts) != "Tổng":
                thue_dv_theo_ts.append({
                    "thue_suat": safe_float(ts),
                    "doanh_so": row["DoanhSoChuaThue"],
                    "thue": row["ThueGTGT"]
                })

    # === THUẾ ĐẦU RA ===
    ta035 = all_results.get("ta035", {})
    ta035_summary = ta035.get("summary", pd.DataFrame())

    gl0903 = all_results.get("gl0903", {})
    gl0903_summary = gl0903.get("summary", pd.DataFrame())

    bchdon = all_results.get("bchdon", {})
    bchdon_summary = bchdon.get("summary", pd.DataFrame())

    kd4a = all_results.get("kd4a", {})

    # Thuế đầu ra
    thue_dau_ra = {
        "TA035": {},
        "GL0903": {},
        "KD4A": kd4a,
        "BCHDON": {}
    }

    # TA035 theo phân loại
    if not ta035_summary.empty:
        for _, row in ta035_summary.iterrows():
            pl = row["PHANLOAI"]
            if pl != "Cộng":
                thue_dau_ra["TA035"][pl] = {
                    "doanh_so": row["DoanhSoChuaThue"],
                    "thue": row["ThueGTGT"]
                }

    # GL0903 theo phân loại
    if not gl0903_summary.empty:
        for _, row in gl0903_summary.iterrows():
            pl = row["PHANLOAI"]
            if pl != "Tổng cộng":
                thue_dau_ra["GL0903"][pl] = {
                    "co_nt": row["CoNT"]
                }

    # BC_HDon theo TRONG/CUOITHANG
    if not bchdon_summary.empty:
        for _, row in bchdon_summary.iterrows():
            tg = row["THOIGIAN"]
            thue_dau_ra["BCHDON"][tg] = {
                "tien_dien": row["TienDien"],
                "thue_dien": row["ThueDien"],
                "tong_dien": row["TongTienDien"],
                "tien_cspk": row["TienCSPK"],
                "thue_cspk": row["ThueCSPK"],
                "tong_cspk": row["TongCSPK"]
            }

    # === CROSS-CHECK ===
    cross_checks = []

    # 1. KD4A vs GL0903 (chỉ so sánh DIEN01 - Điện có thuế)
    # KD4A cột "DT ĐTDL" đã bao gồm cả DIEN00, nên phải trừ DIEN00 ra
    kd4a_dt_dtdl = safe_float(kd4a.get("DoanhThu_DTDL", 0))

    # Lấy DIEN00 từ GL0903 để trừ
    gl0903_dien00 = 0
    gl0903_dien01 = 0
    if not gl0903_summary.empty:
        dien00_row = gl0903_summary[gl0903_summary["PHANLOAI"] == "DIEN00"]
        if not dien00_row.empty:
            gl0903_dien00 = safe_float(dien00_row.iloc[0]["CoNT"])
        dien01_row = gl0903_summary[gl0903_summary["PHANLOAI"] == "DIEN01"]
        if not dien01_row.empty:
            gl0903_dien01 = safe_float(dien01_row.iloc[0]["CoNT"])

    kd4a_dien01 = kd4a_dt_dtdl - gl0903_dien00  # Trừ DIEN00 để còn DIEN01

    cross_checks.append({
        "name": "KD4A vs GL0903 (DIEN01)",
        "source_a": "KD4A (DT ĐTDL - DIEN00)",
        "value_a": kd4a_dien01,
        "source_b": "GL0903 DIEN01 (Có NT)",
        "value_b": gl0903_dien01,
        "chenh_lech": kd4a_dien01 - gl0903_dien01,
        "status": "✓ OK" if abs(kd4a_dien01 - gl0903_dien01) < 1 else "⚠ LỆCH"
    })

    # 2. TA036 vs TA030_1331 (chi tiết từng TK)
    tong_ta036 = 0
    tong_ta030 = 0
    for tk in ["TK13311", "TK13313"]:
        cl = thue_dau_vao["chenh_lech"].get(tk, 0)
        val_a = thue_dau_vao["TA036"].get(tk, {}).get("thue", 0)
        val_b = thue_dau_vao["TA030"].get(tk, {}).get("ps_no", 0)
        tong_ta036 += val_a
        tong_ta030 += val_b
        cross_checks.append({
            "name": f"TA036 vs TA030 ({tk})",
            "source_a": f"TA036 ({tk})",
            "value_a": val_a,
            "source_b": f"TA030 PS Nợ ({tk})",
            "value_b": val_b,
            "chenh_lech": cl,
            "status": "✓ OK" if abs(cl) < 1 else "⚠ LỆCH"
        })

    # 3. Tổng TK1331 (TK13311 + TK13313) - kiểm tra bù trừ điều chỉnh
    cl_tong = tong_ta036 - tong_ta030
    cross_checks.append({
        "name": "TỔNG TK1331 (13311+13313)",
        "source_a": "TA036 Tổng TK1331",
        "value_a": tong_ta036,
        "source_b": "TA030 Tổng PS Nợ",
        "value_b": tong_ta030,
        "chenh_lech": cl_tong,
        "status": "✓ OK (bù trừ ĐC)" if abs(cl_tong) < 1 else "⚠ LỆCH"
    })

    # === KIỂM DÒ THEO TÀI KHOẢN (Layout chữ T) ===
    ta030_3331 = all_results.get("ta030_3331", {})
    gl038 = all_results.get("gl038", {})

    # Khởi tạo cấu trúc T-account
    kiem_do = {
        "TK333111": {
            "du_dau_ky_hcm": 0,      # Số dư HCM đầu kỳ (bên Nợ)
            "ps_co": 0,               # PS tháng này (bên Có)
            "kc_13311": 0,            # KC 13311→333111 SXKD (bên Nợ)
            "kc_13313": 0,            # KC 13313→333111 XDCB (bên Nợ)
            "du_cuoi_ky_hcm": 0,      # Số dư HCM cuối kỳ = PS Có - KC
            "du_dau_ky_spc": 1500000000,  # Số dư SPC (giữ nguyên)
        },
        "TK13311": {
            "du_dau_ky": 0,           # Số dư SPC đầu kỳ (bên Nợ)
            "ps_no": 0,               # PS tháng này (bên Nợ)
            "kc_333": 0,              # KC 133→333 (bên Có)
            "du_cuoi_ky_spc": 0,      # Số dư SPC cuối kỳ (giữ nguyên)
        }
    }

    # === TK333111: Lấy dữ liệu ===
    # Dư đầu kỳ Nợ từ TA030_TK3331
    du_dk_3331 = ta030_3331.get("du_dau_ky", {}).get("TK33311", {})
    du_dk_tong = du_dk_3331.get("no", 0)
    du_spc_333 = kiem_do["TK333111"]["du_dau_ky_spc"]  # 1.500.000.000
    # HCM = Tổng dư đầu kỳ - SPC (SPC luôn = 1,5 tỷ → HCM = 0)
    kiem_do["TK333111"]["du_dau_ky_hcm"] = max(du_dk_tong - du_spc_333, 0)

    # PS Có TK333111 từ TA030_TK3331
    ta030_3331_detail = ta030_3331.get("detail", pd.DataFrame())
    if not ta030_3331_detail.empty:
        tong_333 = ta030_3331_detail[
            (ta030_3331_detail["TaiKhoan"] == "TK33311") &
            (ta030_3331_detail["LoaiDong"] == "TONG")
        ]
        if not tong_333.empty:
            kiem_do["TK333111"]["ps_co"] = safe_float(tong_333.iloc[0]["PSCo"])

    # KC 13311→333111 và KC 13313→333111 từ TA030_TK1331
    ta030_1331_detail = ta030_1331.get("detail", pd.DataFrame())
    if not ta030_1331_detail.empty:
        tong_13311 = ta030_1331_detail[
            (ta030_1331_detail["TaiKhoan"] == "TK13311") &
            (ta030_1331_detail["LoaiDong"] == "TONG")
        ]
        if not tong_13311.empty:
            kiem_do["TK333111"]["kc_13311"] = safe_float(tong_13311.iloc[0]["PSCo"])
            kiem_do["TK13311"]["ps_no"] = safe_float(tong_13311.iloc[0]["PSNo"])
            kiem_do["TK13311"]["kc_333"] = safe_float(tong_13311.iloc[0]["PSCo"])

        tong_13313 = ta030_1331_detail[
            (ta030_1331_detail["TaiKhoan"] == "TK13313") &
            (ta030_1331_detail["LoaiDong"] == "TONG")
        ]
        if not tong_13313.empty:
            kiem_do["TK333111"]["kc_13313"] = safe_float(tong_13313.iloc[0]["PSCo"])

    # Tính dư cuối kỳ TK333111 HCM = PS Có - KC 13311 - KC 13313 + Dư đầu kỳ HCM
    ps_co_333 = kiem_do["TK333111"]["ps_co"]
    kc_tong = kiem_do["TK333111"]["kc_13311"] + kiem_do["TK333111"]["kc_13313"]
    kiem_do["TK333111"]["du_cuoi_ky_hcm"] = ps_co_333 - kc_tong + kiem_do["TK333111"]["du_dau_ky_hcm"]

    # === TK13311: Lấy dữ liệu ===
    # Dư đầu kỳ Nợ từ TA030_TK1331
    du_dk_1331 = ta030_1331.get("du_dau_ky", {}).get("TK13311", {})
    kiem_do["TK13311"]["du_dau_ky"] = du_dk_1331.get("no", 0)

    # Số dư SPC cuối kỳ = 60.549.888.887 (cố định - thuế không được khấu trừ)
    kiem_do["TK13311"]["du_cuoi_ky_spc"] = config.get("tk13311_du_spc", 60549888887)

    return {
        "thue_dau_vao": thue_dau_vao,
        "thue_dau_ra": thue_dau_ra,
        "thue_dv_theo_ts": thue_dv_theo_ts,
        "cross_checks": cross_checks,
        "kiem_do_tk": kiem_do
    }

