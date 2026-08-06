# -*- coding: utf-8 -*-
"""
Processor Tờ Khai Thuế 01/GTGT.
Tính toán các ô [21] → [43] của tờ khai thuế GTGT mẫu 01.
"""
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils import safe_float


def process_tokhai01(all_results: dict, taxvta: dict, config: dict) -> dict:
    """
    Tính toán tờ khai thuế 01/GTGT.

    Returns dict với key = tên ô (o21, o22, ..., o43)
    """
    # === DỮ LIỆU ĐẦU VÀO ===
    ta036 = all_results.get("ta036", {})
    ta035 = all_results.get("ta035", {})
    ta036_detail = ta036.get("detail", pd.DataFrame())
    ta035_detail = ta035.get("detail", pd.DataFrame())

    # Tổng mua vào
    tong_mv_chuathue = 0
    tong_mv_thue = 0
    if not ta036_detail.empty:
        tong_mv_chuathue = safe_float(ta036_detail["DoanhSoChuaThue"].sum())
        tong_mv_thue = safe_float(ta036_detail["ThueGTGT"].sum())

    # Tổng bán ra & phân loại theo thuế suất
    tong_br_chuathue = 0
    tong_br_thue = 0
    br_0_chuathue = 0    # Thuế suất 0%
    br_5_chuathue = 0    # Thuế suất 5%
    br_5_thue = 0
    br_10_chuathue = 0   # Thuế suất 10% (bao gồm cả 8% vì QĐ gốc = 10%)
    br_10_thue = 0
    br_ktt_chuathue = 0  # Không tính thuế (NaN)
    br_kct_chuathue = 0  # Không chịu thuế (0% nhưng là hàng XK hoặc không chịu thuế)

    if not ta035_detail.empty:
        tong_br_chuathue = safe_float(ta035_detail["DoanhSoChuaThue"].sum())
        tong_br_thue = safe_float(ta035_detail["ThueGTGT"].sum())

        for _, row in ta035_detail.iterrows():
            ds = safe_float(row.get("DoanhSoChuaThue", 0))
            thue = safe_float(row.get("ThueGTGT", 0))
            ts = row.get("ThueSuat", None)

            if pd.isna(ts) or ts is None:
                br_ktt_chuathue += ds
            else:
                ts_val = safe_float(ts)
                if ts_val == 0:
                    br_0_chuathue += ds
                elif ts_val == 5:
                    br_5_chuathue += ds
                    br_5_thue += thue
                elif ts_val in [8, 10]:
                    br_10_chuathue += ds
                    br_10_thue += thue

    # === TÍNH CÁC Ô TỜ KHAI ===
    # Lấy số dư thuế GTGT đầu kỳ từ kiểm dò
    kiem_do = taxvta.get("kiem_do_tk", {})
    tk1331 = kiem_do.get("TK13311", {})

    # [22] Thuế GTGT còn được khấu trừ kỳ trước chuyển sang
    # = Số dư Nợ TK13311 đầu kỳ - SPC (phần được khấu trừ thực tế)
    du_dau_ky_1331 = safe_float(tk1331.get("du_dau_ky", 0))
    du_spc_1331 = safe_float(config.get("tk13311_du_spc", 60549888887))
    o22 = max(du_dau_ky_1331 - du_spc_1331, 0)

    # [21] Không phát sinh hoạt động mua bán
    o21 = ""  # Checkbox - để trống

    # [23] Giá trị HHDV mua vào
    o23 = tong_mv_chuathue
    # [24] Thuế GTGT HHDV mua vào
    o24 = tong_mv_thue
    # [23a] Hàng nhập khẩu
    o23a = 0
    # [24a] Thuế nhập khẩu
    o24a = 0
    # [25] Thuế GTGT mua vào được khấu trừ kỳ này
    o25 = o24

    # [26] HHDV bán ra không chịu thuế
    o26 = 0
    # [27] HHDV bán ra chịu thuế GTGT
    # = [29]+[30]+[32]+[32a]
    o29 = br_0_chuathue        # Thuế suất 0%
    o30 = br_5_chuathue        # Thuế suất 5%
    o31 = br_5_thue            # Thuế 5%
    o32 = br_10_chuathue       # Thuế suất 10% (gồm 8%)
    o32a = br_ktt_chuathue     # Không tính thuế
    o33 = br_10_thue           # Thuế GTGT thuế suất 10%

    o27 = tong_br_chuathue
    o28 = tong_br_thue

    # [34] Tổng DT hàng hóa bán ra
    o34 = o26 + o27
    # [35] Tổng thuế GTGT bán ra
    o35 = o28

    # [36] Thuế GTGT phát sinh trong kỳ = [35] - [25]
    o36 = o35 - o25

    # [37] Điều chỉnh giảm
    o37 = 0
    # [38] Điều chỉnh tăng
    o38 = 0
    # [39a] Thuế GTGT nhận bàn giao
    o39a = 0

    # [40a] Thuế GTGT phải nộp từ HĐSXKD
    # = [36] - [22] + [37] - [38] - [39a] nếu > 0
    o40a_raw = o36 - o22 + o37 - o38 - o39a
    o40a = max(o40a_raw, 0)

    # [40b] Thuế GTGT mua vào dự án đầu tư
    o40b = 0

    # [40] Thuế còn phải nộp = [40a] - [40b]
    o40 = max(o40a - o40b, 0)

    # [41] Thuế GTGT chưa khấu trừ hết
    # = abs([40a_raw]) nếu [40a_raw] < 0
    o41 = abs(o40a_raw) if o40a_raw < 0 else 0

    # [42] Thuế GTGT đề nghị hoàn
    o42 = 0

    # [43] Thuế GTGT chuyển kỳ sau = [41] - [42]
    o43 = o41 - o42

    # Thông tin kỳ kê khai
    report_month = config.get("report_month", "")

    return {
        "report_month": report_month,
        "o21": o21, "o22": o22,
        "o23": o23, "o24": o24, "o23a": o23a, "o24a": o24a,
        "o25": o25,
        "o26": o26, "o27": o27, "o28": o28,
        "o29": o29, "o30": o30, "o31": o31,
        "o32": o32, "o32a": o32a, "o33": o33,
        "o34": o34, "o35": o35,
        "o36": o36,
        "o37": o37, "o38": o38, "o39a": o39a,
        "o40a": o40a, "o40b": o40b, "o40": o40,
        "o41": o41, "o42": o42, "o43": o43,
    }
