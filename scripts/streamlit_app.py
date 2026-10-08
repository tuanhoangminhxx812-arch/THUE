# -*- coding: utf-8 -*-
"""
Streamlit Web Application - Đối Chiếu Thuế GTGT (Mua vào - Bán ra).
Công ty Điện lực Vũng Tàu.
"""
import sys
import os
import io
import traceback
from datetime import datetime

# Đảm bảo import đúng path
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import pandas as pd

from utils import (
    load_config, resolve_config_filenames, get_report_month_label,
    get_output_dir, format_number, safe_float, scan_available_months
)
from processors.ta035 import process_ta035
from processors.ta030 import process_ta030_3331, process_ta030_1331
from processors.ta036 import process_ta036
from processors.gl038 import process_gl038
from processors.gl0903 import process_gl0903
from processors.bchdon import process_bchdon
from processors.nhomtc import process_nhomtc
from processors.kd4a import process_kd4a
from processors.taxvta import process_taxvta
from processors.bangke01 import process_bangke01
from processors.tokhai01 import process_tokhai01
from report_builder import build_report
from st_theme import apply_theme
from tax_vta_manager import (
    ensure_tax_vta_file, update_tax_vta_file,
    get_tax_vta_summary_data, get_month_tax_vta_file
)

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Đối Chiếu Thuế GTGT - ĐLVT",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

apply_theme()


# ============================================================
# SESSION STATE INITIALIZATION
# ============================================================
if "results" not in st.session_state:
    st.session_state.results = None
if "taxvta" not in st.session_state:
    st.session_state.taxvta = None
if "config" not in st.session_state:
    st.session_state.config = None
if "processed_month" not in st.session_state:
    st.session_state.processed_month = None
if "bangke01" not in st.session_state:
    st.session_state.bangke01 = None
if "tokhai01" not in st.session_state:
    st.session_state.tokhai01 = None


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def fmt(num):
    """Format số tiền VND."""
    if num is None or pd.isna(num):
        return ""
    try:
        n = float(num)
        if n == 0:
            return "0"
        return f"{n:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return str(num)


def highlight_diff(val):
    """Highlight ô chênh lệch khác 0."""
    try:
        v = float(val)
        if abs(v) > 0.5:
            return "background-color: rgba(239, 83, 80, 0.2); color: #ef5350; font-weight: 700"
    except (ValueError, TypeError):
        pass
    return ""


def highlight_ok(val):
    """Highlight ô OK."""
    if isinstance(val, str) and ("OK" in val or "KHỚP" in val):
        return "background-color: rgba(102, 187, 106, 0.1); color: #66bb6a"
    if isinstance(val, str) and "LỆCH" in val:
        return "background-color: rgba(239, 83, 80, 0.2); color: #ef5350"
    return ""


def style_applymap(styler, func, subset=None):
    """Hỗ trợ cả Pandas cũ (.applymap) và Pandas mới 2.1+ (.map)."""
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)


def df_to_display(df, max_rows=500):
    """Prepare DataFrame for display."""
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        return pd.DataFrame()
    return df.head(max_rows)


def format_df(df):
    """Format toàn bộ cột số trong DataFrame thành chuỗi dấu chấm phân cách.
    Ví dụ: 1234567 → '1.234.567'
    """
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        return df
    df_out = df.copy()
    for col in df_out.columns:
        if pd.api.types.is_numeric_dtype(df_out[col]):
            df_out[col] = df_out[col].apply(
                lambda x: fmt(x) if pd.notna(x) else ""
            )
    return df_out


def run_all_processors(config):
    """Chạy tất cả processor và trả về kết quả."""
    results = {}
    progress = st.progress(0, text="Đang khởi tạo...")

    steps = [
        ("ta035", "TA_035 Bán ra", lambda: process_ta035(config)),
        ("ta030_3331", "TA_030 TK3331", lambda: process_ta030_3331(config)),
        ("ta030_1331", "TA_030 TK1331", lambda: process_ta030_1331(config)),
        ("ta036", "TA_036 Mua vào", lambda: process_ta036(config)),
        ("gl038", "GL_038 TK33895", lambda: process_gl038(config)),
        ("gl0903", "GL_0903 TK511", lambda: process_gl0903(config)),
        ("bchdon", "BC_HDon GCS", lambda: process_bchdon(config)),
        ("nhomtc", "Nhôm Toàn Cầu", lambda: process_nhomtc(config)),
        ("kd4a", "rptKDDN4A", lambda: process_kd4a(config)),
    ]

    for i, (key, label, func) in enumerate(steps):
        progress.progress((i + 1) / (len(steps) + 1), text=f"[{i+1}/{len(steps)}] Đang xử lý {label}...")
        results[key] = func()

    progress.progress(1.0, text="Đang tổng hợp báo cáo...")
    taxvta = process_taxvta(results, config)

    progress.empty()
    return results, taxvta


def detect_cash_risk(ta036_detail):
    """Phát hiện HĐ mua vào ≥ 20 triệu thanh toán tiền mặt."""
    if ta036_detail is None or ta036_detail.empty:
        return pd.DataFrame()

    df = ta036_detail.copy()
    risks = []

    for _, row in df.iterrows():
        tong_tt = safe_float(row.get("SoTienDaTT", 0))
        ghi_chu = str(row.get("GhiChu", "")).upper()
        mat_hang = str(row.get("MatHang", "")).upper()

        is_cash = any(kw in ghi_chu for kw in ["TM", "TIỀN MẶT", "TIEN MAT"]) or \
                  any(kw in mat_hang for kw in ["TM", "TIỀN MẶT"])

        if tong_tt >= 20_000_000 and is_cash:
            risks.append(row)

    if risks:
        return pd.DataFrame(risks)
    return pd.DataFrame()


def detect_duplicate_invoices(detail_df, ky_hieu_col="KyHieuHD", so_hd_col="SoHD"):
    """Phát hiện HĐ bị trùng số trong kỳ."""
    if detail_df is None or detail_df.empty:
        return pd.DataFrame()

    df = detail_df.copy()
    if ky_hieu_col not in df.columns or so_hd_col not in df.columns:
        return pd.DataFrame()

    df["_key"] = df[ky_hieu_col].astype(str).str.strip() + "|" + df[so_hd_col].astype(str).str.strip()
    dupes = df[df.duplicated(subset="_key", keep=False)].copy()
    dupes = dupes.drop(columns=["_key"], errors="ignore")
    return dupes


def drill_down_compare(df_a, df_b, key_col="SoHD", name_a="Hóa đơn", name_b="Sổ cái"):
    """So sánh 2 DataFrame dựa trên cột key để tìm records thiếu/thừa."""
    if df_a is None or df_a.empty:
        df_a = pd.DataFrame(columns=[key_col])
    if df_b is None or df_b.empty:
        df_b = pd.DataFrame(columns=[key_col])

    if key_col not in df_a.columns or key_col not in df_b.columns:
        return pd.DataFrame(), pd.DataFrame()

    keys_a = set(df_a[key_col].astype(str).str.strip().unique())
    keys_b = set(df_b[key_col].astype(str).str.strip().unique())

    only_in_a = keys_a - keys_b
    only_in_b = keys_b - keys_a

    result_a = df_a[df_a[key_col].astype(str).str.strip().isin(only_in_a)].copy()
    result_b = df_b[df_b[key_col].astype(str).str.strip().isin(only_in_b)].copy()

    return result_a, result_b


def get_month_from_quarter(quarter, year):
    """Trả về danh sách tháng thuộc quý."""
    q_map = {1: [1, 2, 3], 2: [4, 5, 6], 3: [7, 8, 9], 4: [10, 11, 12]}
    return [f"{year}-{m:02d}" for m in q_map.get(quarter, [])]


def build_t_account_html(kiem_do):
    """Render T-account (chữ T kế toán) bằng HTML."""
    tk333 = kiem_do.get("TK333111", {})
    tk133 = kiem_do.get("TK13311", {})

    def t_row(label, debit, credit, note="", is_total=False, color="blue"):
        cls = "t-account-row total" + (" right" if color == "green" else "") if is_total else "t-account-row"
        c = color

        # Hiển thị "-" nếu giá trị = 0 hoặc None
        if debit is not None:
            debit_val = safe_float(debit)
            debit_html = f'<div class="t-cell amount {c}">{fmt(debit_val) if abs(debit_val) > 0.5 else "-"}</div>'
        else:
            debit_html = '<div class="t-cell amount empty">—</div>'

        if credit is not None:
            credit_val = safe_float(credit)
            credit_html = f'<div class="t-cell amount {c}">{fmt(credit_val) if abs(credit_val) > 0.5 else "-"}</div>'
        else:
            credit_html = '<div class="t-cell amount empty">—</div>'

        note_html = f'<div class="note">{note}</div>' if note else ""
        return f'<div class="{cls}"><div class="t-cell label">{label}{note_html}</div>{debit_html}{credit_html}</div>'

    # === TK333111 (bên trái) ===
    left = '<div class="t-account">'
    left += '<div class="t-account-title left">TK 333111 - Thuế GTGT Phải Nộp</div>'
    left += '<div class="t-account-header left"><span></span><span>NỢ</span><span>CÓ</span></div>'
    left += t_row("Dư Nợ đầu kỳ HCM", tk333.get("du_dau_ky_hcm"), None, "Số nộp thừa kỳ trước chuyển sang")
    if safe_float(tk333.get("du_dau_ky_spc", 0)) > 0:
        left += t_row("Số dư SPC", tk333.get("du_dau_ky_spc"), None, "Số nộp dư của SPC Nợ TK333111")
    left += t_row("PS Có trong kỳ", None, tk333.get("ps_co"), "Thuế đầu ra phát sinh (TA030_TK3331)")
    left += t_row("KC 13311→333111 (SXKD)", tk333.get("kc_13311"), None, "Bù trừ thuế mua vào SXKD")
    left += t_row("KC 13313→333111 (XDCB)", tk333.get("kc_13313"), None, "Bù trừ thuế mua vào XDCB")

    tong_no = safe_float(tk333.get("du_dau_ky_hcm", 0)) + safe_float(tk333.get("kc_13311", 0)) + safe_float(tk333.get("kc_13313", 0))
    left += t_row("Tổng phát sinh & Dư ĐK", tong_no, tk333.get("ps_co"), is_total=True)
    
    du_co_cuoi = safe_float(tk333.get("ps_co", 0)) - tong_no
    if du_co_cuoi >= 0:
        left += t_row("Dư Có cuối kỳ (Phải nộp)", None, du_co_cuoi, "Nghĩa vụ thuế GTGT phải nộp NSNN", is_total=True, color="blue")
    else:
        left += t_row("Dư Nợ cuối kỳ (Nộp thừa)", abs(du_co_cuoi), None, "Số thuế còn được khấu trừ/cấn trừ kỳ sau", is_total=True, color="blue")
    left += '</div>'

    # === TK13311 (bên phải) ===
    right = '<div class="t-account">'
    right += '<div class="t-account-title right">TK 13311 - Thuế GTGT Được Khấu Trừ</div>'
    right += '<div class="t-account-header right"><span></span><span>NỢ</span><span>CÓ</span></div>'
    right += t_row("Dư Nợ SPC đầu kỳ", tk133.get("du_dau_ky"), None, "Số dư cố định 60.549.888.887 đ (giữ nguyên)", color="green")
    right += t_row("PS Nợ trong kỳ", tk133.get("ps_no"), None, "Thuế mua vào phát sinh kỳ này (TA036)", color="green")
    right += t_row("KC 133→333", None, tk133.get("kc_333"), "Kết chuyển sang TK 333111 bù trừ", color="green")

    tong_no_133 = safe_float(tk133.get("du_dau_ky", 0)) + safe_float(tk133.get("ps_no", 0))
    tong_co_133 = safe_float(tk133.get("kc_333", 0))
    right += t_row("Tổng Nợ / Có", tong_no_133, tong_co_133, is_total=True, color="green")

    du_cuoi = tong_no_133 - tong_co_133
    right += t_row("Dư Nợ cuối kỳ", du_cuoi, None, "Dư Nợ cuối kỳ tài khoản 13311", color="green")

    du_spc = safe_float(tk133.get("du_cuoi_ky_spc", 0))
    con_kt = du_cuoi - du_spc
    right += t_row("Dư Nợ SPC cố định", du_spc, None, "Yêu cầu bắt buộc giữ nguyên", color="green")
    right += t_row("Còn được KT thực tế", con_kt, None, "Số thuế còn được khấu trừ kỳ sau", is_total=True, color="green")
    right += '</div>'

    html = f'<div class="t-account-container">{left}<div class="t-divider">⟷</div>{right}</div>'
    return html


# ============================================================
# SIDEBAR
# ============================================================
def render_sidebar():
    """Render sidebar controls."""
    with st.sidebar:
        st.markdown('<div class="logo-container"><div class="logo-icon">⚡</div></div>', unsafe_allow_html=True)
        st.markdown("# THUẾ GTGT")
        st.caption("Công ty Điện lực Vũng Tàu")

        st.divider()

        # Kỳ tính thuế
        ky_thue = st.radio("📅 Kỳ tính thuế", ["Tháng", "Quý"], horizontal=True)

        # Quét tháng có dữ liệu
        months = scan_available_months()

        if ky_thue == "Tháng":
            if months:
                month_options = {m["label"]: m["month"] for m in months}
                selected_label = st.selectbox(
                    "Chọn tháng",
                    options=list(month_options.keys()),
                    index=0
                )
                selected_month = month_options[selected_label]
                file_count = next((m["file_count"] for m in months if m["month"] == selected_month), 0)
                st.caption(f"📁 {file_count} file trong thư mục")
            else:
                selected_month = "2026-05"
                st.warning("Chưa có thư mục tháng nào trong ĐẦU VÀO/")
        else:
            config = load_config()
            year = config.get("fiscal_year", 2026)
            quarter = st.selectbox("Chọn quý", [1, 2, 3, 4], format_func=lambda q: f"Quý {q}/{year}")
            quarter_months = get_month_from_quarter(quarter, year)
            available_in_quarter = [m for m in months if m["month"] in quarter_months]
            if available_in_quarter:
                selected_month = available_in_quarter[0]["month"]
                st.caption(f"📁 {len(available_in_quarter)} tháng có dữ liệu trong quý")
            else:
                selected_month = quarter_months[0] if quarter_months else "2026-01"
                st.warning(f"Chưa có dữ liệu cho Quý {quarter}")

        # Nút chạy & cập nhật + xuất Excel trực tiếp 1-click
        col1, col2 = st.columns(2)
        with col1:
            btn_process = st.button("🔄 Cập Nhật Dữ Liệu", type="primary", use_container_width=True)
        with col2:
            master_file = get_month_tax_vta_file(selected_month)
            if master_file.exists():
                with open(master_file, "rb") as f:
                    excel_bytes = f.read()
                st.download_button(
                    label="📥 Xuất Excel",
                    data=excel_bytes,
                    file_name=master_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"btn_export_dl_{selected_month}"
                )
            else:
                st.button("📥 Xuất Excel", disabled=True, use_container_width=True)

        # Trạng thái
        if st.session_state.processed_month:
            st.success(f"✓ Đã nạp: {st.session_state.processed_month}")

        return selected_month, btn_process


# ============================================================
# TAB 1: DASHBOARD KIỂM DÒ THUẾ 4 KHỐI
# ============================================================
def render_tab_tonghop(summary_data, taxvta=None):
    """
    Tab Dashboard Kiểm Dò Thuế GTGT - Rút gọn, tinh giản theo đúng yêu cầu:
    1. Đối chiếu chênh lệch giữa bảng TA036 với TA030 của 2 TK 13311 và 13313.
    2. Đối chiếu chênh lệch giữa 2 bảng TA035 với TA030 của các TK 333111, 333113, 333114.
    3. Đối chiếu chênh lệch giữa các bảng GCS, 4A và 0903.
    """
    st.markdown("### 📊 DASHBOARD KIỂM DÒ ĐỐI SOÁT THUẾ GTGT")
    st.caption("Đối soát 3 chiều giữa Hóa đơn (TA35/TA36) - Kinh doanh (GCS/4A) - Sổ cái tài chính (GL)")

    cards = summary_data.get("summary_cards", {})
    is_all_pass = cards.get("is_all_pass", False)
    spc_valid = cards.get("spc_valid", True)
    spc_target = cards.get("spc_target", 60549888887)

    # === KHỐI 1: 4 THẺ KPI TỔNG QUAN & CẢNH BÁO LỆCH ===
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(
            f"""
            <div style="background: rgba(46, 125, 50, 0.12); border: 1px solid #4caf50; border-radius: 8px; padding: 12px; text-align: center;">
                <div style="color: #81c784; font-size: 0.85rem; font-weight: 600; text-transform: uppercase;">🟢 Thuế Đầu Vào [25]</div>
                <div style="color: #4caf50; font-size: 1.45rem; font-weight: 700; margin-top: 4px;">{fmt(cards.get("thue_dau_vao"))}</div>
                <div style="color: #9e9e9e; font-size: 0.75rem; margin-top: 2px;">Khấu trừ kỳ này</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f"""
            <div style="background: rgba(239, 108, 0, 0.12); border: 1px solid #ff9800; border-radius: 8px; padding: 12px; text-align: center;">
                <div style="color: #ffb74d; font-size: 0.85rem; font-weight: 600; text-transform: uppercase;">🔴 Thuế Đầu Ra [28]</div>
                <div style="color: #ff9800; font-size: 1.45rem; font-weight: 700; margin-top: 4px;">{fmt(cards.get("thue_dau_ra"))}</div>
                <div style="color: #9e9e9e; font-size: 0.75rem; margin-top: 2px;">Phát sinh trong kỳ</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col3:
        st.markdown(
            f"""
            <div style="background: rgba(2, 136, 209, 0.12); border: 1px solid #03a9f4; border-radius: 8px; padding: 12px; text-align: center;">
                <div style="color: #4fc3f7; font-size: 0.85rem; font-weight: 600; text-transform: uppercase;">⚖️ Thuế Phải Nộp [36]</div>
                <div style="color: #03a9f4; font-size: 1.45rem; font-weight: 700; margin-top: 4px;">{fmt(cards.get("thue_phai_nop"))}</div>
                <div style="color: #9e9e9e; font-size: 0.75rem; margin-top: 2px;">Nghĩa vụ NSNN</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col4:
        status_bg = "rgba(46, 125, 50, 0.2)" if is_all_pass else "rgba(211, 47, 47, 0.2)"
        status_border = "#4caf50" if is_all_pass else "#f44336"
        status_color = "#66bb6a" if is_all_pass else "#ef5350"
        status_text = "KHỚP DỮ LIỆU (PASS)" if is_all_pass else "LỆCH SỐ LIỆU (CẦN KIỂM TRA)"
        status_sub = "Số liệu 3 chiều đồng bộ" if is_all_pass else "Phát hiện chênh lệch"

        st.markdown(
            f"""
            <div style="background: {status_bg}; border: 2px solid {status_border}; border-radius: 8px; padding: 12px; text-align: center;">
                <div style="color: {status_color}; font-size: 0.85rem; font-weight: 600; text-transform: uppercase;">🛡️ TRẠNG THÁI KIỂM DÒ</div>
                <div style="color: {status_color}; font-size: 1.15rem; font-weight: 800; margin-top: 6px;">{status_text}</div>
                <div style="color: #9e9e9e; font-size: 0.75rem; margin-top: 2px;">{status_sub}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Widget cảnh báo số dư Nợ SPC
    if spc_valid:
        st.markdown(
            f"""
            <div style="background: rgba(33, 150, 243, 0.08); border-left: 4px solid #2196f3; padding: 8px 14px; border-radius: 4px; margin-top: 12px; font-size: 0.9rem;">
                🔒 <strong>Số Dư Nợ SPC Cố Định:</strong> {fmt(spc_target)} đ 
                <span style="color: #4caf50; font-weight: 600; margin-left: 8px;">(✓ Khớp chuẩn - Không bị sai lệch)</span>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f"""
            <div style="background: rgba(244, 67, 54, 0.15); border-left: 4px solid #f44336; padding: 8px 14px; border-radius: 4px; margin-top: 12px; font-size: 0.9rem; color: #ff5252;">
                🚨 <strong>CẢNH BÁO: SAI LỆCH SỐ DƯ NỢ SPC!</strong> Số dư bên Nợ TK 13311 khác giá trị cố định <strong>{fmt(spc_target)} đ</strong>! Vui lòng kiểm tra lại dữ liệu đầu vào.
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # === BẢNG 1: ĐỐI SOÁT MUA VÀO (TA36 vs SỔ CÁI 13311) ===
    st.markdown("#### 1️⃣ Bảng 1: Đối Chiếu Chênh Lệch Giữa Bảng TA036 với TA030 (TK 13311 & TK 13313)")
    k2 = summary_data.get("khoi2_muavao", {})
    k2_rows = k2.get("rows", [])
    k2_tot = k2.get("total", {})

    if k2_rows:
        display_rows = []
        for r in k2_rows:
            display_rows.append({
                "STT": r["stt"],
                "Mã TK": r["ma_tk"],
                "Nội Dung Diễn Giải": r["noi_dung"],
                "Doanh Số TA36": fmt(r["doanh_so"]),
                "Tiền Thuế TA36": fmt(r["thue"]),
                "Tiền Thuế Sổ Cái": fmt(r["so_cai"]),
                "Chênh Lệch (TA36 - Sổ Cái)": fmt(r["chenh_lech"]),
                "Trạng Thái": r["status"]
            })
        if k2_tot:
            display_rows.append({
                "STT": "Cộng",
                "Mã TK": "",
                "Nội Dung Diễn Giải": k2_tot["noi_dung"],
                "Doanh Số TA36": fmt(k2_tot["doanh_so"]),
                "Tiền Thuế TA36": fmt(k2_tot["thue"]),
                "Tiền Thuế Sổ Cái": fmt(k2_tot["so_cai"]),
                "Chênh Lệch (TA36 - Sổ Cái)": fmt(k2_tot["chenh_lech"]),
                "Trạng Thái": k2_tot["status"]
            })

        df_k2 = pd.DataFrame(display_rows)
        styled_k2 = style_applymap(style_applymap(df_k2.style, highlight_diff, subset=["Chênh Lệch (TA36 - Sổ Cái)"]), highlight_ok, subset=["Trạng Thái"])
        st.dataframe(styled_k2, use_container_width=True, hide_index=True)

    st.divider()

    # === BẢNG 2: ĐỐI SOÁT BÁN RA (TA35 vs SỔ CÁI 3331) ===
    st.markdown("#### 2️⃣ Bảng 2: Đối Chiếu Chênh Lệch Giữa Bảng TA035 với TA030 (TK 333111, 333113, 333114)")
    b2 = summary_data.get("bang2_banra", {})
    b2_rows = b2.get("rows", [])
    b2_tot = b2.get("total", {})

    if b2_rows:
        display_b2 = []
        for r in b2_rows:
            display_b2.append({
                "STT": r["stt"],
                "Mã TK": r["ma_tk"],
                "Nội Dung Diễn Giải": r["noi_dung"],
                "Doanh Số TA35": fmt(r["doanh_so"]),
                "Tiền Thuế TA35": fmt(r["thue"]),
                "Tiền Thuế Sổ Cái": fmt(r["so_cai"]),
                "Chênh Lệch (TA35 - Sổ Cái)": fmt(r["chenh_lech"]),
                "Trạng Thái": r["status"]
            })
        if b2_tot:
            display_b2.append({
                "STT": "Cộng",
                "Mã TK": "",
                "Nội Dung Diễn Giải": b2_tot["noi_dung"],
                "Doanh Số TA35": fmt(b2_tot["doanh_so"]),
                "Tiền Thuế TA35": fmt(b2_tot["thue"]),
                "Tiền Thuế Sổ Cái": fmt(b2_tot["so_cai"]),
                "Chênh Lệch (TA35 - Sổ Cái)": fmt(b2_tot["chenh_lech"]),
                "Trạng Thái": b2_tot["status"]
            })

        df_b2 = pd.DataFrame(display_b2)
        styled_b2 = style_applymap(style_applymap(df_b2.style, highlight_diff, subset=["Chênh Lệch (TA35 - Sổ Cái)"]), highlight_ok, subset=["Trạng Thái"])
        st.dataframe(styled_b2, use_container_width=True, hide_index=True)

    st.divider()

    # === BẢNG 3: MA TRẬN ĐỐI SOÁT ĐẦU RA 3 CHIỀU (GCS vs 4A vs GL0903) ===
    st.markdown("#### 3️⃣ Bảng 3: Đối Chiếu Chênh Lệch Giữa Các Bảng GCS, 4A và 0903")
    k3 = summary_data.get("khoi3_matran3d", {})
    k3_rows = k3.get("rows", [])
    k3_tot = k3.get("total", {})

    if k3_rows:
        display_k3 = []
        for r in k3_rows:
            display_k3.append({
                "Hạng Mục Doanh Thu": r["hang_muc"],
                "Mã DFF": r["ma_dff"],
                "Doanh Số GCS": fmt(r["gcs"]),
                "Doanh Số 4A": fmt(r["kd4a"]),
                "Doanh Số GL 0903": fmt(r["gl0903"]),
                "Lệch (GCS - 4A)": fmt(r["lech_gcs_4a"]),
                "Lệch (4A - GL0903)": fmt(r["lech_4a_gl"]),
                "Đánh Giá": r["danh_gia"]
            })
        if k3_tot:
            display_k3.append({
                "Hạng Mục Doanh Thu": k3_tot["hang_muc"],
                "Mã DFF": "",
                "Doanh Số GCS": fmt(k3_tot["gcs"]),
                "Doanh Số 4A": fmt(k3_tot["kd4a"]),
                "Doanh Số GL 0903": fmt(k3_tot["gl0903"]),
                "Lệch (GCS - 4A)": fmt(k3_tot["lech_gcs_4a"]),
                "Lệch (4A - GL0903)": fmt(k3_tot["lech_4a_gl"]),
                "Đánh Giá": k3_tot["danh_gia"]
            })

        df_k3 = pd.DataFrame(display_k3)
        styled_k3 = style_applymap(style_applymap(df_k3.style, highlight_diff, subset=["Lệch (GCS - 4A)", "Lệch (4A - GL0903)"]), highlight_ok, subset=["Đánh Giá"])
        st.dataframe(styled_k3, use_container_width=True, hide_index=True)





# ============================================================
# TAB 2: BÁN RA (TA035)
# ============================================================
def render_tab_banra(results):
    """Tab bán ra - TA035."""
    st.markdown("### 📤 Bảng Kê Bán Ra (TA_035)")

    ta035 = results.get("ta035", {})
    summary = ta035.get("summary", pd.DataFrame())
    detail = ta035.get("detail", pd.DataFrame())

    if not summary.empty:
        st.markdown("#### 📊 Group Theo Phân Loại")
        st.dataframe(format_df(summary), use_container_width=True, hide_index=True)

    if not detail.empty:
        st.markdown("#### 📋 Chi Tiết Bảng Kê Bán Ra")
        dupes = detect_duplicate_invoices(detail)
        if not dupes.empty:
            st.warning(f"⚠️ Phát hiện {len(dupes)} dòng trùng số hóa đơn!")
        st.dataframe(format_df(detail.head(200)), use_container_width=True, hide_index=True)
        st.caption(f"Hiển thị {min(200, len(detail))}/{len(detail)} dòng")


# ============================================================
# TAB 3: MUA VÀO (TA036)
# ============================================================
def render_tab_muavao(results):
    """Tab mua vào - TA036."""
    st.markdown("### 📥 Bảng Kê Mua Vào (TA_036)")

    ta036 = results.get("ta036", {})
    summary = ta036.get("summary", pd.DataFrame())
    summary_ts = ta036.get("summary_thue_suat", pd.DataFrame())
    mst_warnings = ta036.get("mst_warnings", pd.DataFrame())
    detail = ta036.get("detail", pd.DataFrame())

    # MST Warnings
    if isinstance(mst_warnings, pd.DataFrame) and not mst_warnings.empty:
        st.markdown('<div class="alert-warning">⚠️ <strong>Cảnh Báo MST</strong></div>', unsafe_allow_html=True)
        st.dataframe(format_df(mst_warnings), use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)
    with col1:
        if not summary.empty:
            st.markdown("#### 📊 Group Theo Tài Khoản")
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)
    with col2:
        if not summary_ts.empty:
            st.markdown("#### 📊 Group Theo Thuế Suất")
            st.dataframe(format_df(summary_ts), use_container_width=True, hide_index=True)

    if not detail.empty:
        st.markdown("#### 📋 Chi Tiết Bảng Kê Mua Vào")
        dupes = detect_duplicate_invoices(detail)
        if not dupes.empty:
            st.warning(f"⚠️ Phát hiện {len(dupes)} dòng trùng số hóa đơn!")
        st.dataframe(format_df(detail.head(300)), use_container_width=True, hide_index=True)
        st.caption(f"Hiển thị {min(300, len(detail))}/{len(detail)} dòng")


# ============================================================
# TAB 4: KIỂM DÒ (CHỮ T)
# ============================================================
def render_tab_kiemdo(taxvta):
    """Tab kiểm dò theo tài khoản - Layout chữ T kế toán."""
    st.markdown("### 🔍 Kiểm Dò Theo Tài Khoản")

    # Cross-check table
    cross_checks = taxvta.get("cross_checks", [])
    if cross_checks:
        cc_df = pd.DataFrame(cross_checks)
        cc_display = cc_df[["name", "source_a", "value_a", "source_b", "value_b", "chenh_lech", "status"]].copy()
        cc_display.columns = ["Nội dung", "Nguồn A", "Giá trị A", "Nguồn B", "Giá trị B", "Chênh lệch", "Trạng thái"]

        # Format số với dấu chấm
        for col in ["Giá trị A", "Giá trị B", "Chênh lệch"]:
            cc_display[col] = cc_display[col].apply(lambda x: fmt(x) if pd.notna(x) else "")

        styled = style_applymap(style_applymap(cc_display.style, highlight_diff, subset=["Chênh lệch"]), highlight_ok, subset=["Trạng thái"])

        st.dataframe(styled, use_container_width=True, hide_index=True)

    st.divider()

    # T-Account HTML
    kiem_do = taxvta.get("kiem_do_tk", {})
    if kiem_do:
        html = build_t_account_html(kiem_do)
        st.markdown(html, unsafe_allow_html=True)


# ============================================================
# TAB 5: CẢNH BÁO RỦI RO
# ============================================================
def render_tab_canhbao(results):
    """Tab cảnh báo rủi ro."""
    st.markdown("### ⚠️ Cảnh Báo Rủi Ro")

    ta036 = results.get("ta036", {})
    ta035 = results.get("ta035", {})
    detail_mv = ta036.get("detail", pd.DataFrame())
    detail_br = ta035.get("detail", pd.DataFrame())

    # === 1. Tiền mặt ≥ 20 triệu ===
    st.markdown("#### 💵 HĐ Mua vào ≥ 20 triệu - Thanh toán Tiền mặt")
    st.caption("Theo quy định, HĐ mua hàng/dịch vụ ≥ 20 triệu thanh toán bằng tiền mặt sẽ không được khấu trừ thuế GTGT")

    cash_risks = detect_cash_risk(detail_mv)
    if not cash_risks.empty:
        st.markdown(f'<div class="alert-danger">🚨 Phát hiện <strong>{len(cash_risks)}</strong> hóa đơn có rủi ro!</div>', unsafe_allow_html=True)
        display_cols = ["KyHieuHD", "SoHD", "NgayPhatHanh", "TenNguoiBan", "MSTNguoiBan", "MatHang", "DoanhSoChuaThue", "ThueGTGT", "SoTienDaTT", "GhiChu"]
        available_cols = [c for c in display_cols if c in cash_risks.columns]
        st.dataframe(format_df(cash_risks[available_cols]), use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="alert-success">✅ Không phát hiện HĐ tiền mặt ≥ 20 triệu</div>', unsafe_allow_html=True)

    st.divider()

    # === 2. Trùng HĐ mua vào ===
    st.markdown("#### 🔄 HĐ Mua vào bị trùng số trong kỳ")
    dupes_mv = detect_duplicate_invoices(detail_mv)
    if not dupes_mv.empty:
        st.markdown(f'<div class="alert-danger">🚨 Phát hiện <strong>{len(dupes_mv)}</strong> dòng trùng!</div>', unsafe_allow_html=True)
        display_cols = ["KyHieuHD", "SoHD", "NgayPhatHanh", "TenNguoiBan", "DoanhSoChuaThue", "ThueGTGT"]
        available_cols = [c for c in display_cols if c in dupes_mv.columns]
        st.dataframe(format_df(dupes_mv[available_cols]), use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="alert-success">✅ Không phát hiện HĐ trùng (Mua vào)</div>', unsafe_allow_html=True)

    st.divider()

    # === 3. Trùng HĐ bán ra ===
    st.markdown("#### 🔄 HĐ Bán ra bị trùng số trong kỳ")
    dupes_br = detect_duplicate_invoices(detail_br)
    if not dupes_br.empty:
        st.markdown(f'<div class="alert-danger">🚨 Phát hiện <strong>{len(dupes_br)}</strong> dòng trùng!</div>', unsafe_allow_html=True)
        display_cols = ["KyHieuHD", "SoHD", "NgayPhatHanh", "TenNguoiMua", "DoanhSoChuaThue", "ThueGTGT"]
        available_cols = [c for c in display_cols if c in dupes_br.columns]
        st.dataframe(format_df(dupes_br[available_cols]), use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="alert-success">✅ Không phát hiện HĐ trùng (Bán ra)</div>', unsafe_allow_html=True)

    st.divider()

    # === 4. MST không hợp lệ ===
    st.markdown("#### 🆔 MST Không Hợp Lệ (≠ 10 ký tự)")
    mst_warnings = ta036.get("mst_warnings", pd.DataFrame())
    if isinstance(mst_warnings, pd.DataFrame) and not mst_warnings.empty:
        mst_count = ta036.get("mst_warning_count", len(mst_warnings))
        st.markdown(f'<div class="alert-warning">⚠️ Phát hiện <strong>{mst_count}</strong> MST cần kiểm tra</div>', unsafe_allow_html=True)
        st.dataframe(format_df(mst_warnings), use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="alert-success">✅ Tất cả MST đều hợp lệ</div>', unsafe_allow_html=True)


# ============================================================
# TAB 6: DRILL DOWN
# ============================================================
def render_tab_drilldown(results):
    """Tab drill down - tìm HĐ thiếu/thừa."""
    st.markdown("### 🔎 Tìm Kiếm Chi Tiết Lỗi (Drill Down)")
    st.caption("So sánh số hóa đơn giữa Bảng kê (TA035/TA036) và Sổ cái (TA030) để tìm HĐ thiếu/thừa")

    # === Bán ra: TA035 vs TA030_3331 ===
    st.markdown("#### 📤 Bán Ra: TA_035 vs Sổ cái TK33311")

    ta035_detail = results.get("ta035", {}).get("detail", pd.DataFrame())
    ta030_3331_detail = results.get("ta030_3331", {}).get("detail", pd.DataFrame())

    so_ct_col_3331 = "SoCT" if "SoCT" in (ta030_3331_detail.columns if isinstance(ta030_3331_detail, pd.DataFrame) and not ta030_3331_detail.empty else []) else None

    if not ta035_detail.empty and so_ct_col_3331:
        only_ta035, only_ta030 = drill_down_compare(
            ta035_detail, ta030_3331_detail,
            key_col="SoHD",
            name_a="TA035", name_b="TA030"
        )

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Có trên TA035 nhưng KHÔNG có trên Sổ cái** ({len(only_ta035)} dòng)")
            if not only_ta035.empty:
                display_cols = [c for c in ["SoHD", "KyHieuHD", "NgayPhatHanh", "TenNguoiMua", "DoanhSoChuaThue", "ThueGTGT"] if c in only_ta035.columns]
                st.dataframe(format_df(only_ta035[display_cols]), use_container_width=True, hide_index=True)
            else:
                st.markdown('<div class="alert-success">✅ Khớp hoàn toàn</div>', unsafe_allow_html=True)

        with col2:
            st.markdown(f"**Có trên Sổ cái nhưng KHÔNG có trên TA035** ({len(only_ta030)} dòng)")
            if not only_ta030.empty:
                display_cols = [c for c in ["SoCT", "NgayGD", "DienGiai", "PSCo", "PSNo"] if c in only_ta030.columns]
                st.dataframe(format_df(only_ta030[display_cols]), use_container_width=True, hide_index=True)
            else:
                st.markdown('<div class="alert-success">✅ Khớp hoàn toàn</div>', unsafe_allow_html=True)
    else:
        st.info("ℹ️ Không đủ dữ liệu để so sánh TA035 vs Sổ cái TK33311")

    st.divider()

    # === Mua vào: TA036 vs TA030_1331 ===
    st.markdown("#### 📥 Mua Vào: TA_036 vs Sổ cái TK1331")

    ta036_detail = results.get("ta036", {}).get("detail", pd.DataFrame())
    ta030_1331_detail = results.get("ta030_1331", {}).get("detail", pd.DataFrame())

    if not ta036_detail.empty and isinstance(ta030_1331_detail, pd.DataFrame) and not ta030_1331_detail.empty:
        only_ta036, only_ta030_1 = drill_down_compare(
            ta036_detail, ta030_1331_detail,
            key_col="SoHD",
            name_a="TA036", name_b="TA030"
        )

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Có trên TA036 nhưng KHÔNG có trên Sổ cái** ({len(only_ta036)} dòng)")
            if not only_ta036.empty:
                display_cols = [c for c in ["SoHD", "KyHieuHD", "NgayPhatHanh", "TenNguoiBan", "DoanhSoChuaThue", "ThueGTGT"] if c in only_ta036.columns]
                st.dataframe(format_df(only_ta036[display_cols]), use_container_width=True, hide_index=True)
            else:
                st.markdown('<div class="alert-success">✅ Khớp hoàn toàn</div>', unsafe_allow_html=True)

        with col2:
            st.markdown(f"**Có trên Sổ cái nhưng KHÔNG có trên TA036** ({len(only_ta030_1)} dòng)")
            if not only_ta030_1.empty:
                display_cols = [c for c in ["SoCT", "NgayGD", "DienGiai", "PSNo", "PSCo"] if c in only_ta030_1.columns]
                st.dataframe(format_df(only_ta030_1[display_cols]), use_container_width=True, hide_index=True)
            else:
                st.markdown('<div class="alert-success">✅ Khớp hoàn toàn</div>', unsafe_allow_html=True)
    else:
        st.info("ℹ️ Không đủ dữ liệu để so sánh TA036 vs Sổ cái TK1331")


# ============================================================
# TAB 7: CHI TIẾT
# ============================================================
def render_tab_chitiet(results):
    """Tab chi tiết - các bảng phụ."""
    st.markdown("### 📋 Chi Tiết Các Bảng Dữ Liệu")

    sub_tabs = st.tabs([
        "GL0903 (TK511)", "KD4A", "GCS", "TK333111",
        "TK33895", "Nhôm TC", "TK13311", "TK13313"
    ])

    with sub_tabs[0]:
        gl0903 = results.get("gl0903", {})
        summary = gl0903.get("summary", pd.DataFrame())
        detail = gl0903.get("detail", pd.DataFrame())
        if not summary.empty:
            st.markdown("**Group Theo Phân Loại**")
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)
        if not detail.empty:
            st.markdown("**Chi Tiết Giao Dịch**")
            st.dataframe(format_df(detail.head(200)), use_container_width=True, hide_index=True)

    with sub_tabs[1]:
        kd4a = results.get("kd4a", {})
        if kd4a:
            stats = {k: fmt(v) for k, v in kd4a.items() if isinstance(v, (int, float))}
            cols = st.columns(min(len(stats), 4))
            for i, (k, v) in enumerate(stats.items()):
                with cols[i % len(cols)]:
                    st.metric(k.replace("_", " ").title(), v)

    with sub_tabs[2]:
        bchdon = results.get("bchdon", {})
        summary = bchdon.get("summary", pd.DataFrame())
        detail = bchdon.get("detail", pd.DataFrame())
        if not summary.empty:
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)
        if not detail.empty:
            st.dataframe(format_df(detail.head(100)), use_container_width=True, hide_index=True)

    with sub_tabs[3]:
        ta030_3331 = results.get("ta030_3331", {})
        du_dk = ta030_3331.get("du_dau_ky", {})
        summary = ta030_3331.get("summary", pd.DataFrame())
        if du_dk:
            cols = st.columns(len(du_dk))
            for i, (tk, vals) in enumerate(du_dk.items()):
                with cols[i]:
                    st.metric(f"{tk} Dư ĐK Nợ", fmt(vals.get("no", 0)))
        if not summary.empty:
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)

    with sub_tabs[4]:
        gl038 = results.get("gl038", {})
        cols = st.columns(4)
        with cols[0]:
            st.metric("Dư Đầu Kỳ", fmt(gl038.get("du_dau_ky", 0)))
        with cols[1]:
            st.metric("Tổng PS Nợ", fmt(gl038.get("tong_ps_no", 0)))
        with cols[2]:
            st.metric("Tổng PS Có", fmt(gl038.get("tong_ps_co", 0)))
        with cols[3]:
            st.metric("Dư Cuối Kỳ", fmt(gl038.get("du_cuoi_ky", 0)))
        summary = gl038.get("summary", pd.DataFrame())
        if not summary.empty:
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)

    with sub_tabs[5]:
        nhomtc = results.get("nhomtc", {})
        st.metric("Tổng Tiền (DIEN00)", fmt(nhomtc.get("total", 0)))
        detail = nhomtc.get("detail", pd.DataFrame())
        if not detail.empty:
            st.dataframe(format_df(detail.head(100)), use_container_width=True, hide_index=True)

    with sub_tabs[6]:
        ta030_1331 = results.get("ta030_1331", {})
        du_dk = ta030_1331.get("du_dau_ky", {})
        summary = ta030_1331.get("summary", pd.DataFrame())
        if du_dk:
            cols = st.columns(min(len(du_dk), 4))
            for i, (tk, vals) in enumerate(du_dk.items()):
                with cols[i % len(cols)]:
                    st.metric(f"{tk} Dư ĐK Nợ", fmt(vals.get("no", 0)))
        if not summary.empty:
            st.dataframe(format_df(summary), use_container_width=True, hide_index=True)

    with sub_tabs[7]:
        ta030_1331 = results.get("ta030_1331", {})
        detail = ta030_1331.get("detail", pd.DataFrame())
        if isinstance(detail, pd.DataFrame) and not detail.empty:
            tk13313 = detail[detail["TaiKhoan"] == "TK13313"] if "TaiKhoan" in detail.columns else pd.DataFrame()
            if not tk13313.empty:
                st.dataframe(format_df(tk13313), use_container_width=True, hide_index=True)
            else:
                st.info("Không có dữ liệu TK13313")
        else:
            st.info("Không có dữ liệu TK13313")


# ============================================================
# TAB 8: BẢNG KÊ 01 (NQ142 - Giảm thuế GTGT)
# ============================================================
def render_tab_bangke(bangke01, config):
    """Tab Bảng Kê 01 - Giảm thuế GTGT theo NQ142."""
    st.markdown("### 📝 Bảng Kê 01 - Giảm Thuế GTGT (NQ142/TT80)")

    if bangke01 is None:
        st.info("ℹ️ Chưa có dữ liệu. Nhấn **Chạy Báo Cáo** trước.")
        return

    month_label = get_report_month_label(config) if config else ""
    st.caption(f"Kỳ tính thuế: **{month_label}**")

    # === PHẦN I: MUA VÀO ===
    st.markdown("#### I. Hàng hóa, dịch vụ mua vào trong kỳ (Thuế suất 8%)")

    mv_summary = bangke01.get("muavao_summary", pd.DataFrame())
    if not mv_summary.empty:
        st.dataframe(format_df(mv_summary), use_container_width=True, hide_index=True)

        # Tổng cộng mua vào
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tổng giá trị chưa thuế", fmt(bangke01.get("tong_muavao_chuathue", 0)))
        with col2:
            st.metric("Tổng thuế GTGT", fmt(bangke01.get("tong_muavao_thue", 0)))
    else:
        st.info("Không có HĐ mua vào thuế suất 8% trong kỳ")

    st.divider()

    # === PHẦN II: BÁN RA ===
    st.markdown("#### II. Hàng hóa, dịch vụ bán ra trong kỳ (Thuế suất 8%)")

    br_summary = bangke01.get("banra_summary", pd.DataFrame())
    if not br_summary.empty:
        st.dataframe(format_df(br_summary), use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tổng giá trị chưa thuế", fmt(bangke01.get("tong_banra_chuathue", 0)))
        with col2:
            st.metric("Tổng thuế GTGT được giảm", fmt(bangke01.get("tong_banra_thue_duocgiam", 0)))
    else:
        st.info("Không có HĐ bán ra thuế suất 8% trong kỳ")

    st.divider()

    # === NÚT XUẤT FILE ===
    st.markdown("#### 📥 Xuất File Bảng Kê 01")
    if st.button("📥 Xuất Excel Bảng Kê 01", key="btn_export_bangke"):
        try:
            output = _export_bangke01_excel(bangke01, config)
            st.download_button(
                label="💾 Tải Bảng Kê 01",
                data=output,
                file_name=f"BangKe01_NQ142_{config.get('report_month', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_bangke"
            )
        except Exception as e:
            st.error(f"❌ Lỗi xuất file: {str(e)}")


def _export_bangke01_excel(bangke01, config):
    """Xuất file Excel Bảng Kê 01."""
    from openpyxl.styles import Font
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet Mua vào
        mv = bangke01.get("muavao_summary", pd.DataFrame())
        if not mv.empty:
            mv.to_excel(writer, sheet_name="I.Mua vào", index=False)

            # Thêm dòng tổng cộng
            ws = writer.sheets["I.Mua vào"]
            r = len(mv) + 2
            ws.cell(row=r, column=1, value="").font = Font(bold=True)
            ws.cell(row=r, column=2, value="TỔNG CỘNG").font = Font(bold=True)
            ws.cell(row=r, column=3, value=bangke01.get("tong_muavao_chuathue", 0)).number_format = '#,##0'
            ws.cell(row=r, column=4, value=bangke01.get("tong_muavao_thue", 0)).number_format = '#,##0'

        # Sheet Bán ra
        br = bangke01.get("banra_summary", pd.DataFrame())
        if not br.empty:
            br.to_excel(writer, sheet_name="II.Bán ra", index=False)

            ws = writer.sheets["II.Bán ra"]
            r = len(br) + 2
            ws.cell(row=r, column=1, value="").font = Font(bold=True)
            ws.cell(row=r, column=2, value="TỔNG CỘNG").font = Font(bold=True)
            ws.cell(row=r, column=3, value=bangke01.get("tong_banra_chuathue", 0)).number_format = '#,##0'
            ws.cell(row=r, column=6, value=bangke01.get("tong_banra_thue_duocgiam", 0)).number_format = '#,##0'

    output.seek(0)
    return output.getvalue()


# ============================================================
# TAB 9: TỜ KHAI 01/GTGT
# ============================================================
def render_tab_tokhai(tokhai01, config):
    """Tab Tờ Khai 01/GTGT."""
    st.markdown("### 📄 Tờ Khai Thuế GTGT (Mẫu 01/GTGT)")

    if tokhai01 is None:
        st.info("ℹ️ Chưa có dữ liệu. Nhấn **Chạy Báo Cáo** trước.")
        return

    month_label = get_report_month_label(config) if config else ""
    st.caption(f"Kỳ tính thuế: **{month_label}** | Đơn vị: đồng Việt Nam")

    # Header thông tin
    st.markdown("""
    **[01a]** Hoạt động sản xuất kinh doanh: **Hoạt động sản xuất kinh doanh điện thông thường**  
    **[04]** Chi nhánh Tổng Công ty Điện lực TP.HCM TNHH - Công ty Điện lực Vũng Tàu  
    **[05]** MST: **0300951119-010**
    """)

    st.divider()

    # === Bảng tờ khai ===
    o = tokhai01  # shortcut

    # Helper để tạo dòng
    def tk_row(stt, chi_tieu, ma_o, gia_tri_hhdv, thue_gtgt):
        return {
            "STT": stt,
            "Chỉ tiêu": chi_tieu,
            "Mã ô": ma_o,
            "Giá trị HHDV": gia_tri_hhdv,
            "Thuế GTGT": thue_gtgt
        }

    rows = []

    # A - Không phát sinh
    rows.append(tk_row("A", "Không phát sinh hoạt động mua, bán trong kỳ (đánh dấu \"X\")", "[21]", "", ""))

    # B - Thuế GTGT còn được khấu trừ kỳ trước
    rows.append(tk_row("B", "Thuế GTGT còn được khấu trừ kỳ trước chuyển sang", "[22]", "", o["o22"]))

    # C - Kê khai thuế GTGT phải nộp ngân sách
    rows.append(tk_row("C", "Kê khai thuế GTGT phải nộp NSNN", "", "", ""))

    # I - Mua vào
    rows.append(tk_row("I", "Hàng hóa, dịch vụ mua vào trong kỳ", "", "", ""))
    rows.append(tk_row("1", "Giá trị và thuế GTGT của HHDV mua vào", "[23]", o["o23"], o["o24"]))
    rows.append(tk_row("", "  Trong đó: hàng hóa, dịch vụ nhập khẩu", "[23a]", o["o23a"], o["o24a"]))
    rows.append(tk_row("2", "Thuế GTGT của HHDV mua vào được khấu trừ kỳ này", "[25]", "", o["o25"]))

    # II - Bán ra
    rows.append(tk_row("II", "Hàng hóa, dịch vụ bán ra trong kỳ", "", "", ""))
    rows.append(tk_row("", "  HHDV bán ra không chịu thuế GTGT", "[26]", o["o26"], ""))
    rows.append(tk_row("1", "  HHDV bán ra chịu thuế GTGT ([27]=[29]+[30]+[32]+[32a])", "[27]", o["o27"], o["o28"]))
    rows.append(tk_row("a", "    HHDV bán ra chịu thuế suất 0%", "[29]", o["o29"], ""))
    rows.append(tk_row("b", "    HHDV bán ra chịu thuế suất 5%", "[30]", o["o30"], o["o31"]))
    rows.append(tk_row("c", "    HHDV bán ra chịu thuế suất 10%", "[32]", o["o32"], o["o33"]))
    rows.append(tk_row("d", "    HHDV bán ra không tính thuế", "[32a]", o["o32a"], ""))
    rows.append(tk_row("3", "Tổng DT và thuế GTGT của HHDV bán ra ([34]=[26]+[27]; [35]=[28])", "[34]", o["o34"], o["o35"]))

    # III - Thuế phát sinh
    rows.append(tk_row("III", "Thuế GTGT phát sinh trong kỳ ([36]=[35]-[25])", "[36]", "", o["o36"]))

    # IV - Điều chỉnh
    rows.append(tk_row("IV", "Điều chỉnh tăng, giảm thuế GTGT còn được khấu trừ các kỳ trước", "", "", ""))
    rows.append(tk_row("1", "  Điều chỉnh giảm", "[37]", "", o["o37"]))
    rows.append(tk_row("2", "  Điều chỉnh tăng", "[38]", "", o["o38"]))

    # V - Thuế nhận bàn giao
    rows.append(tk_row("V", "Thuế GTGT nhận bàn giao được khấu trừ trong kỳ", "[39a]", "", o["o39a"]))

    # VI - Xác định nghĩa vụ
    rows.append(tk_row("VI", "Xác định nghĩa vụ thuế GTGT phải nộp trong kỳ", "", "", ""))
    rows.append(tk_row("1", "  Thuế GTGT phải nộp từ HĐSXKD ({[40a]=[36]-[22]+[37]-[38]-[39a]} ≥ 0)", "[40a]", "", o["o40a"]))
    rows.append(tk_row("2", "  Thuế GTGT mua vào dự án đầu tư bù trừ", "[40b]", "", o["o40b"]))
    rows.append(tk_row("3", "  Thuế GTGT còn phải nộp trong kỳ ([40]=[40a]-[40b])", "[40]", "", o["o40"]))
    rows.append(tk_row("4", "  Thuế GTGT chưa khấu trừ hết kỳ này", "[41]", "", o["o41"]))
    rows.append(tk_row("4.1", "  Thuế GTGT đề nghị hoàn ([42] ≤ [41])", "[42]", "", o["o42"]))
    rows.append(tk_row("4.2", "  Thuế GTGT chuyển kỳ sau ([43]=[41]-[42])", "[43]", "", o["o43"]))

    tk_df = pd.DataFrame(rows)

    # Format số
    for col in ["Giá trị HHDV", "Thuế GTGT"]:
        tk_df[col] = tk_df[col].apply(lambda x: fmt(x) if isinstance(x, (int, float)) and x != 0 else (str(x) if x != "" else ""))

    # Highlight styling
    def style_tokhai(row):
        styles = [""] * len(row)
        stt = str(row["STT"])
        if stt in ["A", "B", "C", "I", "II", "III", "IV", "V", "VI"]:
            styles = ["font-weight: bold; background-color: rgba(70,130,180,0.15)"] * len(row)
        elif stt in ["1", "2", "3", "4"]:
            styles = ["font-weight: 600"] * len(row)
        return styles

    styled = tk_df.style.apply(style_tokhai, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True, height=900)

    st.divider()

    # === NÚT XUẤT FILE ===
    st.markdown("#### 📥 Xuất File Tờ Khai 01/GTGT")
    if st.button("📥 Xuất Excel Tờ Khai 01", key="btn_export_tokhai"):
        try:
            output = _export_tokhai01_excel(tokhai01, config)
            st.download_button(
                label="💾 Tải Tờ Khai 01",
                data=output,
                file_name=f"ToKhai_01GTGT_{config.get('report_month', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_tokhai"
            )
        except Exception as e:
            st.error(f"❌ Lỗi xuất file: {str(e)}")


def _export_tokhai01_excel(tokhai01, config):
    """Xuất file Excel Tờ Khai 01/GTGT."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Tờ Khai 01-GTGT"

    # Styles
    hdr_font = Font(name="Times New Roman", bold=True, size=12)
    data_font = Font(name="Times New Roman", size=11)
    bold_font = Font(name="Times New Roman", bold=True, size=11)
    num_fmt = '#,##0'
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    header_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    # Header
    ws.merge_cells("A1:E1")
    ws["A1"] = "TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG (MẪU SỐ 01/GTGT)"
    ws["A1"].font = Font(name="Times New Roman", bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:E2")
    month_label = get_report_month_label(config) if config else ""
    ws["A2"] = f"Kỳ tính thuế: {month_label}"
    ws["A2"].font = data_font
    ws["A2"].alignment = Alignment(horizontal="center")

    # Column widths
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22

    # Table header
    row = 4
    headers = ["STT", "Chỉ tiêu", "Mã ô", "Giá trị HHDV\n(chưa có thuế GTGT)", "Thuế GTGT"]
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=j + 1, value=h)
        c.font = bold_font
        c.fill = header_fill
        c.border = thin_border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    o = tokhai01
    # Data rows
    data_rows = [
        ("A", "Không phát sinh hoạt động mua, bán trong kỳ", "[21]", "", "", True),
        ("B", "Thuế GTGT còn được khấu trừ kỳ trước chuyển sang", "[22]", None, o["o22"], True),
        ("C", "Kê khai thuế GTGT phải nộp NSNN", "", None, None, True),
        ("I", "Hàng hóa, dịch vụ mua vào trong kỳ", "", None, None, True),
        ("1", "Giá trị và thuế GTGT của HHDV mua vào", "[23]", o["o23"], o["o24"], False),
        ("", "  Trong đó: HHDV nhập khẩu", "[23a]", o["o23a"], o["o24a"], False),
        ("2", "Thuế GTGT HHDV mua vào được khấu trừ kỳ này", "[25]", None, o["o25"], False),
        ("II", "Hàng hóa, dịch vụ bán ra trong kỳ", "", None, None, True),
        ("", "HHDV bán ra không chịu thuế GTGT", "[26]", o["o26"], None, False),
        ("1", "HHDV bán ra chịu thuế GTGT", "[27]", o["o27"], o["o28"], False),
        ("a", "  HHDV bán ra chịu thuế suất 0%", "[29]", o["o29"], None, False),
        ("b", "  HHDV bán ra chịu thuế suất 5%", "[30]", o["o30"], o["o31"], False),
        ("c", "  HHDV bán ra chịu thuế suất 10%", "[32]", o["o32"], o["o33"], False),
        ("d", "  HHDV bán ra không tính thuế", "[32a]", o["o32a"], None, False),
        ("3", "Tổng DT và thuế GTGT HHDV bán ra", "[34]", o["o34"], o["o35"], False),
        ("III", "Thuế GTGT phát sinh trong kỳ ([36]=[35]-[25])", "[36]", None, o["o36"], True),
        ("IV", "Điều chỉnh tăng, giảm thuế GTGT còn được khấu trừ các kỳ trước", "", None, None, True),
        ("1", "  Điều chỉnh giảm", "[37]", None, o["o37"], False),
        ("2", "  Điều chỉnh tăng", "[38]", None, o["o38"], False),
        ("V", "Thuế GTGT nhận bàn giao được khấu trừ trong kỳ", "[39a]", None, o["o39a"], True),
        ("VI", "Xác định nghĩa vụ thuế GTGT phải nộp trong kỳ", "", None, None, True),
        ("1", "  Thuế GTGT phải nộp từ HĐSXKD", "[40a]", None, o["o40a"], False),
        ("2", "  Thuế GTGT mua vào dự án đầu tư bù trừ", "[40b]", None, o["o40b"], False),
        ("3", "  Thuế GTGT còn phải nộp trong kỳ", "[40]", None, o["o40"], False),
        ("4", "  Thuế GTGT chưa khấu trừ hết kỳ này", "[41]", None, o["o41"], False),
        ("4.1", "  Thuế GTGT đề nghị hoàn", "[42]", None, o["o42"], False),
        ("4.2", "  Thuế GTGT chuyển kỳ sau ([43]=[41]-[42])", "[43]", None, o["o43"], False),
    ]

    for stt, ct, ma, gthhdv, tgtgt, is_section in data_rows:
        row += 1
        ws.cell(row=row, column=1, value=stt).font = bold_font if is_section else data_font
        ws.cell(row=row, column=2, value=ct).font = bold_font if is_section else data_font
        ws.cell(row=row, column=3, value=ma).font = data_font

        if gthhdv is not None and gthhdv != "":
            c = ws.cell(row=row, column=4, value=gthhdv)
            c.number_format = num_fmt
        ws.cell(row=row, column=4).font = data_font

        if tgtgt is not None and tgtgt != "":
            c = ws.cell(row=row, column=5, value=tgtgt)
            c.number_format = num_fmt
        ws.cell(row=row, column=5).font = data_font

        # Borders & section fill
        for j in range(1, 6):
            ws.cell(row=row, column=j).border = thin_border
            if is_section:
                ws.cell(row=row, column=j).fill = section_fill

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# MAIN APP
# ============================================================
def main():
    selected_month, btn_process = render_sidebar()

    # === XỬ LÝ KHI NHẤN "CẬP NHẬT DỮ LIỆU" ===
    if btn_process:
        try:
            with st.spinner(f"Đang chép dữ liệu từ các file nguồn vào file Master TAX_VTA tháng {selected_month}..."):
                updated_file = update_tax_vta_file(selected_month)
                st.session_state.processed_month = selected_month
                st.toast(f"✅ Đã nạp và lưu đè file gốc {os.path.basename(updated_file)} thành công!", icon="✅")
                st.rerun()
        except Exception as e:
            st.error(f"❌ Lỗi cập nhật dữ liệu: {str(e)}")
            with st.expander("Chi tiết lỗi"):
                st.code(traceback.format_exc())

    # === ĐỌC DỮ LIỆU MASTER VÀ CHI TIẾT CÁC TAB ===
    summary_data = get_tax_vta_summary_data(selected_month)

    if st.session_state.results is None or st.session_state.processed_month != selected_month:
        try:
            config = load_config()
            config["report_month"] = selected_month
            resolve_config_filenames(config)
            st.session_state.config = config

            results, taxvta = run_all_processors(config)
            st.session_state.results = results
            st.session_state.taxvta = taxvta
            st.session_state.bangke01 = process_bangke01(results, config)
            st.session_state.tokhai01 = process_tokhai01(results, taxvta, config)
            st.session_state.processed_month = selected_month
        except Exception:
            st.session_state.results = {}
            st.session_state.taxvta = {}

    results = st.session_state.results or {}
    taxvta = st.session_state.taxvta or {}

    month_label = f"Tháng {selected_month[5:]}/{selected_month[:4]}"
    time_label = datetime.now().strftime('%H:%M:%S %d/%m/%Y')

    # Header text
    st.markdown(
        f'<h2 style="text-align:center;color:#4fc3f7;margin:0 0 2px 0;font-size:1.4rem;font-weight:700;letter-spacing:1px;">⚡ ỨNG DỤNG KIỂM DÒ THUẾ GTGT PCVT</h2>'
        f'<p style="text-align:center;color:#8899aa;font-size:0.8rem;margin:0 0 4px 0;">📅 Dữ liệu File Master: <strong>{month_label}</strong> &nbsp;|&nbsp; Xử lý lúc: {time_label}</p>',
        unsafe_allow_html=True
    )

    # === TABS: CHỈ GIỮ 4 SHEET CỐT LÕI THEO YÊU CẦU ===
    tabs = st.tabs([
        "📊 Tổng Hợp",
        "🔍 Kiểm Dò",
        "📝 Bảng Kê 01",
        "📄 Tờ Khai 01"
    ])

    with tabs[0]:
        render_tab_tonghop(summary_data, taxvta)
    with tabs[1]:
        render_tab_kiemdo(taxvta)
    with tabs[2]:
        render_tab_bangke(st.session_state.bangke01, st.session_state.config)
    with tabs[3]:
        render_tab_tokhai(st.session_state.tokhai01, st.session_state.config)


if __name__ == "__main__":
    main()
