# -*- coding: utf-8 -*-
"""Custom CSS theme for Streamlit dark mode - premium financial dashboard."""

DARK_THEME_CSS = """
<style>
/* ===== Import font ===== */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ===== Ẩn các nút mặc định của Streamlit ===== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* ===== Toàn trang ===== */
.stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* ===== Metric: thu nhỏ font để hiện đầy đủ số ===== */
[data-testid="stMetricValue"] {
    font-size: 1.3rem !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.75rem !important;
}

/* ===== Sidebar ===== */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f1923 0%, #1a2332 100%);
    border-right: 1px solid #2a3f55;
}

[data-testid="stSidebar"] .stMarkdown h1 {
    color: #4fc3f7;
    text-align: center;
    font-size: 1.4rem;
    letter-spacing: 2px;
}

[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stRadio label {
    color: #8899aa;
    font-weight: 500;
}

/* ===== Metric cards ===== */
[data-testid="stMetric"] {
    background: #1e2d3d;
    border: 1px solid #2a3f55;
    border-radius: 10px;
    padding: 16px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.3);
}

[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-variant-numeric: tabular-nums;
}

/* ===== Tabs ===== */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #1a2332;
    border-radius: 10px;
    padding: 4px;
    border: 1px solid #2a3f55;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 8px 16px;
    font-weight: 600;
    font-size: 13px;
}

.stTabs [aria-selected="true"] {
    background: rgba(79, 195, 247, 0.15) !important;
    color: #4fc3f7 !important;
}

/* ===== DataFrame / Table ===== */
[data-testid="stDataFrame"] {
    border: 1px solid #2a3f55;
    border-radius: 10px;
    overflow: hidden;
}

/* ===== Buttons ===== */
.stButton > button {
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.3s ease;
}

.stButton > button:hover {
    transform: translateY(-1px);
}

/* ===== Download button ===== */
.stDownloadButton > button {
    background: linear-gradient(135deg, #388e3c, #66bb6a);
    color: white;
    border: none;
    font-weight: 600;
    border-radius: 8px;
}

/* ===== Alerts ===== */
.alert-danger {
    background: rgba(239, 83, 80, 0.1);
    border: 1px solid #ef5350;
    border-radius: 10px;
    padding: 12px 16px;
    color: #ef5350;
    margin-bottom: 12px;
}

.alert-success {
    background: rgba(102, 187, 106, 0.1);
    border: 1px solid #66bb6a;
    border-radius: 10px;
    padding: 12px 16px;
    color: #66bb6a;
    margin-bottom: 12px;
}

.alert-warning {
    background: rgba(255, 167, 38, 0.1);
    border: 1px solid #ffa726;
    border-radius: 10px;
    padding: 12px 16px;
    color: #ffa726;
    margin-bottom: 12px;
}

/* ===== T-Account (Chữ T kế toán) ===== */
.t-account-container {
    display: grid;
    grid-template-columns: 1fr 40px 1fr;
    gap: 0;
    margin: 16px 0;
}

.t-account {
    background: #1e2d3d;
    border: 2px solid #2a3f55;
    border-radius: 10px;
    overflow: hidden;
}

.t-account-title {
    text-align: center;
    padding: 14px;
    font-size: 16px;
    font-weight: 700;
    letter-spacing: 1px;
    color: white;
}

.t-account-title.left {
    background: linear-gradient(135deg, #0288d1, #4fc3f7);
}

.t-account-title.right {
    background: linear-gradient(135deg, #388e3c, #66bb6a);
}

.t-account-header {
    display: grid;
    grid-template-columns: 1.2fr 1fr 1fr;
    background: #243044;
    font-weight: 700;
    font-size: 12px;
    text-transform: uppercase;
    border-bottom: 2px solid #2a3f55;
}

.t-account-header.left span { color: #4fc3f7; }
.t-account-header.right span { color: #66bb6a; }

.t-account-header span {
    padding: 8px 12px;
    text-align: center;
    border-right: 1px solid #2a3f55;
}

.t-account-header span:last-child { border-right: none; }

.t-account-row {
    display: grid;
    grid-template-columns: 1.2fr 1fr 1fr;
    border-bottom: 1px solid rgba(42, 63, 85, 0.4);
}

.t-account-row:hover { background: rgba(79, 195, 247, 0.03); }

.t-account-row.total {
    background: rgba(79, 195, 247, 0.08);
    border-top: 2px solid #4fc3f7;
    border-bottom: 2px solid #4fc3f7;
    font-weight: 700;
}

.t-account-row.total.right {
    border-top-color: #66bb6a;
    border-bottom-color: #66bb6a;
}

.t-cell {
    padding: 10px 14px;
    font-size: 13px;
    border-right: 1px solid rgba(42, 63, 85, 0.3);
}

.t-cell:last-child { border-right: none; }
.t-cell.label { font-weight: 500; }

.t-cell.amount {
    text-align: right;
    font-variant-numeric: tabular-nums;
    font-family: 'Inter', monospace;
    font-weight: 600;
}

.t-cell.amount.blue { color: #4fc3f7; }
.t-cell.amount.green { color: #66bb6a; }
.t-cell.amount.empty { color: transparent; }

.t-cell .note {
    font-size: 10px;
    color: #5a6a7a;
    font-style: italic;
    margin-top: 2px;
}

.t-divider {
    display: flex;
    align-items: center;
    justify-content: center;
    color: #5a6a7a;
    font-size: 20px;
}

/* ===== Section headers ===== */
.section-header {
    font-size: 16px;
    font-weight: 700;
    color: #4fc3f7;
    padding: 8px 0;
    margin: 16px 0 8px;
    border-bottom: 1px solid #2a3f55;
}

/* ===== Highlight styles for pandas ===== */
.highlight-red { background-color: rgba(239, 83, 80, 0.15) !important; color: #ef5350 !important; }
.highlight-green { background-color: rgba(102, 187, 106, 0.1) !important; color: #66bb6a !important; }

/* ===== Logo ===== */
.logo-container {
    text-align: center;
    padding: 16px 0;
}

.logo-container .logo-icon {
    font-size: 48px;
    filter: drop-shadow(0 0 12px rgba(79, 195, 247, 0.5));
}

/* ===== Expander ===== */
.streamlit-expanderHeader {
    font-weight: 600;
    font-size: 14px;
}
</style>
"""


def apply_theme():
    """Inject custom CSS into Streamlit app."""
    import streamlit as st
    st.markdown(DARK_THEME_CSS, unsafe_allow_html=True)
