# -*- coding: utf-8 -*-
"""
Tiện ích dùng chung cho toàn bộ hệ thống báo cáo thuế.
"""
import os
import yaml
import pandas as pd
from datetime import datetime, date
from pathlib import Path


def get_project_root() -> Path:
    """Trả về đường dẫn gốc của dự án (thư mục DATA_THUE)."""
    return Path(__file__).resolve().parent.parent


def load_config() -> dict:
    """Đọc file cấu hình settings.yaml."""
    config_path = get_project_root() / "config" / "settings.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_config_filenames(config: dict) -> dict:
    """
    Thay thế placeholder {MM-YYYY} trong tên file bằng tháng thực tế.
    Ví dụ: report_month = '2026-06' → {MM-YYYY} = '06-2026'
    """
    month_str = config.get("report_month", "")
    if month_str and "-" in month_str:
        # '2026-06' → '06-2026'
        parts = month_str.split("-")
        mm_yyyy = f"{parts[1]}-{parts[0]}"
        for key, fname in config.get("input_files", {}).items():
            if "{MM-YYYY}" in fname:
                config["input_files"][key] = fname.replace("{MM-YYYY}", mm_yyyy)
    return config


def get_input_dir(config: dict = None) -> Path:
    """Trả về thư mục đầu vào theo tháng báo cáo.
    Ví dụ: ĐẦU VÀO/2026-05/
    Fallback: nếu thư mục tháng không tồn tại, trả về ĐẦU VÀO/ gốc.
    """
    base = get_project_root() / "ĐẦU VÀO"
    if config:
        month_dir = base / config["report_month"]
        if month_dir.exists():
            return month_dir
    return base


def get_input_base_dir() -> Path:
    """Trả về thư mục ĐẦU VÀO gốc (chứa các subfolder tháng)."""
    return get_project_root() / "ĐẦU VÀO"


def scan_available_months() -> list:
    """
    Quét thư mục ĐẦU VÀO/ và trả về danh sách các tháng có dữ liệu.
    Mỗi tháng là 1 subfolder dạng YYYY-MM.
    Returns: list of dicts [{"month": "2026-05", "file_count": 9}, ...]
    """
    base = get_project_root() / "ĐẦU VÀO"
    months = []
    if not base.exists():
        return months

    for item in sorted(base.iterdir(), reverse=True):
        if item.is_dir():
            name = item.name
            # Kiểm tra format YYYY-MM
            try:
                datetime.strptime(name, "%Y-%m")
                file_count = len([f for f in item.iterdir() if f.is_file() and not f.name.startswith('~$')])
                months.append({
                    "month": name,
                    "file_count": file_count,
                    "label": f"Tháng {name[5:]}/{name[:4]}"
                })
            except ValueError:
                continue

    return months


def get_ref_dir() -> Path:
    return get_project_root() / "THAM KHẢO"


def get_output_dir() -> Path:
    return get_project_root() / "ĐẦU RA"


def get_report_month(config: dict) -> str:
    """Trả về tháng báo cáo dạng 'YYYY-MM'."""
    return config["report_month"]


def get_report_month_date(config: dict) -> date:
    """Trả về ngày đầu tiên của tháng báo cáo."""
    rm = config["report_month"]
    return datetime.strptime(rm, "%Y-%m").date()


def get_report_month_end(config: dict) -> date:
    """Trả về ngày cuối cùng của tháng báo cáo."""
    d = get_report_month_date(config)
    # Ngày cuối tháng
    if d.month == 12:
        return date(d.year + 1, 1, 1).replace(day=1) - pd.Timedelta(days=1)
    else:
        return date(d.year, d.month + 1, 1) - pd.Timedelta(days=1)


def get_report_month_label(config: dict) -> str:
    """Trả về nhãn tháng dạng 'Tháng MM/YYYY'."""
    d = get_report_month_date(config)
    return f"Tháng {d.month:02d}/{d.year}"


def resolve_file(filename: str, config: dict, file_key: str = None) -> Path:
    """
    Tìm file trong thư mục tháng theo cơ chế nhận diện thông minh:
    1. Tìm file chính xác theo tên.
    2. Nhận diện file theo từ khóa ERP chuẩn (hỗ trợ cả tên file ngắn như TA35.xlsx, GCS.xlsx,
       0903.xlsx, 4A.xlsx, 333111.xlsx, 33895.xlsx, TA36.xlsx, 13311.xlsx, Nhom_TC.xlsx
       lẫn các định dạng file xuất đầy đủ từ hệ thống ERP).
    3. Tìm gần đúng (loại bỏ dấu gạch, khoảng trắng, tiền tố).
    """
    input_dir = get_input_dir(config)
    exact = input_dir / filename

    # 1. Tìm chính xác
    if exact.exists():
        return exact

    files = [f for f in input_dir.iterdir() if f.is_file() and not f.name.startswith("~$") and f.suffix.lower() in [".xls", ".xlsx", ".csv"]]

    # Suy luận file_key nếu chưa truyền vào
    if not file_key:
        cfg_inputs = config.get("input_files", {})
        for k, v in cfg_inputs.items():
            if v == filename:
                file_key = k
                break

    # 2. Quy tắc nhận diện từ khóa ERP (Chuẩn theo Hướng dẫn cách viết ứng dụng ERP)
    # Khớp chính xác theo quy tắc từ khóa tên file:
    erp_rules = {
        "bchdon": lambda name: "GCS" in name or "BCHDON" in name or "HDON" in name,
        "gl0903": lambda name: "0903" in name or "GL0903" in name or "GL00903" in name or "511" in name or "903" in name,
        "kd4a": lambda name: "4A" in name or "04A" in name or "KDDN4A" in name or "BIEU04A" in name or "RPTKDDN4A" in name,
        "ta35": lambda name: ("TA35" in name or "035" in name or "01-1" in name or "01_1" in name) and ("36" not in name and "036" not in name),
        "ta035": lambda name: ("TA35" in name or "035" in name or "01-1" in name or "01_1" in name) and ("36" not in name and "036" not in name),
        "ta030_3331": lambda name: ("333111" in name or ("3331" in name and "1331" not in name) or "030_3331" in name),
        "gl038": lambda name: "33895" in name or "038" in name or "GL038" in name or "GL_038" in name,
        "ta36": lambda name: ("TA36" in name or "036" in name or "01-2" in name or "01_2" in name) and ("35" not in name and "035" not in name),
        "ta036": lambda name: ("TA36" in name or "036" in name or "01-2" in name or "01_2" in name) and ("35" not in name and "035" not in name),
        "ta030_1331": lambda name: ("13311" in name or ("1331" in name and "3331" not in name) or "030_1331" in name),
        "nhomtc": lambda name: "NHOM" in name or "TC" in name or "TOANCAU" in name or "SANLUONG" in name,
    }

    # Nếu biết file_key, ưu tiên tìm theo rule của key đó
    if file_key and file_key.lower() in erp_rules:
        rule_fn = erp_rules[file_key.lower()]
        matched_erp = [f for f in files if rule_fn(f.name.upper())]
        if len(matched_erp) == 1:
            return matched_erp[0]
        elif len(matched_erp) > 1:
            ext = Path(filename).suffix.lower()
            for c in matched_erp:
                if c.suffix.lower() == ext:
                    return c
            return matched_erp[0]

    # Nếu không có file_key, suy đoán loại file từ filename
    fn_upper = filename.upper()
    for k, rule_fn in erp_rules.items():
        if rule_fn(fn_upper):
            matched_erp = [f for f in files if rule_fn(f.name.upper())]
            if len(matched_erp) == 1:
                return matched_erp[0]
            elif len(matched_erp) > 1:
                ext = Path(filename).suffix.lower()
                for c in matched_erp:
                    if c.suffix.lower() == ext:
                        return c
                return matched_erp[0]

    # 3. Tìm gần đúng: loại bỏ _, space, TK, dots
    def clean_key(s: str) -> str:
        s_upper = s.upper()
        s_clean = s_upper.replace("_", "").replace("-", "").replace(" ", "").replace(".", "").replace("TK", "")
        return s_clean

    stem_clean = clean_key(Path(filename).stem)
    candidates = []
    for f in files:
        f_clean = clean_key(f.stem)
        if stem_clean in f_clean or f_clean in stem_clean:
            candidates.append(f)

    if len(candidates) == 1:
        return candidates[0]
    elif len(candidates) > 1:
        ext = Path(filename).suffix.lower()
        for c in candidates:
            if c.suffix.lower() == ext:
                return c
        return candidates[0]

    raise FileNotFoundError(
        f"Không tìm thấy file '{filename}' trong thư mục '{input_dir}'. "
        f"Các file hiện có: {[f.name for f in files]}"
    )


def read_excel_file(filename: str, config: dict, file_key: str = None, **kwargs) -> pd.DataFrame:
    """
    Đọc file Excel từ thư mục ĐẦU VÀO/YYYY-MM/.
    Tự động tìm file gần đúng hoặc theo từ khóa ERP.
    Tự động chọn engine phù hợp (.xls vs .xlsx).
    """
    filepath = resolve_file(filename, config, file_key=file_key)
    engine = "xlrd" if str(filepath).endswith(".xls") else "openpyxl"
    return pd.read_excel(filepath, engine=engine, **kwargs)


def read_excel_sheets(filename: str, config: dict, file_key: str = None) -> pd.ExcelFile:
    """Mở file Excel và trả về ExcelFile object để đọc nhiều sheet."""
    filepath = resolve_file(filename, config, file_key=file_key)
    engine = "xlrd" if str(filepath).endswith(".xls") else "openpyxl"
    return pd.ExcelFile(filepath, engine=engine)


def format_number(value) -> str:
    """Format số thành chuỗi có dấu phẩy phân cách hàng nghìn."""
    if pd.isna(value) or value is None:
        return ""
    try:
        v = float(value)
        if v == int(v):
            return f"{int(v):,}".replace(",", ".")
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(value)


def safe_float(value, default=0.0) -> float:
    """Chuyển đổi giá trị sang float an toàn."""
    if pd.isna(value) or value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def clean_string(value) -> str:
    """Làm sạch chuỗi: strip, loại bỏ ký tự thừa."""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()


def parse_date_flexible(value) -> date:
    """
    Parse ngày từ nhiều format khác nhau.
    Hỗ trợ: datetime, string dd/mm/yyyy, string yyyy-mm-dd, số Excel.
    """
    if pd.isna(value) or value is None:
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date() if hasattr(value, 'date') else value
    if isinstance(value, date):
        return value
    s = str(value).strip()
    # Thử dd/mm/yyyy
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m/%Y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Thử parse số Excel (số ngày kể từ 1900-01-01)
    try:
        n = int(float(s))
        if 40000 < n < 60000:  # Khoảng hợp lý cho ngày Excel
            return (datetime(1899, 12, 30) + pd.Timedelta(days=n)).date()
    except (ValueError, TypeError):
        pass
    return None
