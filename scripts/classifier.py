# -*- coding: utf-8 -*-
"""
Module phân loại diễn giải/nội dung thành mã phân loại.
Dựa trên quy tắc trong config/settings.yaml.
"""
from utils import load_config, clean_string


class Classifier:
    """Phân loại diễn giải thành mã: DIEN00, DIEN01, CSPK00, CSPK02, 1388DD, 4500TT, 5118QT, CHUARO."""

    def __init__(self, config: dict = None):
        if config is None:
            config = load_config()
        self.rules = config.get("classification_rules", [])
        self.default = config.get("default_classification", "CHUARO")

    def classify(self, text: str) -> str:
        """
        Phân loại một chuỗi diễn giải.
        Trả về mã phân loại (DIEN00, DIEN01, ..., CHUARO).
        """
        if not text:
            return self.default

        text_upper = clean_string(text).upper()

        for rule in self.rules:
            code = rule["code"]
            keywords = rule.get("keywords", [])
            for kw in keywords:
                if kw.upper() in text_upper:
                    return code

        return self.default

    def classify_series(self, series) -> list:
        """Phân loại toàn bộ Series pandas."""
        return [self.classify(str(v)) for v in series]

    def get_all_codes(self) -> list:
        """Trả về danh sách tất cả mã phân loại (theo thứ tự config)."""
        codes = [r["code"] for r in self.rules]
        if self.default not in codes:
            codes.append(self.default)
        return codes

    def get_description(self, code: str) -> str:
        """Trả về mô tả cho mã phân loại."""
        for rule in self.rules:
            if rule["code"] == code:
                return rule.get("description", code)
        if code == self.default:
            return "Chưa rõ phân loại"
        return code
