# -*- coding: utf-8 -*-
"""Xử lý rptKDDN4A (Tổng hợp bán điện)."""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from data_loader import load_kd4a


def process_kd4a(config: dict) -> dict:
    """
    Xử lý file rptKDDN4A.
    Returns dict trực tiếp chứa các số liệu.
    """
    return load_kd4a(config)
