# -*- coding: utf-8 -*-
"""Explore TAX_VTA reference - full TAXVTA sheet and key sheets."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
ref_dir = os.path.join(base, 'THAM KHẢO')

filepath = os.path.join(ref_dir, 'TAX_VTA_2026_05.xlsx')
xls = pd.ExcelFile(filepath, engine='openpyxl')

# TAXVTA sheet - continue from row 20
df = pd.read_excel(filepath, sheet_name='TAXVTA', header=None, engine='openpyxl')
print(f'TAXVTA sheet - rows 20 to end (shape: {df.shape})')
for i in range(20, min(80, len(df))):
    row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
    print(f'  Row {i}: {" | ".join(row_vals)}')

if len(df) > 80:
    print(f'\n  ... rows 80 to {len(df)-1} ...\n')
    for i in range(max(80, len(df)-20), len(df)):
        row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
        print(f'  Row {i}: {" | ".join(row_vals)}')
