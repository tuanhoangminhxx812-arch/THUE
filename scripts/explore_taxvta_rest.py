# -*- coding: utf-8 -*-
"""Explore TAX_VTA reference - rows 80-134 and key sheets."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
ref_dir = os.path.join(base, 'THAM KHẢO')

filepath = os.path.join(ref_dir, 'TAX_VTA_2026_05.xlsx')

# TAXVTA sheet rows 80-134
df = pd.read_excel(filepath, sheet_name='TAXVTA', header=None, engine='openpyxl')
print('TAXVTA rows 80-134:')
for i in range(80, min(135, len(df))):
    row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
    non_empty = [r for r in row_vals if r.strip()]
    if non_empty:
        print(f'  Row {i}: {" | ".join(row_vals)}')

# Key sheets: TA35, 333111, 33895, TA36, Nhom TC
for sheet_name in ['TA35', '333111', '33895', 'TA36', 'Nhom TC', 'GCS', '0903']:
    print(f'\n{"="*60}')
    print(f'Sheet: [{sheet_name}]')
    print(f'{"="*60}')
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine='openpyxl')
    print(f'Shape: {df.shape}')
    for i in range(min(20, len(df))):
        row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
        non_empty = [r for r in row_vals if r.strip()]
        if non_empty or i < 5:
            print(f'  Row {i}: {" | ".join(row_vals)}')
    if len(df) > 20:
        print(f'  ... ({len(df)-20} more rows)')
        for i in range(max(20, len(df)-5), len(df)):
            row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
            print(f'  Row {i}: {" | ".join(row_vals)}')
