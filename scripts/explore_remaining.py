# -*- coding: utf-8 -*-
"""Explore remaining data: GL_0903_TK511 first rows, TAX_VTA sheets."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')
ref_dir = os.path.join(base, 'THAM KHẢO')

# GL_0903_TK511 - first rows
filepath = os.path.join(input_dir, 'GL_0903_TK511.xlsx')
df = pd.read_excel(filepath, header=None, engine='openpyxl')
print(f'GL_0903_TK511 first 10 rows:')
for i in range(10):
    row_vals = [str(v)[:80] if pd.notna(v) else '' for v in df.iloc[i].values]
    print(f'  Row {i}: {" | ".join(row_vals)}')

# Check diễn giải column patterns in GL_0903_TK511
print('\n--- Diễn giải patterns in GL_0903_TK511 ---')
# The diễn giải appears to be in column index 17 (based on row 5 header)
for i in range(6, len(df)):
    val = df.iloc[i, 17] if pd.notna(df.iloc[i, 17]) else ''
    if val:
        print(f'  Row {i}: {str(val)[:100]}')

# TAX_VTA reference - get sheet names and first few sheets
print(f'\n{"="*80}')
filepath = os.path.join(ref_dir, 'TAX_VTA_2026_05.xlsx')
xls = pd.ExcelFile(filepath, engine='openpyxl')
print(f'TAX_VTA Sheets: {xls.sheet_names}')

# Read the "thuế đầu ra" or main summary sheet
for sheet in ['THUE DAU RA', 'Tổng hợp', 'TT', 'KIEM DO']:
    if sheet in xls.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, engine='openpyxl')
        print(f'\n--- Sheet [{sheet}] ---')
        print(f'Shape: {df.shape}')
        for i in range(min(25, len(df))):
            row_vals = [str(v)[:50] if pd.notna(v) else '' for v in df.iloc[i].values]
            print(f'  Row {i}: {" | ".join(row_vals)}')

# Read first sheet
first_sheet = xls.sheet_names[0]
df = pd.read_excel(filepath, sheet_name=first_sheet, header=None, engine='openpyxl')
print(f'\n--- First Sheet [{first_sheet}] ---')
print(f'Shape: {df.shape}')
for i in range(min(20, len(df))):
    row_vals = [str(v)[:60] if pd.notna(v) else '' for v in df.iloc[i].values]
    print(f'  Row {i}: {" | ".join(row_vals)}')
