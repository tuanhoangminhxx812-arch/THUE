# -*- coding: utf-8 -*-
"""Script to explore GL_0903_TK511 and BC_HDon BC sheet."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')

# GL_0903_TK511
filepath = os.path.join(input_dir, 'GL_0903_TK511.xlsx')
df = pd.read_excel(filepath, header=None, engine='openpyxl')
print(f'GL_0903_TK511 Shape: {df.shape}')
print(f'Num columns: {df.shape[1]}')

for i in range(len(df)):
    row_vals = []
    for v in df.iloc[i].values:
        if pd.isna(v):
            row_vals.append('')
        else:
            row_vals.append(str(v)[:80])
    print(f'  Row {i}: {" | ".join(row_vals)}')

# BC_HDon - Sheet BC (main report sheet)
print(f'\n{"="*80}')
filepath = os.path.join(input_dir, 'BC_HDon_01_THopTheoNgayGCS 05-2026.xlsx')
xls = pd.ExcelFile(filepath, engine='openpyxl')
print(f'Sheets: {xls.sheet_names}')

# Read Sheet1 (raw data)
print('\n--- Sheet: [Sheet1] ---')
df1 = pd.read_excel(filepath, sheet_name='Sheet1', header=None, engine='openpyxl')
print(f'Shape: {df1.shape}')
for i in range(len(df1)):
    row_vals = []
    for v in df1.iloc[i].values:
        if pd.isna(v):
            row_vals.append('')
        else:
            row_vals.append(str(v)[:60])
    print(f'  Row {i}: {" | ".join(row_vals)}')
