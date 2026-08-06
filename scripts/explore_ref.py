# -*- coding: utf-8 -*-
"""Script to explore reference file TAX_VTA_2026_05."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
ref_dir = os.path.join(base, 'THAM KHẢO')

filepath = os.path.join(ref_dir, 'TAX_VTA_2026_05.xlsx')
xls = pd.ExcelFile(filepath, engine='openpyxl')
print(f'Sheets: {xls.sheet_names}')

for sheet in xls.sheet_names:
    print(f'\n{"="*80}')
    print(f'Sheet: [{sheet}]')
    print(f'{"="*80}')
    df = pd.read_excel(filepath, sheet_name=sheet, header=None, engine='openpyxl')
    print(f'Shape: {df.shape}')
    
    max_rows = 25
    for i in range(min(max_rows, len(df))):
        row_vals = []
        for v in df.iloc[i].values:
            if pd.isna(v):
                row_vals.append('')
            else:
                row_vals.append(str(v)[:60])
        print(f'  Row {i}: {" | ".join(row_vals)}')
    
    if len(df) > max_rows:
        print(f'  ... ({len(df) - max_rows} more rows)')
        for i in range(max(max_rows, len(df)-5), len(df)):
            row_vals = []
            for v in df.iloc[i].values:
                if pd.isna(v):
                    row_vals.append('')
                else:
                    row_vals.append(str(v)[:60])
            print(f'  Row {i}: {" | ".join(row_vals)}')
