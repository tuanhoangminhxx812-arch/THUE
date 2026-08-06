# -*- coding: utf-8 -*-
"""Script to explore TA_030_TK3331 and TA_036 in detail."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')

def show_file(filepath, max_rows=30):
    fname = os.path.basename(filepath)
    print(f'\n{"="*80}')
    print(f'FILE: {fname}')
    print(f'{"="*80}')
    
    engine = 'xlrd' if filepath.endswith('.xls') else 'openpyxl'
    df = pd.read_excel(filepath, header=None, engine=engine)
    print(f'Shape: {df.shape}')
    
    for i in range(min(max_rows, len(df))):
        row_vals = []
        for v in df.iloc[i].values:
            if pd.isna(v):
                row_vals.append('')
            else:
                row_vals.append(str(v)[:80])
        print(f'  Row {i}: {" | ".join(row_vals)}')
    
    if len(df) > max_rows:
        print(f'  ... ({len(df) - max_rows} more rows)')
        for i in range(max(max_rows, len(df)-5), len(df)):
            row_vals = []
            for v in df.iloc[i].values:
                if pd.isna(v):
                    row_vals.append('')
                else:
                    row_vals.append(str(v)[:80])
            print(f'  Row {i}: {" | ".join(row_vals)}')

# TA_030_TK3331
show_file(os.path.join(input_dir, 'TA_030_TK3331.xlsx'))
