# -*- coding: utf-8 -*-
"""Script to explore TA files and GL files in detail."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')

def show_file(filepath, max_rows=20):
    fname = os.path.basename(filepath)
    print(f'\n{"="*80}')
    print(f'FILE: {fname}')
    print(f'{"="*80}')
    
    engine = 'xlrd' if filepath.endswith('.xls') else 'openpyxl'
    xls = pd.ExcelFile(filepath, engine=engine)
    print(f'Sheets: {xls.sheet_names}')
    
    for sheet in xls.sheet_names[:1]:
        print(f'\n--- Sheet: [{sheet}] ---')
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, engine=engine)
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
            for i in range(max(max_rows, len(df)-3), len(df)):
                row_vals = []
                for v in df.iloc[i].values:
                    if pd.isna(v):
                        row_vals.append('')
                    else:
                        row_vals.append(str(v)[:80])
                print(f'  Row {i}: {" | ".join(row_vals)}')

# TA files
for f in ['TA_030_TK3331.xlsx', 'TA_030_TK1331.xlsx', 'TA_035_TK3331.xlsx', 'GL_038_TK33895.xlsx']:
    show_file(os.path.join(input_dir, f))
