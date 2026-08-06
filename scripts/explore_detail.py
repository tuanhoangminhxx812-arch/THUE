# -*- coding: utf-8 -*-
"""Script to explore remaining input files in detail."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')

def show_file(filepath, max_rows=15):
    fname = os.path.basename(filepath)
    print(f'\n{"="*80}')
    print(f'FILE: {fname}')
    print(f'{"="*80}')
    
    engine = 'xlrd' if filepath.endswith('.xls') else 'openpyxl'
    xls = pd.ExcelFile(filepath, engine=engine)
    print(f'Sheets: {xls.sheet_names}')
    
    # For nhom toan cau, only read last sheet (most recent)
    if 'nhom' in fname.lower() or 'Toan Cau' in fname:
        sheets_to_read = xls.sheet_names[-2:]  # Last 2 sheets
    else:
        sheets_to_read = xls.sheet_names[:2]
    
    for sheet in sheets_to_read:
        print(f'\n--- Sheet: [{sheet}] ---')
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, engine=engine)
        print(f'Shape: {df.shape}')
        
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
            for i in range(max(max_rows, len(df)-3), len(df)):
                row_vals = []
                for v in df.iloc[i].values:
                    if pd.isna(v):
                        row_vals.append('')
                    else:
                        row_vals.append(str(v)[:60])
                print(f'  Row {i}: {" | ".join(row_vals)}')

# Files to explore in detail
files_to_explore = [
    'TA_030_TK3331.xlsx',
    'TA_030_TK1331.xlsx',
    'TA_035_TK3331.xlsx',
    'TA_036_TK1331.xlsx',
    'GL_038_TK33895.xlsx',
    'GL_0903_TK511.xlsx',
    'BC_HDon_01_THopTheoNgayGCS 05-2026.xlsx',
    'BC_San luong nhom Toan Cau 05-2026.xlsx',
    'rptKDDN4A.xls',
]

for f in files_to_explore:
    show_file(os.path.join(input_dir, f))
