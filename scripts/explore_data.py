# -*- coding: utf-8 -*-
"""Script to explore all input files and reference files."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')
ref_dir = os.path.join(base, 'THAM KHẢO')
output_dir = os.path.join(base, 'ĐẦU RA')

def explore_excel(filepath, max_rows=8):
    """Read and display excel file structure."""
    fname = os.path.basename(filepath)
    print(f'\n{"="*80}')
    print(f'FILE: {fname}')
    print(f'{"="*80}')
    
    try:
        if filepath.endswith('.xls'):
            xls = pd.ExcelFile(filepath, engine='xlrd')
        else:
            xls = pd.ExcelFile(filepath, engine='openpyxl')
        
        print(f'Sheets: {xls.sheet_names}')
        
        for sheet in xls.sheet_names[:3]:  # Only first 3 sheets
            print(f'\n--- Sheet: [{sheet}] ---')
            df = pd.read_excel(filepath, sheet_name=sheet, header=None,
                             engine='xlrd' if filepath.endswith('.xls') else 'openpyxl')
            print(f'Shape: {df.shape}')
            
            # Print first rows to understand structure
            for i in range(min(max_rows, len(df))):
                row_vals = []
                for v in df.iloc[i].values:
                    if pd.isna(v):
                        row_vals.append('')
                    else:
                        row_vals.append(str(v))
                print(f'  Row {i}: {" | ".join(row_vals)}')
            
            # Print last few rows
            if len(df) > max_rows:
                print(f'  ...')
                for i in range(max(max_rows, len(df)-3), len(df)):
                    row_vals = []
                    for v in df.iloc[i].values:
                        if pd.isna(v):
                            row_vals.append('')
                        else:
                            row_vals.append(str(v))
                    print(f'  Row {i}: {" | ".join(row_vals)}')
                    
    except Exception as e:
        print(f'Error: {e}')
        import traceback
        traceback.print_exc()

# Explore all input files
print('\n' + '#'*80)
print('# INPUT FILES')
print('#'*80)
for f in sorted(os.listdir(input_dir)):
    if f.startswith('~'):
        continue
    explore_excel(os.path.join(input_dir, f))

# Explore reference files
print('\n' + '#'*80)
print('# REFERENCE FILES')
print('#'*80)
for f in sorted(os.listdir(ref_dir)):
    if f.startswith('~'):
        continue
    explore_excel(os.path.join(ref_dir, f))

# Explore output template files
print('\n' + '#'*80)
print('# OUTPUT TEMPLATE FILES')
print('#'*80)
for f in sorted(os.listdir(output_dir)):
    if f.startswith('~'):
        continue
    explore_excel(os.path.join(output_dir, f))
