# -*- coding: utf-8 -*-
"""Script to explore TA_036 and see diễn giải patterns."""
import pandas as pd
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

base = r'd:\DATA\DATA_THUE'
input_dir = os.path.join(base, 'ĐẦU VÀO')

# TA_036_TK1331 
filepath = os.path.join(input_dir, 'TA_036_TK1331.xlsx')
df = pd.read_excel(filepath, header=None, engine='openpyxl')
print(f'TA_036_TK1331 Shape: {df.shape}')

# Print header rows
for i in range(20):
    row_vals = []
    for v in df.iloc[i].values:
        if pd.isna(v):
            row_vals.append('')
        else:
            row_vals.append(str(v)[:80])
    print(f'  Row {i}: {" | ".join(row_vals)}')

# Print some data rows
print('\n--- Sample data rows ---')
for i in range(20, min(35, len(df))):
    row_vals = []
    for v in df.iloc[i].values:
        if pd.isna(v):
            row_vals.append('')
        else:
            row_vals.append(str(v)[:100])
    print(f'  Row {i}: {" | ".join(row_vals)}')

# Print last rows
print('\n--- Last rows ---')
for i in range(max(35, len(df)-10), len(df)):
    row_vals = []
    for v in df.iloc[i].values:
        if pd.isna(v):
            row_vals.append('')
        else:
            row_vals.append(str(v)[:100])
    print(f'  Row {i}: {" | ".join(row_vals)}')

# Look for "đầu tư xây dựng" patterns in diễn giải
print('\n--- Rows containing "đầu tư" or "xây dựng" ---')
for i in range(len(df)):
    for v in df.iloc[i].values:
        if pd.notna(v) and isinstance(v, str):
            vl = v.lower()
            if 'đầu tư' in vl or 'xây dựng' in vl or 'xdcb' in vl:
                row_vals = [str(x)[:60] if pd.notna(x) else '' for x in df.iloc[i].values]
                print(f'  Row {i}: {" | ".join(row_vals)}')
                break

# Check column with MST (mã số thuế)
print('\n--- Column analysis for MST ---')
# Look at header to identify MST column
for i in range(13, 18):
    row_vals = [str(x)[:60] if pd.notna(x) else '' for x in df.iloc[i].values]
    print(f'  Row {i}: {" | ".join(row_vals)}')
