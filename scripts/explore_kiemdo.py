# -*- coding: utf-8 -*-
"""Quick view of the reconciliation template."""
import pandas as pd
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

fp = r'd:\DATA\DATA_THUE\ĐẦU RA\mẫu thuế kiểm dò theo tài khoản.xlsx'
df = pd.read_excel(fp, header=None, engine='openpyxl')
for i in range(len(df)):
    row_vals = [str(v)[:55] if pd.notna(v) else '' for v in df.iloc[i].values]
    print(f'Row {i}: {" | ".join(row_vals)}')
