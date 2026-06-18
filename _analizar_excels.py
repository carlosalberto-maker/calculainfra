"""Analyze both Excel files in detail for La Perla."""
import pandas as pd
import os

f1 = r'c:\Users\ECONAS12\Downloads\Implementación Ecosistema Digital en Salud - Estado de México (v2).xlsx'
f2 = r'c:\Users\ECONAS12\Downloads\34fbb4f25843_MCIMB009215_EDS_HG_LA_PERLA_NEZA_actualizado.xlsx'

# File 1 - Find La Perla row in all sheets
print('='*80)
print('FILE 1:', os.path.basename(f1))
print('='*80)
xls1 = pd.ExcelFile(f1)
for s in xls1.sheet_names:
    df = pd.read_excel(xls1, s, header=None)
    print(f'\n=== Sheet: {s} (shape={df.shape}) ===')
    # Find rows mentioning "PERLA"
    for i, row in df.iterrows():
        row_str = '|'.join([str(v)[:40] for v in row if pd.notna(v)])
        if 'PERLA' in row_str.upper() or i < 5:
            vals = []
            for v in row:
                if pd.notna(v):
                    sv = str(v).replace('\n',' | ')
                    sv = sv[:40] if len(sv)>40 else sv
                    vals.append(sv)
                else:
                    vals.append('')
            print(f'  Row {i}: {vals}')

print('\n\n')
print('='*80)
print('FILE 2:', os.path.basename(f2))
print('='*80)
xls2 = pd.ExcelFile(f2)
print('Sheets:', xls2.sheet_names)
for s in xls2.sheet_names:
    df = pd.read_excel(xls2, s, header=None)
    print(f'\n=== {s} (shape={df.shape}) ===')
    for i, row in df.iterrows():
        vals = []
        for v in row:
            if pd.notna(v):
                sv = str(v).replace('\n',' | ')
                sv = sv[:50] if len(sv)>50 else sv
                vals.append(sv)
            else:
                vals.append('')
        if len(vals) > 10:
            vals = vals[:10] + ['...']
        print(f'  Row {i}: {vals}')
    if len(df) > 25:
        print(f'  ... ({len(df)} rows total)')
