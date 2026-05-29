import pandas as pd
import sys
p = r'C:\IMALab\manipuri-english-translation-pipeline\output_BasicsofFinancialAccounting_output.xlsx'
try:
    df = pd.read_excel(p, engine='openpyxl')
except Exception as e:
    print('error', e)
    sys.exit(1)
col = 'manipuri'
if col not in df.columns:
    print('missing_col', df.columns.tolist())
    sys.exit(0)
total = len(df)
translated = (df[col].astype(str).str.strip() != '').sum()
print('total', total)
print('translated', translated)
print('\nSample translated rows (first 10 non-empty):')
nonempty = df[df[col].astype(str).str.strip() != ''][[ 'english', col ]].head(10)
for i, row in nonempty.iterrows():
    print(i+1, row['english'][:80].replace('\n',' '), '->', row[col][:80].replace('\n',' '))
