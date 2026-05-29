import pandas as pd

df = pd.DataFrame({'english': ['hello', 'test'], 'manipuri': ['', '']})
df.to_excel('translation/test_input.xlsx', index=False)
print('wrote translation/test_input.xlsx')
