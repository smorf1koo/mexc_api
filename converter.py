import pandas as pd

df = pd.read_excel('/Users/glebpetrov/Downloads/mexc_futures-2.xlsx')

print(df.head())

df['column'] = df.apply(
    lambda row: f"MEXC:{row['Базовый актив']}USDT.P",
    axis=1
)

tokens = list(df['column'])
print(tokens)


with open ("file.txt", 'w') as f:
    f.write(",".join(tokens))