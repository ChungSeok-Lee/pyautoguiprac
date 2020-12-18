import pandas as pd 

df = pd.read_excel('./File/1차_합산본_20201215.xlsx')

df['ST'] = df['ST'].astype(str)
df['ET'] = df['ET'].astype(str)
# df['Category'] = df['Category'].astype(str)
# df['Index'] = df['Index'].astype(str)


df.to_excel('./tmp.xlsx')