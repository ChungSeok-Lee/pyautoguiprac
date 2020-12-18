import pandas as pd 

df = pd.read_excel('./File/1차_합산본_20201215.xlsx')

def EditName(i):
    FullName = df['Name'][i]
    FileName = FullName[2:-12]
    return FullName, FileName

