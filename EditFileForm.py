import pandas as pd 

df = pd.read_excel('./File/1차_합산본_20201215.xlsx')
def EditFileForm(col):
    tlst = []
    for i in range(len(df)):
        txt = str(df.iloc[i][col]).split(":")[0]
        if len(txt)<2:
            tlst += ["0"+str(df.iloc[i][col])]
        else:
            tlst += [str(df.iloc[i][col])]
    df[col] = tlst
