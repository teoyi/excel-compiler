import pandas as pd

FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\Request_IKEA_copy.xlsx"
df2 = pd.read_excel(FILE_PATH2)

# print(df2.duplicated)
df2_dups = df2.duplicated()
df2_dups.to_excel(
    r'C:\\Users\\yipen\\Desktop\\ikea_dups.xlsx', index=False)
