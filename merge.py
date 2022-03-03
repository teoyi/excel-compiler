import pandas as pd


FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\14dec_copy.xlsx'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\31dec_copy.xlsx"
df1 = pd.read_excel(FILE_PATH1)
df2 = pd.read_excel(FILE_PATH2)

print(df1.info())
print(df2.info())
data = []
# assuming all sheets have the same headers
# 01 JAN 2022', '02 Jan 2022'
# for required_sheet in required_sheets:
#     print(required_sheet)
# for required_sheet in required_sheets:
#     sheet = pd.read_excel(xl, sheet_name=required_sheet)
#     sheet['Original Sheet'] = required_sheet
#     data.append(sheet)
data.append(df1)
data.append(df2)

# for datas in data:
#     print(datas)
compiled_data = pd.concat(data, axis=0, ignore_index=True)
print(compiled_data.info())
# compiled_data.drop(
#     compiled_data.index[compiled_data['ServicePriceExclGST'] == 'ServicePriceExclGST'], inplace=True)
# compiled_data.drop(
#     compiled_data.index[compiled_data['DocumentNo'] == 'DocumentNo'], inplace=True)
# compiled_data.drop(
#     compiled_data.index[compiled_data['ServiceOrderNo'] == 'ServiceOrderNo'], inplace=True)
compiled_data.to_excel(
    r'C:\\Users\\yipen\\Desktop\\compiled_DEC.xlsx', index=False)
