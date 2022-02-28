import pandas as pd


FILE_PATH = 'C:\\Users\\yipen\\Desktop\\SOT - JAN 2022.xlsx'
xl = pd.ExcelFile(FILE_PATH)
# print(xl.sheet_names)

required_sheets = []
for sheet_names in xl.sheet_names:
    # print(sheet_names.split(' '))
    if sheet_names.split(' ')[0].isnumeric() and len(sheet_names.split(' ')) == 3:
        # print(sheet_names)
        required_sheets.append(sheet_names)

# print(required_sheets)

data = []
# assuming all sheets have the same headers
# 01 JAN 2022', '02 Jan 2022'

for required_sheet in required_sheets:
    sheet = pd.read_excel(xl, sheet_name=required_sheet)
    sheet['Original Sheet'] = required_sheet
    data.append(sheet)

# for datas in data:
#     print(datas)
compiled_data = pd.concat(data, axis=0, ignore_index=True)
compiled_data.to_excel(
    r'C:\\Users\\yipen\\Desktop\\compiled_SOT.xlsx', index=False)
