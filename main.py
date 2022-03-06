import pandas as pd


FILE_PATH = 'C:\\Users\\yipen\\Desktop\\SOT - FEB 2022.xlsx'
xl = pd.ExcelFile(FILE_PATH)
# print(xl.sheet_names)

required_sheets = []
for sheet_names in xl.sheet_names:
    # print(sheet_names.split(' '))
    # if sheet_names.split(' ')[0].isnumeric() and len(sheet_names.split(' ')) == 3:
    # print(sheet_names)
    if (sheet_names.split(' ')[0].isnumeric() and sheet_names.split(' ')[2].isnumeric()):
        # print(sheet_names)
        required_sheets.append(sheet_names)

# print(required_sheets)

data = []
# assuming all sheets have the same headers
# 01 JAN 2022', '02 Jan 2022'
for required_sheet in required_sheets:
    print(required_sheet)
for required_sheet in required_sheets:
    sheet = pd.read_excel(xl, sheet_name=required_sheet)
    sheet['Date'] = required_sheet
    data.append(sheet)

# for datas in data:
#     print(datas)
compiled_data = pd.concat(data, axis=0, ignore_index=True)
# print(compiled_data.dtypes)
# compiled_data['Service Remarks'] = compiled_data['ServiceComment'].astype(str) + ';' + \
#     compiled_data['Remark/ Issue (Reschedule/Cancel Date (CCC/CR: date & time inform MGL)'].astype(str)
# print(compiled_data.columns)
compiled_data.drop(
    compiled_data.index[compiled_data['DocumentNo'] == 'DocumentNo'], inplace=True)
compiled_data.drop(
    compiled_data.index[compiled_data['Team'] == 'Team'], inplace=True)
compiled_data.dropna(subset=["Team"], inplace=True)
headers = ['Team', 'DocumentNo', 'ServiceOrderNo', 'ServiceName',
           'Remark/ Issue (Reschedule/Cancel Date (CCC/CR: date & time inform MGL)', 'Status', 'Date']
filtered_data = compiled_data.loc[:, headers]
# compiled_data.drop(
#     compiled_data.index[compiled_data['Document No.'] == 'Document No.'], inplace=True)
# compiled_data.drop(
#     compiled_data.index[compiled_data['ServiceOrderNo'] == 'ServiceOrderNo'], inplace=True)
filtered_data.to_excel(
    r'C:\\Users\\yipen\\Desktop\\compiled_FEB_SOT_2022.xlsx', index=False)
