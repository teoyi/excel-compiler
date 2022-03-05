import pandas as pd
# pd.set_option('display.max_rows', None)

FILE_PATH = 'C:\\Users\\yipen\\Desktop\\IKEA-MGL-COPY.xlsx'
xl = pd.ExcelFile(FILE_PATH)

required_sheets = []
for sheet_names in xl.sheet_names:
    if (len(sheet_names.split(' ')) > 1):
        print(sheet_names)
        required_sheets.append(sheet_names)

data = []
# for required_sheet in required_sheets:
#     print(required_sheet)
for required_sheet in required_sheets:
    sheet = pd.read_excel(xl, sheet_name=required_sheet)

    # remove first column indexing

    # look for 'TYPE' in the cell and create subframe
    row = sheet.loc[sheet['Unnamed: 1'] == 'TYPE'].index[0]
    headers = sheet.iloc[[row]].values.tolist()
    sheet_subframe = sheet.loc[row:]
    sheet_subframe = sheet_subframe.drop(sheet_subframe.columns[0], axis=1)
    sheet_subframe['Original Sheet'] = required_sheet
    cols = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
            'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 13']
    sheet_subframe.loc[:, cols] = sheet_subframe.loc[:, cols].ffill()
    sheet_subframe = sheet_subframe.reset_index(drop=True)
    print(sheet_subframe)
    data.append(sheet_subframe)
    # sheet_subframe.to_excel(
    #     r'C:\\Users\\yipen\\Desktop\\%s-IKEA-MGL-SORTEDBY-TEAM.xlsx' % required_sheet, index=False)
# df.index = pd.Series(df.index).fillna(method='ffill')


compiled_data = pd.concat(data, axis=0, ignore_index=True)
print(compiled_data)
compiled_data.drop(
    compiled_data.index[compiled_data['Unnamed: 1'] == 'TYPE'], inplace=True)
# compiled_data = compiled_data.sort_values(by=['Unnamed: 13'])
compiled_data.to_excel(
    r'C:\\Users\\yipen\\Desktop\\IKEA-MGL-SORTEDBY-TEAM.xlsx', index=False)
