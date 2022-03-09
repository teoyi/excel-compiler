from enum import unique
import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\NVB_DEC_UPDATED_FILTERED_09032022.xlsx'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\REVISED DEC 21 DATA.xlsx"
# FILE_PATH3 = 'C:\\Users\\yipen\\Desktop\\Jan 2022.xlsx'
# FILE_PATH4 = "C:\\Users\\yipen\\Desktop\\compiled_DEC.xlsx"
# FILE_PATH5 = "C:\\Users\\yipen\\Desktop\\compiled_DEC_SOT_2021.xlsx"
df1 = pd.read_excel(FILE_PATH1)
# df2 = pd.read_excel(FILE_PATH2)
# df3 = pd.read_excel(FILE_PATH3)
# df4 = pd.read_excel(FILE_PATH4)
# df5 = pd.read_excel(FILE_PATH5)

# create list to hold data in sequence
# datas = []

'''
Plan: 
Get list of document no from revised data 
Use the list to get the relevant status values in NVB
create a new column containing the numbers for both sheets 
'''

required_sheets = ['A & R ORDER', 'S ORDER']
for sheet in required_sheets:
    datas = []
    sheet = pd.read_excel(FILE_PATH2, sheet_name=sheet)
    sheet_docuNo = sheet['Document No.'].tolist()
    uniqueDocuNo = set(sheet_docuNo)

    for documentNo in uniqueDocuNo:
        datas

    # data.append(sheet)
    print(len(sheet_docuNo))
    df_droplog = pd.DataFrame()
    filtered_dups = sheet.duplicated(subset='Document No.', keep='first')
    filtered_keep = sheet.loc[~filtered_dups]
    df_droplog = df_droplog.append(sheet.loc[filtered_dups])
    print(df_droplog)


# prepare document no for loop
# df2_docuNo = df2['Document No.'].tolist()
# uniqueDocuNo = set(df2_docuNo)  # remove duplicates

# for documentNo in uniqueDocuNo:
#     datas.append(df2.loc[df2['Document No.'] == documentNo])
#     if (df1['Document No.'].eq(documentNo).any()):
#         datas.append(df1.loc[df1['Document No.'] == documentNo])
    # if (df3['Document No.'].eq(documentNo).any()):
    #     datas.append(df3.loc[df3['Document No.'] == documentNo])
    # if (df4['Document No.'].eq(documentNo).any()):
    #     datas.append(df4.loc[df4['Document No.'] == documentNo])
    # if (df5['Document No.'].eq(documentNo).any()):
    #     datas.append(df5.loc[df5['Document No.'] == documentNo])


# compiled_data = pd.concat(datas, axis=0, ignore_index=True)
# print(compiled_data.info())


# selecting specific columns for display
# filtered = []
# headers = ['Origin', 'Document No.', 'Service Order No.', 'Service Name', 'Service Status',
#    'Service Remarks', 'Customer Remarks', 'Reason', 'Bill to Ikea', 'Status', 'Status_2', 'Flow', 'Sales Channel']

# filtered_data = compiled_data.loc[:, headers]
# df_droplog = pd.DataFrame()
# filtered_dups = filtered_data.duplicated(subset=headers, keep='first')
# filtered_keep = filtered_data.loc[~filtered_dups]
# df_droplog = df_droplog.append(filtered_data.loc[filtered_dups])
# print(df_droplog)
# final_data = filtered_data.drop_duplicates(subset=headers)
# filtered_data = compiled_data.drop_duplicates(
#     subset=compiled_data.columns.values.tolist())

# filtered_data = pd.concat(filtered, axis=0, ignore_index=True)

# def highlight(value):
#     if value == 'SOT':
#         return 'background-color: #91BAD6'
#     elif value == 'IKEA-CN':
#         return 'background-color: red'
#     elif value == 'NVB':
#         return 'background-color: #CBC3E3'
#     elif value == 'EPOD':
#         return 'background-color: green'
#     elif value == 'SOT-DEC2021':
#         return 'background-color: #FFD580'


# compiled_data.style.applymap(highlight, subset=['Origin']).to_excel(
#     r'C:\\Users\\yipen\\Desktop\\compiled_REVISED_DEC_21_DATA.xlsx', index=False)
