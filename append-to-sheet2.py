from enum import unique
import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\NVB_DEC21.xlsx'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\IKEA-CN-DEC21.xlsx"
# FILE_PATH3 = 'C:\\Users\\yipen\\Desktop\\Jan 2022.xlsx'
# FILE_PATH4 = "C:\\Users\\yipen\\Desktop\\compiled_DEC.xlsx"
# FILE_PATH5 = "C:\\Users\\yipen\\Desktop\\compiled_DEC_SOT_2021.xlsx"
df1 = pd.read_excel(FILE_PATH1)
df2 = pd.ExcelFile(FILE_PATH2)
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

# include origin of data
df1.insert(loc=0, column='Origin', value='NVB')
# df2.insert(loc=0, column='Origin', value='IKEA-CN')

# get sheet names and use for loop to get relevant data
required_sheets = []
for sheet_names in df2.sheet_names:
    required_sheets.append(sheet_names)

for required_sheet in required_sheets:
    print(required_sheet)
    final_df = pd.DataFrame()

    sheet = pd.read_excel(df2, sheet_name=required_sheet)
    sheet.insert(loc=0, column='Origin', value='IKEA-CN')

    required_headers = ['Document No.', 'Service Order No.',
                        'Service Name', 'Bill to Ikea', 'Service Status']

    for header in required_headers:
        required_col = sheet[header].tolist()
        if (header == 'Service Status'):
            final_df['IKEA-CN-SERVICE-STATUS'] = pd.Series(required_col).values
        else:
            final_df[header] = pd.Series(required_col).values
    # print(final_df)
    # final_df['NVB-SERVICE-STATUS'] = ""
    nvb_data = []
    for i in final_df.index:
        # print(i)
        df_row = final_df.iloc[i]
        docuNo = df_row["Document No."]
        orderNo = df_row["Service Order No."]
        # nvb_data.append(df1.loc[(df1['Document No.'] == df_row['Document No.']) & (
        #     df1['Service Order No.'] == df_row['Service Order No.'])].iloc[0]['Service Status'])
        # print(df1.loc[(df1['Document No.'] == df_row['Document No.']) & (
        #     df1['Service Order No.'] == df_row['Service Order No.'])].iloc[0]['Service Status'])
        nvb_row = df1.loc[(df1['Document No.'] == df_row['Document No.']) & (
            df1['Service Order No.'] == df_row['Service Order No.'])]
        if not nvb_row.empty:
            nvb_data.append(nvb_row.iloc[0]['Service Status'])
        # print(df_row["Document No."])
        # print(nvbData['Service Status']
    print(nvb_data)
    final_df['NVB-SERVICE-STATUS'] = pd.Series(nvb_data)
    print(final_df)
    final_df.to_excel(
        r'C:\\Users\\yipen\\Desktop\\%s_compiled_REVISED_DEC_21_DATA.xlsx' % required_sheet, index=False)

    # print(final_df)
    # sheet_docuNo = sheet['Document No.'].tolist()
    # final_df['Document No.'] = pd.Series(sheet_docuNo).values

    # sheet_orderNo = sheet['Service Order No.'].tolist()
    # final_df['Service Order No.'] = pd.Series(sheet_orderNo).values

    # sheet_svcName = sheet['Service Name'].tolist()
    # final_df['Service Name'] = pd.Series(sheet_svcName).values

    # sheet_svcStatus = sheet['Service Status'].tolist()
    # final_df['IKEA-CN-SERVICE-STATUS'] = pd.Series(sheet_svcStatus).values

    # checker = zip(sheet_docuNo, sheet_orderNo)
    # print(list(checker)[0][0])

    # nvb_data = []
    # for checks in checker:
    #     # nvb_data.append(sheet.loc[(sheet['Document No.'] == checks[0]) & (
    #     #     sheet['Service Order No.'] == checks[1])])
    #     nvb_df = df1.loc[(df1['Document No.'] == checks[0]) & (
    #         df1['Service Order No.'] == checks[1])]
    #     # print(nvb_df)
    #     # print(NVB_data.empty())
    #     if not nvb_df.empty:
    #         nvb_data.append(nvb_df)
    # print(NVB_orderNo)
    # if (df1.loc[(df1['Document No.'] == checks[0]) & (df1['Service Order No.'] == checks[1])]):
    #     data.append(df1.loc[(df1['Document No.'] == checks[0]) & (
    #         df1['Service Order No.'] == checks[1])])
    # print(data)
    # compiled_nvb_data = pd.concat(nvb_data, axis=0, ignore_index=True)
    # print(compiled_nvb_data)
    # # print(compiled_data)
    # duplicateRowsDF = compiled_nvb_data[compiled_nvb_data.duplicated()]
    # print(duplicateRowsDF)
    # if not duplicateRowsDF.empty:
    #     duplicateRowsDF.to_excel(
    #         r'C:\\Users\\yipen\\Desktop\\%s.xlsx' % required_sheet, index=False)
    # df_droplog = pd.DataFrame()
    # filtered_dups = compiled_nvb_data.duplicated(subset=None, keep='first')
    # filtered_keep = compiled_nvb_data.loc[~filtered_dups]
    # df_droplog = df_droplog.append(compiled_nvb_data.loc[filtered_dups])
    # print('here is droplog')
    # if not df_droplog.empty:
    #     print(df_droplog)
    # filtered_keep_status = filtered_keep['Service Status'].tolist()
    # final_df['NVB-SERVICE-STATUS'] = pd.Series(filtered_keep_status).values

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

    # filtered_keep.style.applymap(highlight, subset=['Origin']).to_excel(
    #     r'C:\\Users\\yipen\\Desktop\\%s_compiled_REVISED_DEC_21_DATA.xlsx' % required_sheet, index=False)

    # final_df.to_excel(r'C:\\Users\\yipen\\Desktop\\%s.xlsx' %
    #                   required_sheet, index=False)

    # data.append(sheet)


# required_sheets = ['A & R ORDER', 'S ORDER']
# for sheet in required_sheets:
#     datas = []
#     sheet = pd.read_excel(FILE_PATH2, sheet_name=sheet)
#     sheet_docuNo = sheet['Document No.'].tolist()
#     uniqueDocuNo = set(sheet_docuNo)

#     for documentNo in uniqueDocuNo:
#         datas

#     # data.append(sheet)
#     print(len(sheet_docuNo))
#     df_droplog = pd.DataFrame()
#     filtered_dups = sheet.duplicated(subset='Document No.', keep='first')
#     filtered_keep = sheet.loc[~filtered_dups]
#     df_droplog = df_droplog.append(sheet.loc[filtered_dups])
#     print(df_droplog)


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
