import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\Dec 21.xls.xls'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\Request_IKEA_copy.xlsx"
FILE_PATH3 = 'C:\\Users\\yipen\\Desktop\\compiled_IKEA_MGL_BILLING.xlsx'
FILE_PATH4 = "C:\\Users\\yipen\\Desktop\\compiled_DEC.xlsx"
FILE_PATH5 = "C:\\Users\\yipen\\Desktop\\compiled_DEC_SOT_2021.xlsx"
df1 = pd.read_excel(FILE_PATH1)
df2 = pd.read_excel(FILE_PATH2)
df3 = pd.read_excel(FILE_PATH3)
df4 = pd.read_excel(FILE_PATH4)
df5 = pd.read_excel(FILE_PATH5)

# df3 = pd.read_excel(FILE_PATH3)
# print(xl.sheet_names)
# print(df1.columns)
# print(df2.columns)


# for datas in data:
#     print(datas)
# compiled_data = pd.concat(data, axis=0, ignore_index=True)

# compiled_data.to_excel(
#     r'C:\\Users\\yipen\\Desktop\\compiled_IKEA_MGL_BILLING.xlsx', index=False)

# df2 contains document no that is used to search for rows in df1

'''
PLAN
let data = []
for each df2 document no
    data.append df2 row
    search for row in df1
    data.append df1 row
'''

# create list to hold data in sequence
datas = []

# set origin for tracing
df1['Origin'] = 'EPOD'
df2['Origin'] = 'IKEA-CN'
df3['Origin'] = 'SUQI'
df4['Origin'] = 'NVB'
df5['Origin'] = 'SOT-DEC2021'
# print(df2.info())

# prepare document no for loop
df2_docuNo = df2['Document No.'].tolist()
uniqueDocuNo = set(df2_docuNo)  # remove duplicates
# contains_duplicates = any(uniqueDocuNo.count(
#     element) > 1 for element in uniqueDocuNo)
# print(contains_duplicates)
# print(df2['Origin'])
# print(df2_docuNo)
for documentNo in uniqueDocuNo:
    datas.append(df2.loc[df2['Document No.'] == documentNo])
    if (df1['Document No.'].eq(documentNo).any()):
        datas.append(df1.loc[df1['Document No.'] == documentNo])
    if (df3['Document No.'].eq(documentNo).any()):
        datas.append(df3.loc[df3['Document No.'] == documentNo])
    if (df4['Document No.'].eq(documentNo).any()):
        datas.append(df4.loc[df4['Document No.'] == documentNo])
    if (df5['Document No.'].eq(documentNo).any()):
        datas.append(df5.loc[df5['Document No.'] == documentNo])


compiled_data = pd.concat(datas, axis=0, ignore_index=True)
print(compiled_data.info())

filtered = []
headers = ['Origin', 'Document No.', 'Service Order No.', 'Service Name', 'Service Status',
           'Service Remarks', 'Customer Remarks', 'Reason', 'Bill to Ikea', 'Status', 'Status_2', 'Flow', 'Sales Channel']

filtered_data = compiled_data.loc[:, headers]
df_droplog = pd.DataFrame()
filtered_dups = filtered_data.duplicated(subset=headers, keep='first')
filtered_keep = filtered_data.loc[~filtered_dups]
df_droplog = df_droplog.append(filtered_data.loc[filtered_dups])
print(df_droplog)
# final_data = filtered_data.drop_duplicates(subset=headers)
# filtered_data = compiled_data.drop_duplicates(
#     subset=compiled_data.columns.values.tolist())

# filtered_data = pd.concat(filtered, axis=0, ignore_index=True)


# print(compiled_data.keys)
# print(filtered_data.info())
# print(filtered_data.info())


def highlight(value):
    if value == 'SUQI':
        return 'background-color: #91BAD6'
    elif value == 'IKEA-CN':
        return 'background-color: red'
    elif value == 'NVB':
        return 'background-color: #CBC3E3'
    elif value == 'EPOD':
        return 'background-color: green'
    elif value == 'SOT-DEC2021':
        return 'background-color: #FFD580'


filtered_data.style.applymap(highlight, subset=['Origin']).to_excel(
    r'C:\\Users\\yipen\\Desktop\\filtered_ALL_FINAL.xlsx', index=False)

# filtered_data.to_excel(
#     r'C:\\Users\\yipen\\Desktop\\filtered_EPOD.xlsx', index=False)
# (
#     df.style.applymap(color_negative_red, subset=['Diff'])
#         .applymap(color_recommend, subset=['Recommend'])
# )

# def highlight(col):
#     s = col['Origin']
#     print(s)
#     if s == 'SUQI':
#         return ['background-color: blue']
#     elif s == 'IKEA-CN':
#         return ['background-color: red']
#     elif s == 'NVB':
#         return ['background-color: green']

#
# compiled_data = compiled_data.style.apply(highlight)
# #

# compiled_data.to_excel(
#     r'C:\\Users\\yipen\\Desktop\\filtered.xlsx', index=False)
