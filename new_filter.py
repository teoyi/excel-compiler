import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\Dec 21.xls.xls'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\Request_IKEA_copy.xlsx"
df1 = pd.read_excel(FILE_PATH1)
df2 = pd.read_excel(FILE_PATH2)
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
# df3['Origin'] = 'NVB'

# print(df1.info())
# print(df2.info())
# print(df3.info())
# prepare document no for loop
df2_docuNo = df2['Document No.'].tolist()
# print(df2['Origin'])
# print(df2_docuNo)

for documentNo in df2_docuNo:
    # print(documentNo)
    datas.append(df2.loc[df2['Document No.'] == documentNo])
    if (df1['Document No.'].eq(documentNo).any()):
        datas.append(df1.loc[df1['Document No.'] == documentNo])
    # datas.append(df3.loc[df3['Document No.'] == documentNo])
    # print(df1.loc[df1['Document No.'] == documentNo])

# for data in datas:
#     print(data)

compiled_data = pd.concat(datas, axis=0, ignore_index=True)
print(compiled_data.columns)

filtered = []
headers = ['Origin', 'Document No.', 'Service Order No.', 'Service Name', 'Service Status',
           'Service Remarks', 'Reason', 'Bill to Ikea', 'Status', 'Status_2', 'Flow', 'Sales Channel']

# for header in headers:
#     # print(header)
#     # print(compiled_data.loc[:, header])
#     filtered.append(compiled_data.loc[:, header])
#     # filtered.append(compiled_data.loc[header])
filtered_data = compiled_data.loc[:, headers]
# filtered_data = filtered_data.sort_values(by=['Document No.', 'Service Order No.'])
# print(filtered_data.duplicated(subset=headers))
filtered_data = filtered_data.drop_duplicates(subset=headers)

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
        return 'background-color: green'
    elif value == 'EPOD':
        return 'background-color: green'


filtered_data.style.applymap(highlight, subset=['Origin']).to_excel(
    r'C:\\Users\\yipen\\Desktop\\filtered_EPOD_v4.xlsx', index=False)

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
