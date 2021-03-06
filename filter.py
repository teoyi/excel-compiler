import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\compiled_IKEA_MGL_BILLING.xlsx'
FILE_PATH2 = "C:\\Users\\yipen\\Desktop\\Request_IKEA_copy.xlsx"
FILE_PATH3 = "C:\\Users\\yipen\\Desktop\\compiled_DEC.xlsx"
df1 = pd.read_excel(FILE_PATH1)
df2 = pd.read_excel(FILE_PATH2)
df3 = pd.read_excel(FILE_PATH3)
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
df1['Origin'] = 'SUQI'
df2['Origin'] = 'IKEA-CN'
df3['Origin'] = 'NVB'

print(df1.info())
print(df2.info())
print(df3.info())
# prepare document no for loop
df2_docuNo = df2['Document No.'].tolist()
# print(df2['Origin'])
# print(df2_docuNo)

for documentNo in df2_docuNo:

    datas.append(df2.loc[df2['Document No.'] == documentNo])
    if (df1['Document No.'].eq(documentNo).any()):
        datas.append(df1.loc[df1['Document No.'] == documentNo])
    if (df3['Document No.'].eq(documentNo).any()):
        datas.append(df3.loc[df3['Document No.'] == documentNo])
# print(df1.loc[df1['Document No.'] == documentNo])

# for data in datas:
#     print(data)

compiled_data = pd.concat(datas, axis=0, ignore_index=True)

# print(compiled_data.keys)
print(compiled_data.columns)
print(compiled_data.info())


def highlight(value):
    if value == 'SUQI':
        return 'background-color: #91BAD6'
    elif value == 'IKEA-CN':
        return 'background-color: red'
    elif value == 'NVB':
        return 'background-color: green'


compiled_data.style.applymap(highlight, subset=['Origin']).to_excel(
    r'C:\\Users\\yipen\\Desktop\\filtered_v4.xlsx', index=False)


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
