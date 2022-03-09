import pandas as pd

FILE_PATH1 = 'C:\\Users\\yipen\\Desktop\\NVB_DEC_UPDATED_09032022.xlsx'

df1 = pd.read_excel(FILE_PATH1)


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

removedValues = ['Bathroom Installation', 'Call Out Kitchen/Bathroom',
                 'Picking Service', 'Parcel Delivery', 'Collection Service']

for value in removedValues:
    df1.drop(
        df1.index[df1['Service Name'] == value], inplace=True)

df1.to_excel(
    r'C:\\Users\\yipen\\Desktop\\NVB_DEC_UPDATED_FILTERED_09032022.xlsx', index=False)
