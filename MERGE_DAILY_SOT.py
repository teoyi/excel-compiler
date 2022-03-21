import pandas as pd


FILE_PATH = "C:\\Users\\yipen\\Desktop\\SOT - FEB 2022 FROM SHARON 12032022.xlsx"
xl = pd.ExcelFile(FILE_PATH)

required_sheets = []
for sheet_names in xl.sheet_names:
    if sheet_names.split(" ")[0].isnumeric() and sheet_names.split(" ")[2].isnumeric():
        required_sheets.append(sheet_names)


data = []
for required_sheet in required_sheets:
    print(required_sheet)
for required_sheet in required_sheets:
    sheet = pd.read_excel(xl, sheet_name=required_sheet)
    sheet["Date"] = required_sheet
    data.append(sheet)

compiled_data = pd.concat(data, axis=0, ignore_index=True)

# handle dropping titles for subsequent headers
compiled_data.drop(
    compiled_data.index[compiled_data["DocumentNo"] == "DocumentNo"], inplace=True
)
compiled_data.drop(compiled_data.index[compiled_data["Team"] == "Team"], inplace=True)

# handle dropping na for document no and service order no
na_free = compiled_data.dropna(subset=["DocumentNo", "ServiceOrderNo"])
only_na = compiled_data[~compiled_data.index.isin(na_free.index)]

# compiled_data.dropna(subset=["Team"], inplace=True)
# headers = [
#     "Team",
#     "DocumentNo",
#     "ServiceOrderNo",
#     "ServiceName",
#     "Remark/ Issue (Reschedule/Cancel Date (CCC/CR: date & time inform MGL)",
#     "Status",
#     "Date",
# ]
# filtered_data = compiled_data.loc[:, headers]
# compiled_data.drop(
#     compiled_data.index[compiled_data['Document No.'] == 'Document No.'], inplace=True)
# compiled_data.drop(
#     compiled_data.index[compiled_data['ServiceOrderNo'] == 'ServiceOrderNo'], inplace=True)
# compiled_data.to_excel(
#     r"C:\\Users\\yipen\\Desktop\\compiled_JAN_SOT_2022.xlsx", index=False
# )
merge_sot_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\compiled_FEB_SOT_2022.xlsx", engine="xlsxwriter"
)
na_free.to_excel(merge_sot_writer, sheet_name="filtered", index=False)
only_na.to_excel(merge_sot_writer, sheet_name="dropped", index=False)
merge_sot_writer.save()
