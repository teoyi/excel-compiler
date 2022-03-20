import pandas as pd


FILE_PATH = "C:\\Users\\yipen\\Desktop\\SOT - JAN 2022 (12-03-22 FROM SHARON).xlsx"
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
compiled_data.drop(
    compiled_data.index[compiled_data["DocumentNo"] == "DocumentNo"], inplace=True
)
compiled_data.drop(compiled_data.index[compiled_data["Team"] == "Team"], inplace=True)
compiled_data.dropna(subset=["Team"], inplace=True)
headers = [
    "Team",
    "DocumentNo",
    "ServiceOrderNo",
    "ServiceName",
    "Remark/ Issue (Reschedule/Cancel Date (CCC/CR: date & time inform MGL)",
    "Status",
    "Date",
]
filtered_data = compiled_data.loc[:, headers]
# compiled_data.drop(
#     compiled_data.index[compiled_data['Document No.'] == 'Document No.'], inplace=True)
# compiled_data.drop(
#     compiled_data.index[compiled_data['ServiceOrderNo'] == 'ServiceOrderNo'], inplace=True)
compiled_data.to_excel(
    r"C:\\Users\\yipen\\Desktop\\compiled_JAN_SOT_2022.xlsx", index=False
)
