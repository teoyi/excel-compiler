import pandas as pd

"""'
    Headers from revised data
    EPOD NVBC SOT 

    SAD -> A + R document no. 
    HAPPY -> S document no. 

    -> NVBC Document No. + Service Order No. as ref 
    -> match with SOT and EPOD 

    3 match 1 df 
    1 match 1df 
    2 match 1df 
    keep NVBC no match bot 

    -> SOT as ref / dont need split sad happy 
    -> match with NVBC and EPOD 

    Route Date = Service Date 

    NOTE: 
    1. There are 2 Alt doc number in SOT, add it into required_headers 
    2. There are 3 service goods value in SOT, add it into required_headers
    3. EPOD team column need trimming of MGL_ from the string 
    4. SOT missing Service goods value, need to copy over from NVBC 
"""

# DECLARE HEADERS FOR COMPILED SHEET
required_headers = [
    # "EPOD",
    "SOT",
    "NVBC",
    "Driver Team",  # SOT, EPOD
    "Document No.",  # NVBC
    "Service Order No.",  # NVBC
    "Service Name",  # NVBC, SOT, EPOD
    "Service Date",  # NVBC, SOT
    "GRN",
    "Service Status",  # NVBC
    "Service Goods Value",  # NVBC, SOT
    "Service Goods Value1",  # SOT
    "Service Goods Value2",  # SOT
    "Manual Value",
    "Service Price Excl. GST",  # NVBC, SOT
    "Payout to Subco",
    "Bill to Ikea",
    "Service Comment",  # NVBC, SOT
    "CRM Case ID.",  # NVBC
    "Remark/ Issue",  # SOT
    "Status",  # SOT, EPOD
    "Service Remarks",  # EPOD
    "Status No.2",
    "Reason",  # EPOD
    "Sell-to Customer Name",  # NVBC
    "Sell-to-Address",  # NVBC
    "Routed Date",  # NVBC service date
    "Alt Doc Number (up to 3)",  # NVBC, SOT
    "Alt Doc Number1 (up to 3)",  # SOT
    "Flow",  # NVBC based on AR or S
    "Sales Channel",  # NVBC
]

# READ EXCEL FILES AND CREATE DATAFRAME
# FILE_PATH_EPOD = "C:\\Users\\yipen\\Desktop\\EPOD - JAN 22.xls"
FILE_PATH_SOT = "C:\\Users\\yipen\\Desktop\\compiled_MAR_SOT_2022.xlsx"
FILE_PATH_NVBC = "C:\\Users\\yipen\\Desktop\\mar_nvbc.xlsx"

# df_epod = pd.read_excel(FILE_PATH_EPOD)
df_sot = pd.read_excel(FILE_PATH_SOT)
df_nvbc = pd.read_excel(FILE_PATH_NVBC)

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
# nvbc_base_writer = pd.ExcelWriter(
#     "C:\\Users\\yipen\\Desktop\\NVBC_BASE_COMPILE.xlsx", engine="xlsxwriter"
# )
sot_base_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\SOT_BASE_MAR_COMPILE.xlsx", engine="xlsxwriter"
)

"""
DATA PREP
1. Trim EPOD Team col
2. Replace SOT Service Goods Value
"""
print("_______________________")
print("DATA PREP START")
# TRIM TEAM COLUMN FOR EPOD
# df_epod["Team"] = df_epod["Team"].str.replace("MGL_", "")
# df_epod["Team"] = df_epod["Team"].str.replace("MGL-", "")
# print("---EPOD TEAM TRIM COMPLETE")

# REPLACE MISSING SOT SERVICE GOODS VALUE
df_sot.insert(loc=3, column="Sales Channel", value="")
sot_headers = df_sot.columns.values.tolist()
missing_rows = []
for i in df_sot.index:
    row = df_sot.iloc[i]
    docuNo = row["Document No."]
    svcOrderNo = row["Service Order No."]
    nvb_match = df_nvbc.loc[
        (df_nvbc["Document No."] == docuNo)
        & (df_nvbc["Service Order No."] == svcOrderNo)
    ]
    # Add navision status to SOT
    if not nvb_match.empty:
        df_sot.at[i, "Navision Status"] = nvb_match.iloc[0]["Service Status"]
        df_sot.at[i, "Sales Channel"] = nvb_match.iloc[0]["Sales Channel"]
    # print(row)
    # Check for Service Goods Value and port it over
    if pd.isnull(row["Service Goods Value"]):
        if nvb_match.empty:
            missing_rows.append(  # change the row into a list and append to missing_rows
                row.values.tolist()
            )
        else:
            df_sot.at[i, "Service Goods Value"] = nvb_match.iloc[0][
                "Service Goods Value"
            ]
df_missingRow = pd.DataFrame(missing_rows, columns=sot_headers)
print("---SOT SERVICE GOODS VALUE & NAVISION STATUS UPDATED")
print("DATA PREP COMPLETE")
print("_______________________")

"""
PART 2: SOT BASE 
"""
print("PART 2 START")
## Creating sot dataframe with document no and service order no
df_sot_compiled = pd.DataFrame(columns=required_headers)
sot_docuNo = df_sot["Document No."].tolist()
sot_svcOrderNo = df_sot["Service Order No."].tolist()
df_sot_compiled["Document No."] = sot_docuNo
df_sot_compiled["Service Order No."] = sot_svcOrderNo


def data_fill(df):
    for i in df.index:
        row = df.iloc[i]
        docuNo = row["Document No."]
        svcOrderNo = row["Service Order No."]

        # find match
        nvbc_match = df_nvbc.loc[
            (df_nvbc["Document No."] == docuNo)
            & (df_nvbc["Service Order No."] == svcOrderNo)
        ]
        sot_match = df_sot.loc[
            (df_sot["Document No."] == docuNo)
            & (df_sot["Service Order No."] == svcOrderNo)
        ]
        # fill exist location
        if not nvbc_match.empty:
            df.at[i, "NVBC"] = "1"
        if not sot_match.empty:
            df.at[i, "SOT"] = "1"
        if nvbc_match.empty:
            df.at[i, "NVBC"] = "0"
        if sot_match.empty:
            df.at[i, "SOT"] = "0"

        # fill relevant data
        if not nvbc_match.empty and not sot_match.empty:  # match 3
            for col in required_headers:
                if col == "Driver Team":
                    df.at[i, col] = sot_match.iloc[0]["Team"]
                elif col == "Service Name":
                    df.at[i, col] = (
                        str(nvbc_match.iloc[0][col]) + "," + str(sot_match.iloc[0][col])
                    )
                elif col == "Service Date":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Status":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = sot_match.iloc[0][col]
                # elif col == "Service Goods Value1":
                #     df.at[i, col] = sot_match.iloc[0][col]
                # elif col == "Service Goods Value2":
                #     df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Comment":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "CRM Case ID.":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Remark/ Issue":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Status":
                    df.at[i, col] = str(sot_match.iloc[0][col])
                elif col == "Sell-to Customer Name":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to-Address":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Routed Date":
                    df.at[i, col] = nvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    df.at[i, col] = (
                        str(nvbc_match.iloc[0][col])
                        # str(nvbc_match.iloc[0][col]) + "," + str(sot_match.iloc[0][col])
                    )
                # elif col == "Alt Doc Number1 (up to 3)":
                #     df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Sales Channel":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                else:
                    pass
        elif not sot_match.empty and nvbc_match.empty:  # only epod
            for col in required_headers:
                if col == "Driver Team":
                    df.at[i, col] = sot_match.iloc[0]["Team"]
                elif col == "Service Name":
                    df.at[i, col] = str(sot_match.iloc[0][col])
                elif col == "Service Date":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = sot_match.iloc[0][col]
                # elif col == "Service Goods Value1":
                #     df.at[i, col] = sot_match.iloc[0][col]
                # elif col == "Service Goods Value2":
                #     df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Service Comment":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Remark/ Issue":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Status":
                    df.at[i, col] = str(sot_match.iloc[0][col])
                # elif col == "Alt Doc Number (up to 3)":
                #     df.at[i, col] = sot_match.iloc[0][col]
                # elif col == "Alt Doc Number1 (up to 3)":
                #     df.at[i, col] = sot_match.iloc[0][col]
                else:
                    pass


print("---Starting SOT BASE")
data_fill(df_sot_compiled)
print("---SOT BASE Complete")
print("PART 2 COMPLETE")
print("_______________________")

"""
EXCEL EXPORT AND STYLES
"""


def highlight(value):
    if value == "1":
        return "background-color: green"
    elif value == "0":
        return "background-color: red"


# df_epod.to_excel(nvbc_base_writer, sheet_name="epod", index=False)
# df_sot.to_excel(nvbc_base_writer, sheet_name="sot", index=False)
# df_missingRow.to_excel(nvbc_base_writer, sheet_name="sot-missing", index=False)
# df_nvbc.to_excel(nvbc_base_writer, sheet_name="nvbc", index=False)
print("EXPORTING...")
# df_sad.style.applymap(highlight, subset=["EPOD", "SOT", "NVBC"]).to_excel(
#     nvbc_base_writer, sheet_name="A&R Order", index=False
# )
# df_happy.style.applymap(highlight, subset=["EPOD", "SOT", "NVBC"]).to_excel(
#     nvbc_base_writer, sheet_name="S Order", index=False
# )
# nvbc_base_writer.save()
# print("NVBC BASE EXPORTED")

df_sot_compiled.style.applymap(highlight, subset=["SOT", "NVBC"]).to_excel(
    sot_base_writer, sheet_name="SOT BASE", index=False
)
sot_base_writer.save()
print("SOT BASE EXPORTED")
print("_______________________")

