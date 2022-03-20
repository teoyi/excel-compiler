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
    "Document No.",  # NVBC
    "Service Order No.",  # NVBC
    "EPOD",
    "SOT",
    "MVBC",
    "EPOD Team",  # epod service personne l!!!!!
    "SOT Team",  # SOT !!!!!
    "EPOD Reason",  # EPOD reason !!!!!!!
    "EPOD Service Remarks",  # EPOD service remakrs !!!
    "SOT Remark/ Issue",  # SOT remark/issue
    "SOT Service Comment",  # SOT service comment @!!!!!
    "MVBC Service Comment",  # NVBC, SOT
    "EPOD Status",  # EPOD STATUS !!!!!
    "SOT Status",  # SOT STATUS !!!!!!!!
    "MVBC Service Status",  # NVBC SERVICE STATUS
    "Service Name",  # NVBC, SOT, EPOD
    "Service Date",  # NVBC, SOT
    "GRN",
    "Service Goods Value",  # NVBC, SOT
    "Service Goods Value1",  # SOT
    "Service Goods Value2",  # SOT
    "Manual Value",
    "Service Price Excl. GST",  # NVBC, SOT
    "Payout to Subco",
    "Bill to Ikea",
    "CRM Case ID.",  # NVBC
    "Service Remarks",  # EPOD
    "Sell-to Customer Name",  # NVBC
    "Sell-to-Address",  # NVBC
    "Routed Date",  # NVBC service date
    "Alt Doc Number (up to 3)",  # NVBC, SOT
    "Alt Doc Number1 (up to 3)",  # SOT
    "Flow",  # NVBC based on AR or S
    "Sales Channel",  # NVBC
]

# READ EXCEL FILES AND CREATE DATAFRAME
FILE_PATH_EPOD = "C:\\Users\\yipen\\Desktop\\JAN-EPOD-COPY.xls"
FILE_PATH_SOT = "C:\\Users\\yipen\\Desktop\\compiled_JAN_SOT_2022.xlsx"
FILE_PATH_NVBC = "C:\\Users\\yipen\\Desktop\\MVBC_JAN.xlsx"

df_epod = pd.read_excel(FILE_PATH_EPOD)
df_sot = pd.read_excel(FILE_PATH_SOT)
df_nvbc = pd.read_excel(FILE_PATH_NVBC, sheet_name="Filtered")

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
nvbc_base_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\NVBC_BASE_COMPILE.xlsx", engine="xlsxwriter"
)

"""
DATA PREP
1. Trim EPOD Team col
2. Replace SOT Service Goods Value
"""
print("_______________________")
print("DATA PREP START")
# TRIM TEAM COLUMN FOR EPOD
df_epod["EPOD Team"] = df_epod["EPOD Team"].str.replace("MGL_", "")
df_epod["EPOD Team"] = df_epod["EPOD Team"].str.replace("MGL-", "")
print("---EPOD TEAM TRIM COMPLETE")

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
        df_sot.at[i, "Navision Status"] = nvb_match.iloc[0]["MVBC Service Status"]
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
PART 1: NVBC BASE 

0. Create Compiled dataframe with headers 
1. Get all Document No. and Service Order No. from NVBC df
2. Turn Document No. and Service Order No. into arrays and zip to create relation 
3. Create empty list 
4. Iterate over  

Possible matches
(EVERY SHEET WILL CONTINUE TO CONTAIN ALL NVBC ENTRY ) 
NVBC SOT EPOD 
1. All 3 
2. NVBC + SOT 
3. NVBC + EPOD 
4. SOT + EPOD 
5. SOT ONLY 
6. EPOD ONLY
"""
print("PART 1 START")

# create sad and happy dataframe
df_sad = pd.DataFrame(columns=required_headers)
df_happy = pd.DataFrame(columns=required_headers)

# extract document no and service order no
nvbc_docuNo = df_nvbc["Document No."].tolist()
nvbc_svcOrderNo = df_nvbc["Service Order No."].tolist()
sad_docuNo = []
sad_svcOrderNo = []
happy_docuNo = []
happy_svcOrderNo = []

if len(nvbc_docuNo) == len(nvbc_svcOrderNo):
    for i in range(len(nvbc_docuNo)):
        docuNo = nvbc_docuNo[i]
        svcOrderNo = nvbc_svcOrderNo[i]
        if docuNo[0] == "A" or docuNo[0] == "R":
            sad_docuNo.append(docuNo)
            sad_svcOrderNo.append(svcOrderNo)
        elif docuNo[0] == "S":
            happy_docuNo.append(docuNo)
            happy_svcOrderNo.append(svcOrderNo)
else:
    print(
        f"ERROR: NVBC Document No. length [{len(nvbc_docuNo)}] does not match Service Order No. Length [{len(nvbc_svcOrderNo)}]"
    )

# assign list values to dataframe
df_sad["Document No."] = sad_docuNo
df_sad["Service Order No."] = sad_svcOrderNo
df_happy["Document No."] = happy_docuNo
df_happy["Service Order No."] = happy_svcOrderNo


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
        epod_match = df_epod.loc[
            (df_epod["Document No."] == docuNo)
            & (df_epod["Service Order No."] == svcOrderNo)
        ]
        # fill exist location
        if not nvbc_match.empty:
            df.at[i, "MVBC"] = "1"
        if not sot_match.empty:
            df.at[i, "SOT"] = "1"
        if not epod_match.empty:
            df.at[i, "EPOD"] = "1"
        if sot_match.empty:
            df.at[i, "SOT"] = "0"
        if epod_match.empty:
            df.at[i, "EPOD"] = "0"

        # fill relevant data
        if (
            not nvbc_match.empty and not sot_match.empty and not epod_match.empty
        ):  # match 3
            for col in required_headers:
                if col == "SOT Team":
                    df.at[i, col] = sot_match.iloc[0]["SOT Team"]
                elif col == "EPOD Team":
                    df.at[i, col] = epod_match.iloc[0]["EPOD Team"]
                elif col == "EPOD Reason":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "EPOD Service Remarks":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "SOT Remark/Issue":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "SOT Service Comment":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "MVBC Service Comment":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "EPOD Status":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "SOT Status":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "MVBC Service Status":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Name":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Service Date":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "CRM Case ID.":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to Customer Name":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to-Address":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Routed Date":
                    df.at[i, col] = nvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    df.at[i, col] = (
                        str(nvbc_match.iloc[0][col]) + "," + str(sot_match.iloc[0][col])
                    )
                elif col == "Alt Doc Number1 (up to 3)":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Sales Channel":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                else:
                    pass
        elif (
            not nvbc_match.empty and not sot_match.empty and epod_match.empty
        ):  # only sot
            for col in required_headers:
                if col == "SOT Team":
                    df.at[i, col] = sot_match.iloc[0]["SOT Team"]
                elif col == "SOT Remark/Issue":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "SOT Service Comment":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "MVBC Service Comment":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "SOT Status":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "MVBC Service Status":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Name":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Service Date":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "CRM Case ID.":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to Customer Name":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to-Address":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Routed Date":
                    df.at[i, col] = nvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    df.at[i, col] = (
                        str(nvbc_match.iloc[0][col]) + "," + str(sot_match.iloc[0][col])
                    )
                elif col == "Alt Doc Number1 (up to 3)":
                    df.at[i, col] = sot_match.iloc[0][col]
                elif col == "Sales Channel":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                else:
                    pass
        elif (
            not nvbc_match.empty and sot_match.empty and not epod_match.empty
        ):  # only epod
            for col in required_headers:
                if col == "EPOD Team":
                    df.at[i, col] = epod_match.iloc[0]["EPOD Team"]
                elif col == "EPOD Reason":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "EPOD Service Remarks":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "MVBC Service Comment":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "EPOD Status":
                    df.at[i, col] = epod_match.iloc[0][col]
                elif col == "MVBC Service Status":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Name":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Service Date":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "CRM Case ID.":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to Customer Name":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to-Address":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Routed Date":
                    df.at[i, col] = nvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Sales Channel":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                else:
                    pass
        elif not nvbc_match.empty and sot_match.empty and epod_match.empty:  # only nvbc
            for col in required_headers:
                if col == "MVBC Service Comment":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "MVBC Service Status":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Name":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Service Date":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Goods Value":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Service Price Excl. GST":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "CRM Case ID.":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to Customer Name":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Sell-to-Address":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                elif col == "Routed Date":
                    df.at[i, col] = nvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    df.at[i, col] = str(nvbc_match.iloc[0][col])
                elif col == "Sales Channel":
                    df.at[i, col] = nvbc_match.iloc[0][col]
                else:
                    pass


print("---Starting Sad")
data_fill(df_sad)
print("---Sad Complete")
print("---Starting Happy")
data_fill(df_happy)
print("---Happy Complete")
print("PART 1 COMPLETE")
print("_______________________")


"""
EXCEL EXPORT AND STYLES
"""


def highlight(value):
    if value == "1":
        return "background-color: #90EE90"
    elif value == "0":
        return "background-color: #FFCCCB"


# df_epod.to_excel(nvbc_base_writer, sheet_name="epod", index=False)
# df_sot.to_excel(nvbc_base_writer, sheet_name="sot", index=False)
# df_missingRow.to_excel(nvbc_base_writer, sheet_name="sot-missing", index=False)
# df_nvbc.to_excel(nvbc_base_writer, sheet_name="nvbc", index=False)
print("EXPORTING...")
df_sad.style.applymap(highlight, subset=["EPOD", "SOT", "MVBC"]).to_excel(
    nvbc_base_writer, sheet_name="A&R Order", index=False
)
df_happy.style.applymap(highlight, subset=["EPOD", "SOT", "MVBC"]).to_excel(
    nvbc_base_writer, sheet_name="S Order", index=False
)
nvbc_base_writer.save()
print("NVBC BASE EXPORTED")

print("_______________________")

