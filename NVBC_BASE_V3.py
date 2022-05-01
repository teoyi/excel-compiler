import pandas as pd
from cowpy import cow

msg = cow.milk_random_cow("Welcome to your data processing script!")
print(msg)

# import progress bar
# Print iterations progress
def printProgressBar(
    iteration,
    total,
    prefix="",
    suffix="",
    decimals=1,
    length=50,
    fill="█",
    printEnd="",
):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + "-" * (length - filledLength)
    print(f"{prefix} |{bar}| {percent}% {suffix}\r", end="", flush=True)
    # Print New Line on Complete
    if iteration == total:
        print()


"""'
    Headers from revised data
    EPOD NVBC SOT 

    SAD -> A + R document no. 
    HAPPY -> S document no. 

    -> NVBC Document No. + Service Order No. as ref 
    -> match with SOT first to get all possible team numbers 
    -> then match with EPOD to fill remaining data

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
    "Manual Value",
    "Service Price Excl. GST",  # NVBC, SOT
    "Capacity Value Weight",
    "Capacity Value Volume",
    "Subco Overweight Status",
    "Ikea Overweight Status",
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
FILE_PATH_EPOD = "C:\\Users\\yipen\\Desktop\\EPOD - MAR 22 (18042022).xls"
FILE_PATH_SOT = "C:\\Users\\yipen\\Desktop\\filtered_ALL_MAR2022.xlsx"
FILE_PATH_NVBC = "C:\\Users\\yipen\\Desktop\\Navition March 2022 (18042022).xlsx"

df_epod = pd.read_excel(FILE_PATH_EPOD)
df_sot = pd.read_excel(FILE_PATH_SOT, sheet_name="filtered")
df_sot_dropped = pd.read_excel(FILE_PATH_SOT, sheet_name="dropped")
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
print("\nCleaning Original Data...")
# TRIM TEAM COLUMN FOR EPOD
df_epod["EPOD Team"] = df_epod["EPOD Team"].str.replace("MGL_", "")
df_epod["EPOD Team"] = df_epod["EPOD Team"].str.replace("MGL-", "")


# REPLACE MISSING SOT SERVICE GOODS VALUE
df_sot.insert(loc=3, column="Sales Channel", value="")
sot_headers = df_sot.columns.values.tolist()
missing_rows = []
for i in df_sot.index:
    printProgressBar(i + 1, len(df_sot))
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

"""
PART 1: MVBC BASE 

1. Take from MVBC the unique set of document no. and service order no. turn into list to iterate (index values are the same so can iterate over length)
2. Create compiled dataframe for appending relevant data 
3. Iterate over document no. and service order no. and get the sot match. 
    For each match: 
        iterate over headers, if headers match return value else return null 
4. From that match epod based on document no. service order no. and team

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

# extract document no and service order no
nvbc_docuNo = df_nvbc["Document No."].tolist()
nvbc_svcOrderNo = df_nvbc["Service Order No."].tolist()
zipped_mvbc = list(zip(nvbc_docuNo, nvbc_svcOrderNo))
unique_mvbc = list(set(list(zipped_mvbc)))
# mvbc tends not to have duplicates so this is just to ensure any outliers are taken care of

sad_info = []
happy_info = []

print("\nSorting Happy and Sad...")
for i in range(len(unique_mvbc)):  # iterating over unique docuno and svcorderno set
    printProgressBar(i + 1, len(unique_mvbc))
    if unique_mvbc[i][0][0] == "A" or unique_mvbc[i][0][0] == "R":
        sad_info.append(unique_mvbc[i])
    elif unique_mvbc[i][0][0] == "S":
        happy_info.append(unique_mvbc[i])

# create sad and happy dataframe
# df_sad = pd.DataFrame(columns=required_headers)
# df_happy = pd.DataFrame(columns=required_headers)

# with document no and service order no, iterate over and use the document no and service order no to match mvbc and sot first

mvbc_col = [
    "MVBC Service Comment",
    "MVBC Service Status",
    "Service Name",
    "Service Date",
    "Service Goods Value",
    "Service Price Excl. GST",
    "CRM Case ID.",
    "Sell-to Customer Name",
    "Sell-to-Address",
    "Routed Date",
    "Alt Doc Number (up to 3)",
    "Sales Channel",
    "Capacity Value Weight",
    "Capacity Value Volume",
]

sot_col = [
    "SOT Team",
    "SOT Remark/ Issue",
    "SOT Service Comment",
    "SOT Status",
    "Alt Doc Number (up to 3)",
    "Alt Doc Number1 (up to 3)",
]

epod_col = ["EPOD Team", "EPOD Reason", "EPOD Service Remarks", "EPOD Status"]

sotMatchedPairs = []


def sot_fill(info):
    df = pd.DataFrame(columns=required_headers)
    data = []
    print("------ Matching MVBC and SOC...")
    for i in range(len(info)):
        printProgressBar(i + 1, len(info))
        pairs = info[i]
        # print(pairs)
        # Based on MVBC
        # [0] is Document No.
        # [1] is Service Order No.
        docuNo = pairs[0]
        svcOrderNo = pairs[1]

        mvbc_match = df_nvbc.loc[
            (df_nvbc["Document No."] == docuNo)
            & (df_nvbc["Service Order No."] == svcOrderNo)
        ]

        sot_match = df_sot.loc[
            (df_sot["Document No."] == docuNo)
            & (df_sot["Service Order No."] == svcOrderNo)
        ]

        if sot_match.empty:  # no match
            # initializing dictionary to contain row values
            d = {}
            for i in required_headers:
                d[i] = None

            d["SOT"] = "0"
            d["MVBC"] = "1"
            d["Document No."] = docuNo
            d["Service Order No."] = svcOrderNo

            for col in mvbc_col:
                # check if value exissts ,if empty just go next col
                if col == "Routed Date":
                    if pd.isnull(mvbc_match.iloc[0]["Service Date"]):
                        continue
                else:
                    if pd.isnull(mvbc_match.iloc[0][col]):
                        continue

                if col == "Routed Date":
                    d[col] = mvbc_match.iloc[0]["Service Date"]
                elif col == "Service Name" or col == "Alt Doc Number (up to 3)":
                    d[col] = str(mvbc_match.iloc[0][col])
                else:
                    d[col] = mvbc_match.iloc[0][col]

            # append dictionary to dataframe
            data.append(d)

        elif len(sot_match) > 1:  # match more than once
            # append to list of matched sot
            sotMatchedPairs.append(pairs)
            sotMatchCol = sot_match.columns.values.tolist()
            sotRows = sot_match.values.tolist()
            for sotRow in sotRows:
                # initializing dictionary to contain row values
                d = {}
                for i in required_headers:
                    d[i] = None

                d["SOT"] = "1"
                d["MVBC"] = "1"
                d["Document No."] = docuNo
                d["Service Order No."] = svcOrderNo

                for col in mvbc_col:
                    # check if value exissts ,if empty just go next col
                    if col == "Routed Date":
                        if pd.isnull(mvbc_match.iloc[0]["Service Date"]):
                            continue
                    else:
                        if pd.isnull(mvbc_match.iloc[0][col]):
                            continue

                    if col == "Routed Date":
                        d[col] = mvbc_match.iloc[0]["Service Date"]
                    elif col == "Service Name" or col == "Alt Doc Number (up to 3)":
                        d[col] = str(mvbc_match.iloc[0][col])
                    else:
                        d[col] = mvbc_match.iloc[0][col]

                for col in sot_col:
                    # get index of column from the matched col
                    i = sotMatchCol.index(col)
                    if col == "Alt Doc Number (up to 3)":
                        d[col] = str(d[col]) + "," + str(sotRow[i])
                    else:
                        d[col] = sotRow[i]

                # append dictionary to dataframe
                data.append(d)

        elif len(sot_match) == 1:  # match only once
            # Adding matched pairs to list
            sotMatchedPairs.append(pairs)
            # initializing dictionary to contain row values
            d = {}
            for i in required_headers:
                d[i] = None

            d["SOT"] = "1"
            d["MVBC"] = "1"
            d["Document No."] = docuNo
            d["Service Order No."] = svcOrderNo

            for col in mvbc_col:
                # check if value exissts ,if empty just go next col
                if col == "Routed Date":
                    if pd.isnull(mvbc_match.iloc[0]["Service Date"]):
                        continue
                else:
                    if pd.isnull(mvbc_match.iloc[0][col]):
                        continue

                if col == "Routed Date":
                    d[col] = mvbc_match.iloc[0]["Service Date"]
                elif col == "Alt Doc Number (up to 3)":
                    d[col] = str(mvbc_match.iloc[0][col])
                else:
                    d[col] = mvbc_match.iloc[0][col]

            for col in sot_col:
                if col == "Alt Doc Number (up to 3)":
                    d[col] = str(d[col]) + "," + str(sot_match.iloc[0][col])
                else:
                    d[col] = sot_match.iloc[0][col]

            # append dictionary to dataframe
            data.append(d)
    print("------Finalizing SOT Data...")
    for i in range(len(data)):
        printProgressBar(i + 1, len(data))
        dict = data[i]
        df = df.append(dict, ignore_index=True)
    return df


print("\nStart SOT fill")
print("--- Filling Happy SOT")
df_happy_sot = sot_fill(happy_info)
print("--- Filling Sad SOT")
df_sad_sot = sot_fill(sad_info)


def epod_fill(df):
    dfList = df.to_dict(orient="records")
    data = []
    print("------ Matching EPOD...")
    for i in range(len(dfList)):
        printProgressBar(i + 1, len(dfList))
        d = dfList[i]
        docuNo = d["Document No."]
        svcOrderNo = d["Service Order No."]
        epod_match = df_epod.loc[
            (df_epod["Document No."] == docuNo)
            & (df_epod["Service Order No."] == svcOrderNo)
        ]
        if not epod_match.empty:
            d["EPOD"] = "1"
            for col in epod_col:
                d[col] = epod_match.iloc[0][col]
            data.append(d)
        else:
            d["EPOD"] = "0"
            data.append(d)

    df = pd.DataFrame(columns=required_headers)
    print("------ Finalizing Data...")
    for i in range(len(data)):
        printProgressBar(i + 1, len(data))
        dict = data[i]
        df = df.append(dict, ignore_index=True)
    return df


print("\nStarting EPOD fill")
print("--- Filling Happy EPOD")
df_happy_complete = epod_fill(df_happy_sot)
print("--- Filling Sad EPOD")
df_sad_complete = epod_fill(df_sad_sot)

"""
PART 2: CALCULATE PAYMENTS 
"""
required_svcNames = [
    "After Sales Assembly",
    "After Sales Delivery",
    "After Sales Disassembly",
    "Assembly ASIS",
    "Call Out Service",
    "Delivery Service",
    "Furniture Assembly Service",
    "Furniture Disassembly Service",
    "Furniture Removal Service",
    "Return Delivery",
    "Sofa and Sofa Bed Assembly",
]

const_svc = [
    "After Sales Delivery",
    "Call Out Service",
    "Delivery Service",
    "Return Delivery",
    "Furniture Removal Service",
]

sgv_svc = [
    "After Sales Assembly",
    "After Sales Disassembly",
    "Assembly ASIS",
    "Furniture Assembly Service",
    "Furniture Disassembly Service",
    "Sofa and Sofa Bed Assembly",
]

###
# Bill to ikea
###
def bill_to_ikea(df):
    print("----- Calculating...")
    for i in df.index:
        printProgressBar(i + 1, len(df))
        row = df.iloc[i]
        svcName = row["Service Name"]
        sgv = row["Service Goods Value"]
        mvbc_val = row["MVBC"]
        sot_val = row["SOT"]
        epod_val = row["EPOD"]

        if (mvbc_val == "1") and (sot_val == "0") and (epod_val == "0"):
            df.at[i, "Bill to Ikea"] = 0
        else:
            if svcName in const_svc:
                df.at[i, "Bill to Ikea"] = 35
            elif svcName in sgv_svc:
                df.at[i, "Bill to Ikea"] = round(0.095 * sgv, 2)


print("\nStarting Bill to Ikea Calculation")
print("--- For Happy")
bill_to_ikea(df_happy_complete)
print("--- For Sad")
bill_to_ikea(df_sad_complete)

###
# Payout to subco
###

contracted_const_svc = [
    "After Sales Delivery",
    "Call Out Service",
    "Delivery Service",
    "Return Delivery",
    "Furniture Removal Service",
]

contracted_sgv_svc = [
    "Assembly ASIS",
    "Furniture Assembly Service",
    "Furniture Disassembly Service",
    "Sofa and Sofa Bed Assembly",
]

contracted_afterSales = [
    "After Sales Assembly",
    "After Sales Disassembly",
]

df_h_outlier = []
df_s_outlier = []

#  for h
def pay_to_sub(df, df_outlier):
    print("------ Calculating...")
    for i in df.index:
        printProgressBar(i + 1, len(df))
        row = df.iloc[i]
        team = row["SOT Team"]
        mvbc_val = row["MVBC"]
        sot_val = row["SOT"]
        epod_val = row["EPOD"]
        if (mvbc_val == "1") and (sot_val == "0") and (epod_val == "0"):
            df.at[i, "Payout to Subco"] = 0
        else:
            if not pd.isnull(team):
                strLength = len(team)
                if strLength > 5:
                    df_outlier.append(row.values.tolist())
                else:
                    one = team[0]
                    two = team[:2]
                    three = team[:3]

                    # if three == "MGL":
                    #     print(team[3:])

                    # # print(three, two, one)
                    if two != "GS" or three != "HJK":
                        svcName = row["Service Name"]
                        sgv = row["Service Goods Value"]

                        if three == "MGL":
                            teamNum = team[3:]
                            thirty = ["11", "12", "13", "14"]
                            thirty_three = ["10", "15", "16", "17", "20"]

                            if teamNum in thirty:
                                if svcName in const_svc:
                                    df.at[i, "Payout to Subco"] = 30
                                elif svcName in sgv_svc:
                                    df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                            elif teamNum in thirty_three:
                                if svcName in const_svc:
                                    df.at[i, "Payout to Subco"] = 33
                                elif svcName in sgv_svc:
                                    df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                            else:
                                df.at[i, "Payout to Subco"] = 0
                        # elif three == "HJK":
                        #     if svcName in const_svc:
                        #         df.at[i, "Payout to Subco"] = 33
                        #     elif svcName in sgv_svc:
                        #         df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif two == "JX":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif two == "SL":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif three == "TLI":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif one == "T":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif three == "SGP":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.07 * sgv, 2)
                        elif two == "BF":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif two == "KR":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif two == "N1":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.07 * sgv, 2)
                        elif two == "IK":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif two == "PN":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 30
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                    else:
                        svcName = row["Service Name"]
                        sgv = row["Service Goods Value"]

                        if svcName in contracted_const_svc:
                            df.at[i, "Payout to Subco"] = 33
                        elif svcName in contracted_sgv_svc:
                            df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif svcName in contracted_afterSales:
                            df.at[i, "Payout to Subco"] = round(0.07 * sgv, 2)


print("\nStarting Payout to Subco Calculation")
print("--- For Happy")
pay_to_sub(df_happy_complete, df_h_outlier)
print("--- For Sad")
pay_to_sub(df_sad_complete, df_s_outlier)

outlier_h = pd.DataFrame(df_h_outlier, columns=required_headers)
outlier_s = pd.DataFrame(df_s_outlier, columns=required_headers)

"""
PART 3: OVERWEIGHT CALCULATION 
"""

###
# IKEA OVERWEIGHT STATUS
###

overweight_services = [
    "Call Out Service",
    "Delivery Service",
    "Return Delivery",
    "Furniture Removal Service",
    "After Sales Delivery",
]


def ikea_overweight(df):
    df["Ikea Overweight Status"] = df["Ikea Overweight Status"].astype("str")
    print("------ Calculating...")
    for i in df.index:
        printProgressBar(i + 1, len(df))
        row = df.iloc[i]
        cvw = row["Capacity Value Weight"]  # cvw = Capacity Value Weight
        bti = row["Bill to Ikea"]
        mvbc_val = row["MVBC"]
        sot_val = row["SOT"]
        epod_val = row["EPOD"]
        svcName = row["Service Name"]
        if (mvbc_val == "1") and (sot_val == "0") and (epod_val == "0"):
            df.at[i, "Ikea Overweight Status"] = "N/A"
        else:
            if svcName in overweight_services:
                if 600 <= cvw <= 800:
                    df.at[i, "Ikea Overweight Status"] = "600-800 OVERWEIGHT"
                    if not pd.isnull(bti):
                        df.at[i, "Bill to Ikea"] = 180
                elif 801 <= cvw:
                    df.at[i, "Ikea Overweight Status"] = ">801 OVERWEIGHT"
                    if not pd.isnull(bti):
                        df.at[i, "Bill to Ikea"] = 260
                else:
                    df.at[i, "Ikea Overweight Status"] = "N/A"
            else:
                df.at[i, "Ikea Overweight Status"] = "N/A"


print("\nStarting Ikea Overweight Calculations")
print("--- For Happy Main")
ikea_overweight(df_happy_complete)
print("--- For Happy Outlier")
ikea_overweight(outlier_h)
print("--- For Sad Main")
ikea_overweight(df_sad_complete)
print("--- For Sad Outlier")
ikea_overweight(outlier_s)


def subco_overweight(df):
    df["Subco Overweight Status"] = df["Subco Overweight Status"].astype("str")
    print("------ Calculating...")
    for i in df.index:
        printProgressBar(i + 1, len(df))
        row = df.iloc[i]
        cvw = row["Capacity Value Weight"]  # cvw = Capacity Value Weight
        pos = row["Payout to Subco"]
        mvbc_val = row["MVBC"]
        sot_val = row["SOT"]
        epod_val = row["EPOD"]
        svcName = row["Service Name"]
        if (mvbc_val == "1") and (sot_val == "0") and (epod_val == "0"):
            df.at[i, "Subco Overweight Status"] = "N/A"
        else:
            if svcName in overweight_services:
                if 600 <= cvw <= 800:
                    df.at[i, "Subco Overweight Status"] = "600-800 OVERWEIGHT"
                    if not pd.isnull(pos):
                        df.at[i, "Payout to Subco"] = 70
                elif 801 <= cvw <= 1000:
                    df.at[i, "Subco Overweight Status"] = "801-1000 OVERWEIGHT"
                    if not pd.isnull(pos):
                        df.at[i, "Payout to Subco"] = 90
                elif 1001 <= cvw <= 2000:
                    df.at[i, "Subco Overweight Status"] = "1001-2000 OVERWEIGHT"
                    if not pd.isnull(pos):
                        df.at[i, "Payout to Subco"] = 120
                elif 2001 <= cvw:
                    df.at[i, "Subco Overweight Status"] = ">2001 OVERWEIGHT"
                    if not pd.isnull(pos):
                        df.at[i, "Payout to Subco"] = 150
                else:
                    df.at[i, "Subco Overweight Status"] = "N/A"
            else:
                df.at[i, "Subco Overweight Status"] = "N/A"


print("Starting Subco Overweight Calculations")
print("--- For Happy Main")
subco_overweight(df_happy_complete)
print("--- For Happy Outlier")
subco_overweight(outlier_h)
print("--- For Sad Main")
subco_overweight(df_sad_complete)
print("--- For Sad Outlier")
subco_overweight(outlier_s)

"""
Part 4: Getting unmatched values for SOT 
"""

# List containing docu and svc order no are in sotMatchedPairs
# print(sotMatchedPairs)

# create list of index to be dropped later
dropList = []
## iterate over matched pairs and get index of rows that meets the requirement
for pair in sotMatchedPairs:
    docuNo = pair[0]
    svcOrderNo = pair[1]
    indices = df_sot.index[
        (df_sot["Document No."] == docuNo) & (df_sot["Service Order No."] == svcOrderNo)
    ]
    ## append index to list
    for index in indices:
        dropList.append(index)
# print(dropList)
## drop the rows of index at the end of it
df_unmatch_sot = df_sot.drop(dropList, 0)

"""
EXCEL EXPORT AND STYLES
"""


def highlight(value):
    if value == "1":
        return "background-color: #90EE90"
    elif value == "0":
        return "background-color: #FFCCCB"
    elif value == "600-800 OVERWEIGHT":
        return "background-color: #FFCCCB"
    elif value == "801-1000 OVERWEIGHT":
        return "background-color: #FFCCCB"
    elif value == "1001-2000 OVERWEIGHT":
        return "background-color: #FFCCCB"
    elif value == ">2001 OVERWEIGHT":
        return "background-color: #FFCCCB"
    elif value == ">801 OVERWEIGHT":
        return "background-color: #FFCCCB"
    elif value == "N/A":
        return "background-color: #ADD8E6"


# df_epod.to_excel(nvbc_base_writer, sheet_name="epod", index=False)
# df_sot.to_excel(nvbc_base_writer, sheet_name="sot", index=False)
# df_missingRow.to_excel(nvbc_base_writer, sheet_name="sot-missing", index=False)
# df_nvbc.to_excel(nvbc_base_writer, sheet_name="nvbc", index=False)
print("\nEXPORTING...")
df_sad_complete.style.applymap(
    highlight,
    subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
).to_excel(nvbc_base_writer, sheet_name="A&R Order", index=False)
df_happy_complete.style.applymap(
    highlight,
    subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
).to_excel(nvbc_base_writer, sheet_name="S Order", index=False)
outlier_h.style.applymap(
    highlight,
    subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
).to_excel(nvbc_base_writer, sheet_name="S OUTLIER Order", index=False)
outlier_s.style.applymap(
    highlight,
    subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
).to_excel(nvbc_base_writer, sheet_name="A&R OUTLIER Order", index=False)
df_unmatch_sot.to_excel(nvbc_base_writer, sheet_name="Unmatch SOT", index=False)
nvbc_base_writer.save()
print("\nDone. Please check the file at the location you have saved.")
