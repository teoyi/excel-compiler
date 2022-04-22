
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

gs_const_svc = [
    "After Sales Delivery",
    "Call Out Service",
    "Delivery Service",
    "Return Delivery",
    "Furniture Removal Service",
]

gs_sgv_svc = [
    "Assembly ASIS",
    "Furniture Assembly Service",
    "Furniture Disassembly Service",
    "Sofa and Sofa Bed Assembly",
]

gs_afterSales = [
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
                    if two != "GS":
                        svcName = row["Service Name"]
                        sgv = row["Service Goods Value"]

                        if three == "MGL":
                            teamNum = team[3:]
                            thirty = ["11", "12", "13", "14"]
                            thirty_three = ["10", "15", "16", "17"]

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
                        elif three == "HJK":
                            if svcName in const_svc:
                                df.at[i, "Payout to Subco"] = 33
                            elif svcName in sgv_svc:
                                df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
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
                    else:
                        svcName = row["Service Name"]
                        sgv = row["Service Goods Value"]

                        if svcName in gs_const_svc:
                            df.at[i, "Payout to Subco"] = 33
                        elif svcName in gs_sgv_svc:
                            df.at[i, "Payout to Subco"] = round(0.072 * sgv, 2)
                        elif svcName in gs_afterSales:
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