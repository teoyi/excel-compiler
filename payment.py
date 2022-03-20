import pandas as pd

FILE_PATH_MVBC_BASE = "C:\\Users\\yipen\\Desktop\\NVBC_BASE_COMPILE.xlsx"

df_h = pd.read_excel(FILE_PATH_MVBC_BASE, sheet_name="A&R Order")
df_s = pd.read_excel(FILE_PATH_MVBC_BASE, sheet_name="S Order")

# print(df)

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

"""
After Sales Assembly
After Sales Delivery
After Sales Disassembly
Assembly ASIS
Call Out Service
Delivery Service
Furniture Assembly Service
Furniture Disassembly Service
Furniture Removal Service
Return Delivery
Sofa and Sofa Bed Assembly
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


"""
BILL TO IKEA 
35 || 9.5 
"""


def bill_to_ikea(df):
    for i in df.index:
        row = df.iloc[i]
        # print(row)
        svcName = row["Service Name"]
        sgv = row["Service Goods Value"]
        if svcName in const_svc:
            df.at[i, "Bill to Ikea"] = 35
        elif svcName in sgv_svc:
            df.at[i, "Bill to Ikea"] = round(0.095 * sgv, 2)


bill_to_ikea(df_h)
bill_to_ikea(df_s)

"""
PAY TO SUBCON 
MGL 11 12 13 -> $30 || 7.2% * SGV
MGL 16 17 -> $33 || " 
MGL 14 -> $30 || " 
MGL 10 -> $33 || "
MGL 15 -> $33 || " 
HJK -> " || " 
JX -> "
SL -> "
T -> " 
SGP -> $33 || 7% * SGV 
BF -> $33 || 7.2% * SGV 
KR -> $33 || 7.2% 
TLI -> 33 || 7.2 
N1 -> 33 || 7 
GS -> 33 || 7.2 (ASSEMBLY) || 7 (AFTER SALES ASSEMBLY + DISASSEMBLY)
IK -> 33 || 7.2 
"""
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

# df_h_outlier = pd.DataFrame(columns=required_headers)
# df_s_outlier = pd.DataFrame(columns=required_headers)
df_h_outlier = []
df_s_outlier = []

#  for h
def pay_to_sub(df, df_outlier):
    for i in df.index:
        row = df.iloc[i]
        team = row["SOT Team"]

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
                        thirty = [11, 12, 13, 14]
                        thirty_three = [10, 15, 16, 17]
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


pay_to_sub(df_h, df_h_outlier)
pay_to_sub(df_s, df_s_outlier)

outlier_h = pd.DataFrame(df_h_outlier, columns=required_headers)
outlier_s = pd.DataFrame(df_s_outlier, columns=required_headers)

nvbc_base_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\test_billpay.xlsx", engine="xlsxwriter"
)
df_h.to_excel(nvbc_base_writer, sheet_name="S Order", index=False)
outlier_h.to_excel(nvbc_base_writer, sheet_name="S Outlier Order", index=False)
df_s.to_excel(nvbc_base_writer, sheet_name="A&R Order", index=False)
outlier_s.to_excel(nvbc_base_writer, sheet_name="A&R Outlier Order", index=False)
nvbc_base_writer.save()

