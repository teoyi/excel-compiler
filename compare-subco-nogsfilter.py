import pandas as pd
import math
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


"""
FOR SUBCO COMPARE
- based off of document number and service name (maybe address as well)
- primary columns will be at the front, remainder of mvbc compile will be at the back 
"""

required_headers = [
    "Source",
    "GS",
    "SOT",
    "Document No.",
    "GS Service Name",
    "Service Name",
    "Service Order No.",
    "GS Driver Team",
    "SOT Team",
    "GS Service Date",
    "Service Date",
    "GS Service Status",
    "SOT Status",
    "GS Service Goods Value",
    "Service Goods Value",
    "GS Fee Calculation",
    "Payout to Subco",
    "GS Sell-to Customer Name",
    "Sell-to Customer Name",  # NVBC
    "GS Sell-to Address",
    "Sell-to-Address",  # NVBC
    "Capacity Value Weight",
    "Capacity Value Volume",
    "Subco Overweight Status",
    "Ikea Overweight Status",
    "Bill to Ikea",
    "EPOD",
    "MVBC",
    "EPOD Team",  # epod service personne l!!!!!
    "EPOD Reason",  # EPOD reason !!!!!!!
    "EPOD Service Remarks",  # EPOD service remakrs !!!
    "SOT Remark/ Issue",  # SOT remark/issue
    "SOT Service Comment",  # SOT service comment @!!!!!
    "MVBC Service Comment",  # NVBC, SOT
    "EPOD Status",  # EPOD STATUS !!!!!
    "MVBC Service Status",  # NVBC SERVICE STATUS
    "GRN",
    "Manual Value",
    "Service Price Excl. GST",  # NVBC, SOT
    "CRM Case ID.",  # NVBC
    "Service Remarks",  # EPOD
    "Routed Date",  # NVBC service date
    "Alt Doc Number (up to 3)",  # NVBC, SOT
    "Alt Doc Number1 (up to 3)",  # SOT
    "Flow",  # NVBC based on AR or S
    "Sales Channel",  # NVBC
]


# READ EXCEL FILES AND CREATE DATAFRAME
FILE_PATH_SUBCO = "C:\\Users\\yipen\\Desktop\\1st-16th ikea job report.xlsx"
FILE_PATH_SOT = "C:\\Users\\yipen\\Desktop\\NVBC_BASE_COMPILE.xlsx"

df_happy = pd.read_excel(FILE_PATH_SOT, sheet_name="A&R Order")
df_sad = pd.read_excel(FILE_PATH_SOT, sheet_name="S Order")
df_subco = pd.read_excel(FILE_PATH_SUBCO)

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
subco_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\SUBCO_COMPARE_ALL.xlsx", engine="xlsxwriter"
)

"""
Plan: 
1. Get document no and service name from subco df 
2. Create empty dataframe with requried headers
3. Create empty list to contain data 
3. Use document no and service name to find a match in df sad or df happy 
4. append row of subco and append row of sot 
5. combine df and export 
"""

gs_headers = [
    "Document No.",
    "GS Service Name",
    "GS Driver Team",
    "GS Service Date",
    "GS Service Status",
    "GS Service Goods Value",
    "GS Fee Calculation",
    "GS Sell-to Customer Name",
    "GS Sell-to Address",
]

mvbc_headers = [
    "SOT",
    "Document No.",
    "Service Name",
    "Service Order No.",
    "SOT Team",
    "Service Date",
    "SOT Status",
    "Service Goods Value",
    "Payout to Subco",
    "Sell-to Customer Name",  # NVBC
    "Sell-to-Address",  # NVBC
    "Capacity Value Weight",
    "Capacity Value Volume",
    "Subco Overweight Status",
    "Ikea Overweight Status",
    "Bill to Ikea",
    "EPOD",
    "MVBC",
    "EPOD Team",  # epod service personne l!!!!!
    "EPOD Reason",  # EPOD reason !!!!!!!
    "EPOD Service Remarks",  # EPOD service remakrs !!!
    "SOT Remark/ Issue",  # SOT remark/issue
    "SOT Service Comment",  # SOT service comment @!!!!!
    "MVBC Service Comment",  # NVBC, SOT
    "EPOD Status",  # EPOD STATUS !!!!!
    "MVBC Service Status",  # NVBC SERVICE STATUS
    "GRN",
    "Manual Value",
    "Service Price Excl. GST",  # NVBC, SOT
    "CRM Case ID.",  # NVBC
    "Service Remarks",  # EPOD
    "Routed Date",  # NVBC service date
    "Alt Doc Number (up to 3)",  # NVBC, SOT
    "Alt Doc Number1 (up to 3)",  # SOT
    "Flow",  # NVBC based on AR or S
    "Sales Channel",  # NVBC
]

subco_docuNo = df_subco["Document No."].tolist()
subco_svcName = df_subco["GS Service Name"].tolist()
zipped_subco = list(zip(subco_docuNo, subco_svcName))
unique_subco = list(set(list(zipped_subco)))

data = []
df = pd.DataFrame(columns=required_headers)

for i in range(len(unique_subco)):
    printProgressBar(i + 1, len(unique_subco))
    info = unique_subco[i]
    docuNo = info[0]
    svcName = info[1]

    subco_match = df_subco.loc[
        (df_subco["Document No."] == docuNo) & (df_subco["GS Service Name"] == svcName)
    ]

    happy_match = df_happy.loc[
        (df_happy["Document No."] == docuNo) & (df_happy["Service Name"] == svcName)
    ]

    sad_match = df_sad.loc[
        (df_sad["Document No."] == docuNo) & (df_sad["Service Name"] == svcName)
    ]

    if not subco_match.empty:
        subcoMatchCol = subco_match.columns.values.tolist()
        subcoRows = subco_match.values.tolist()
        for subcoRow in subcoRows:
            d = {}
            for col in required_headers:
                d[col] = ""
            d["GS"] = "1"
            d["Source"] = "GS"
            if not happy_match.empty or not sad_match.empty:
                d["SOT"] = "1"
            else:
                d["SOT"] = "0"
            for col in gs_headers:
                i = subcoMatchCol.index(col)
                d[col] = subcoRow[i]
            data.append(d)

    if not happy_match.empty:
        happyMatchCol = happy_match.columns.values.tolist()
        happyRows = happy_match.values.tolist()
        for happyRow in happyRows:
            d = {}
            for col in required_headers:
                d[col] = ""
            d["SOT"] = "1"
            d["Source"] = "SOT"
            if not subco_match.empty:
                d["GS"] = "1"
            else:
                d["GS"] = "0"
            for col in mvbc_headers:
                i = happyMatchCol.index(col)
                d[col] = happyRow[i]
            data.append(d)
            continue

    if not sad_match.empty:
        sadMatchCol = sad_match.columns.values.tolist()
        sadRows = sad_match.values.tolist()
        for sadRow in sadRows:
            d = {}
            for col in required_headers:
                d[col] = ""
            d["SOT"] = "1"
            d["Source"] = "SOT"
            if not subco_match.empty:
                d["GS"] = "1"
            else:
                d["GS"] = "0"
            for col in mvbc_headers:
                i = sadMatchCol.index(col)
                d[col] = sadRow[i]
            data.append(d)


# print(data)
for i in range(len(data)):
    printProgressBar(i + 1, len(data))
    dict = data[i]
    df = df.append(dict, ignore_index=True)
print(df)
# d["GS"] = "1"
# d["Document No."] = docuNo
# d["GS Service Name"]

"""
EXCEL EXPORT AND STYLES
"""


def Repeat(x):
    _size = len(x)
    repeated = []
    for i in range(_size):
        k = i + 1
        for j in range(k, _size):
            if x[i] == x[j] and x[i] not in repeated:
                repeated.append(x[i])
    return repeated


print(Repeat(required_headers))


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
    # elif value == "SUBCO-NONE":
    #     return "background-color: black"
    # elif value == "happy-NONE":
    #     return "background-color: black"
    # elif value == "sad-NONE":
    #     return "background-color: black"


def highlightrow(row):
    if row["Source"] == "GS":
        return ["background-color: #CCCCCC"] * len(row)
    # elif row["Source"] == "SOT":
    #     return ["background-color: #ADD8E6"] * len(row)


print("\nEXPORTING...")
df.style.apply(highlightrow, axis=1).to_excel(subco_writer, index=False)
subco_writer.save()
print("\nDone. Please check the file at the location you have saved.")
