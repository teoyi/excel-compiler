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


# READ EXCEL FILES AND CREATE DATAFRAME
FILE_PATH_MINE = "C:\\Users\\yipen\\Desktop\\JOB LIST 17 - 30 APR.xlsx"
FILE_PATH_SUBCO = (
    "C:\\Users\\yipen\\Desktop\\GS Billing for 17-30 Apr 2022 (from GS).xlsx"
)

df_mine = pd.read_excel(FILE_PATH_MINE, sheet_name="17 - 30 APR")
df_subco = pd.read_excel(FILE_PATH_SUBCO, sheet_name="Fee calculation")

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
compare_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\SUBCO_COMPARE.xlsx", engine="xlsxwriter"
)

print(df_mine)
print(df_subco)

"""
iterate over mine df 
use document no. and service name for matching 

create 2 df
1) matched with subco -> return same row with svs goods value and fee calculation filled out
2) no match with subco 
"""

# get column names and convert into header list
headers_mine = df_mine.columns.tolist()
headers_subco = df_subco.columns.tolist()

ori = []
unmatched_data = []
multi_data = []

for i in df_mine.index:
    printProgressBar(i + 1, len(df_mine))
    row = df_mine.iloc[i]
    docuNo = row["DocumentNo"]
    svcName = row["ServiceName"]
    team = row["Team"]

    s_match = df_subco.loc[
        (df_subco["Document No."] == docuNo)
        & (df_subco["Service Name"] == svcName)
        & (df_subco["Driver Team"] == team)
    ]

    if (
        s_match.empty
    ):  # add mine rows to matched data and filling up the gs realted columsn to missing
        d = {}  # initialize dictionary
        for i in headers_mine:  # create empty dictionary
            d[i] = ""
        # unmatched_data.append(s_match)
        for col in headers_mine:
            if col == "GS svs good value" or col == "GS Fee calculation":
                d[col] = "MISSING"
            else:
                d[col] = row[col]
        ori.append(d)
    elif len(s_match.index) > 1:
        # print(s_match)
        mine = {}  # continue to put into ori but leave gs related as check multi
        for i in headers_mine:
            mine[i] = ""
        for col in headers_mine:
            if col == "GS svs good value" or col == "GS Fee calculation":
                mine[col] = "CHECK MULTI"
            else:
                mine[col] = row[col]
        ori.append(mine)

        multiRows = s_match.values.tolist()
        for multiRow in multiRows:
            # initializing dictionary to contain row values
            d = {}
            for i in headers_subco:
                d[i] = ""
            for col in headers_subco:
                i = headers_subco.index(col)
                d[col] = multiRow[i]

            multi_data.append(d)
    else:
        d = {}  # initialize dictionary
        for i in headers_mine:  # create empty dictionary
            d[i] = ""
        # unmatched_data.append(s_match)
        for col in headers_mine:
            if col == "GS svs good value":
                d[col] = s_match.iloc[0]["Service Good Value"]
            elif col == "GS Fee calculation":
                d[col] = s_match.iloc[0]["Fee Calculation"]
            else:
                d[col] = row[col]
        ori.append(d)

df_ori = pd.DataFrame(columns=headers_mine)
for i in range(len(ori)):
    printProgressBar(i + 1, len(ori))
    dict = ori[i]
    df_ori = df_ori.append(dict, ignore_index=True)


df_multi = pd.DataFrame(columns=headers_subco)
for i in range(len(multi_data)):
    printProgressBar(i + 1, len(multi_data))
    dict = multi_data[i]
    df_multi = df_multi.append(dict, ignore_index=True)
df_multi = df_multi.drop_duplicates(keep="first")


for i in df_subco.index:
    printProgressBar(i + 1, len(df_subco))
    row = df_subco.iloc[i]
    docuNo = row["Document No."]
    svcName = row["Service Name"]
    team = row["Driver Team"]

    m_match = df_mine.loc[
        (df_mine["DocumentNo"] == docuNo)
        & (df_mine["ServiceName"] == svcName)
        & (df_mine["Team"] == team)
    ]

    if m_match.empty:
        d = {}
        for i in headers_subco:
            d[i] = ""
        for col in headers_subco:
            d[col] = row[col]

        unmatched_data.append(d)

df_unmatched = pd.DataFrame(columns=headers_subco)
for i in range(len(unmatched_data)):
    printProgressBar(i + 1, len(unmatched_data))
    dict = unmatched_data[i]
    df_unmatched = df_unmatched.append(dict, ignore_index=True)

"""
EXCEL EXPORT AND STYLES
"""


def highlight(value):
    if value == "CHECK MULTI":
        return "background-color: #90EE90"
    elif value == "MISSING":
        return "background-color: #FFCCCB"
    elif value == "N/A":
        return "background-color: #ADD8E6"


# df_epod.to_excel(nvbc_base_writer, sheet_name="epod", index=False)
# df_sot.to_excel(nvbc_base_writer, sheet_name="sot", index=False)
# df_missingRow.to_excel(nvbc_base_writer, sheet_name="sot-missing", index=False)
# df_nvbc.to_excel(nvbc_base_writer, sheet_name="nvbc", index=False)
# print("\nEXPORTING...")
df_ori.style.applymap(
    highlight, subset=["GS svs good value", "GS Fee calculation"],
).to_excel(compare_writer, sheet_name="ORIGINAL FILL", index=False)
df_multi.to_excel(compare_writer, sheet_name="MULTI", index=False)
df_unmatched.to_excel(compare_writer, sheet_name="NO MATCH", index=False)
# df_happy_complete.style.applymap(
#     highlight,
#     subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
# ).to_excel(nvbc_base_writer, sheet_name="S Order", index=False)
# outlier_h.style.applymap(
#     highlight,
#     subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
# ).to_excel(nvbc_base_writer, sheet_name="S OUTLIER Order", index=False)
# outlier_s.style.applymap(
#     highlight,
#     subset=["EPOD", "SOT", "MVBC", "Subco Overweight Status", "Ikea Overweight Status"],
# ).to_excel(nvbc_base_writer, sheet_name="A&R OUTLIER Order", index=False)
# df_unmatch_sot.to_excel(nvbc_base_writer, sheet_name="Unmatch SOT", index=False)
compare_writer.save()
# print("\nDone. Please check the file at the location you have saved.")
