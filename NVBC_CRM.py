import pandas as pd
import numpy as np
from collections import Counter
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
FILE_PATH_NVBC = "C:\\Users\\yipen\\Desktop\\nvbc(1-18may).xlsx"
FILE_PATH_CRM = "C:\\Users\\yipen\\Desktop\\21-05-2022-MGL Service Order Report.csv"

df_nvbc = pd.read_excel(FILE_PATH_NVBC, sheet_name="Filtered")
df_crm = pd.read_csv(FILE_PATH_CRM)

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
compare_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\NVBC_CRM_merge.xlsx", engine="xlsxwriter"
)

# print(df_nvbc)
# print(df_crm)

print("\nGenerating headers...")
headers_crm = df_crm.columns.values.tolist()
headers_crm_ori = df_crm.columns.values.tolist()
headers_crm = list(
    map(lambda x: x.replace("SERVICE STATUS", "CRM SERVICE STATUS"), headers_crm)
)
headers_crm.insert(0, "NVBC")
status_index = headers_crm.index("CRM SERVICE STATUS")
headers_crm.insert(status_index, "NVBC SERVICE STATUS")
print("\nHeaders to be used")
for header in headers_crm:
    print(header)
print("\nHeaders generated")

"""
crm as base using docu no and svc order no -> keep crm all columsn 
included columns from nvbc =

nvbc ------ crm 
Service Status SERVICE STATUS


if nvbc have crm don't have add to new sheet 
"""

crm_docuNo = df_crm["DOCUMENT NO."].tolist()
crm_svcOrderNo = df_crm["SERVICE ORDER NO."].tolist()
zipped_crm = list(zip(crm_docuNo, crm_svcOrderNo))

print("Checking for duplicates")
d = Counter(zipped_crm)
res = [k for k, v in d.items() if v > 1]
print(f"Number of duplicates: {len(res)}")
if len(res) > 1:
    print(f"Duplicates Array: {res}")

df_final = pd.DataFrame(columns=headers_crm)
headers_nvbc = df_nvbc.columns.values.tolist()
df_unmatched = pd.DataFrame(columns=headers_nvbc)
data = []
crm_matched = []

for i in range(len(zipped_crm)):
    printProgressBar(i + 1, len(zipped_crm))
    pairs = zipped_crm[i]
    docuNo = pairs[0]
    svcOrderNo = pairs[1]

    nvbc_match = df_nvbc.loc[
        (df_nvbc["Document No."] == docuNo)
        & (df_nvbc["Service Order No."] == svcOrderNo)
    ]

    crm_match = df_crm.loc[
        (df_crm["DOCUMENT NO."] == docuNo) & (df_crm["SERVICE ORDER NO."] == svcOrderNo)
    ]

    if not crm_match.empty:
        d = {}
        crm_matched.append(pairs)

        for header in headers_crm:
            d[header] = "MISSING"

        for header in headers_crm_ori:
            header_data = crm_match.iloc[0][header]
            if header == "SERVICE STATUS":
                d["CRM SERVICE STATUS"] = header_data
            else:
                d[header] = header_data

        if nvbc_match.empty:
            d["NVBC"] = "0"
        else:
            d["NVBC"] = "1"
            d["NVBC SERVICE STATUS"] = nvbc_match.iloc[0]["Service Status"]
        data.append(d)

print("Building DataFrame")
for i in range(len(data)):
    printProgressBar(i + 1, len(data))
    dict = data[i]
    df_final = df_final.append(dict, ignore_index=True)


dropList = []
for pair in crm_matched:
    docuNo = pair[0]
    svcOrderNo = pair[1]
    indices = df_nvbc.index[
        (df_nvbc["Document No."] == docuNo)
        & (df_nvbc["Service Order No."] == svcOrderNo)
    ]

    for index in indices:
        dropList.append(index)
df_unmatch_nvbc = df_nvbc.drop(dropList, 0)


"""
EXCEL EXPORT AND STYLES
"""


def highlight(value):
    if value == "CHECK MULTI":
        return "background-color: #90EE90"
    elif value == "0":
        return "background-color: #FFCCCB"
    elif value == "MISSING":
        return "background-color: #FFCCCB"


df_final.style.applymap(highlight, subset=["NVBC", "NVBC SERVICE STATUS"],).to_excel(
    compare_writer, sheet_name="CRM BASE", index=False
)
# df_multi.to_excel(compare_writer, sheet_name="MULTI", index=False)
df_unmatch_nvbc.to_excel(compare_writer, sheet_name="NO MATCH", index=False)
compare_writer.save()
print("\nDone. Please check the file at the location you have saved.")
