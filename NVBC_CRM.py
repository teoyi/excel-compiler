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
FILE_PATH_NVBC = "C:\\Users\\yipen\\Desktop\\nvbc(1-18may).xlsx"
FILE_PATH_CRM = "C:\\Users\\yipen\\Desktop\\21-05-2022-MGL Service Order Report.csv"

df_mine = pd.read_excel(FILE_PATH_NVBC, sheet_name="17 - 30 APR")
df_subco = pd.read_excel(FILE_PATH_CRM, sheet_name="Fee calculation")

# CREATE WRITER TO OUTPUT MULTIPLE SHEETS
compare_writer = pd.ExcelWriter(
    "C:\\Users\\yipen\\Desktop\\NVBC_CRM_merge.xlsx", engine="xlsxwriter"
)

print(df_mine)
print(df_subco)


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
