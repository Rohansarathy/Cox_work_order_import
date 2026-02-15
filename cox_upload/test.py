

# import sys
# print(sys.executable)

import pandas as pd
import numpy as np
import glob
import os
import re
from datetime import datetime
INPUT_FOLDER = r"C:\Users\RohansarathyGoudhama\Downloads\cox"
COX_FILE = r"C:\Users\RohansarathyGoudhama\Downloads\cox\cox_error_logs\Cox.xlsx"
SHEET_NAME = "MergedData"



# def transform_value(value):
#     if not isinstance(value, str):
#         return value
#     # value = value.strip()
#     # if not value:
#     #     return value
#     if value and value[0] in {"C", "G", "A", "H", "M", "T"}:
#         try:
#             start = value.index("-") + 2
#             end = value.index("(", start)
#             return value[start:end].strip()
#         except ValueError:
#             return value
#     print(f"Unchanged value: {value}")
#     return value


def extract_tech_id(value):
    if not isinstance(value, str):
        return None

    match = re.search(r'-\s*(\d+)', value)
    if match:
        return match.group(1)


    match = re.search(r'(\d+)\s*\(', value)
    if match:
        # return int(match.group(1))
        return match.group(1)
    return None


csv_files  = glob.glob(os.path.join(INPUT_FOLDER, "*.csv"))
if not csv_files:
    raise Exception("No CSV files found in the input folder")

dataframes = []

for file in csv_files:
    print(f"Reading: {file}")
    df = pd.read_csv(file)
    df["Source_File"] = os.path.basename(file)
    dataframes.append(df)

merged_df = pd.concat(dataframes, ignore_index=True)
print("Total rows merged:", len(merged_df))
source_col = merged_df.iloc[:, 0]
# merged_df.insert(1, 'Tech ID', merged_df.iloc[:, 0].apply(transform_value))
# merged_df.insert(1, 'Tech ID',  source_col.apply(transform_value))
# merged_df["Tech ID"] =  source_col.apply(extract_tech_id)

### Insert Tech ID column ###
merged_df.insert(1, 'Tech ID', source_col.apply(extract_tech_id))
merged_df["Tech ID"] = pd.to_numeric(merged_df["Tech ID"],errors="coerce")
merged_df.to_excel(COX_FILE, index=False)


filenames = [COX_FILE]
# target_word = "Bulk"
target_word = "Bulk|Default|Bucket|Backfill Bucket"
for filename in filenames:
    df = pd.read_excel(filename)
    # mask = df.apply(lambda r: r.astype(str).str.contains(target_word, case=False).any(), axis=1)


    #### HOLD Activity Status for Bulk, Default, Bucket, Backfill Bucket ####
    mask = df["HST/UHT"].astype(str).str.contains(target_word, case=False, regex=True)
    if mask.any():
        print(f"Found '{target_word}' in {filename}")
        print(df[mask])
        df.loc[mask, "Activity Status"] = "Hold"
        df.to_excel(filename, index=False)
    else:
        print(f"'{target_word}' not found in {filename}")
    
    print("Columns in file:")
    print(df.columns.tolist())
    df.columns = df.columns.str.strip()
    ## Remove HST/UHT column ###
    # if "HST/UHT" in df.columns:
    #     df.drop(columns=["HST/UHT"], inplace=True)

    ### Append Date header in start Header###
    if "Start" in df.columns:
        df.rename(columns={"Start": "Date"}, inplace=True)
    
    #### Update current DateTime in Date column ####
    current_date = datetime.now().strftime("%m/%d/%Y")                                        
    df["Date"] = current_date
    df.to_excel(filename, index=False)

    ### Delete the Empty rows ###
    # if "Service window" in df.columns:
    #     df["Service window"] = df["Service window"].astype(str).str.strip()
    #     # try: 
    #     print("Check first condition method")
    #     print("Total rows before:", len(df))
    #     df = df[
    #         df["Service window"].notna() &
    #         (df["Service window"].astype(str).str.strip() != "")
    #     ]
    #     print("Total rows after:", len(df))
    #     df.to_excel(filename, index=False)
    # else:
    #     print("'Service window' column not found in the DataFrame.")

    #### Filter Work Skill #####
    df["Manual ICOMS"] = np.where(
        df["Work Skill"].fillna("").str.contains("Cox - CB", case=False, regex=False),
        "Commercial",
        "Residential"
    )
    df.to_excel(filename, index=False)



# print("Data successfully written to cox.xlsx")  