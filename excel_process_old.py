

# import sys
# print(sys.executable)

import pandas as pd
import numpy as np
import glob
import os
import re
from datetime import datetime, timedelta

from bulk_import import cox_bulk_import
# INPUT_FOLDER = r"C:\Users\RohansarathyGoudhama\Downloads\cox"
# COX_FILE = r"C:\Users\RohansarathyGoudhama\OneDrive - Yitro Technology Solutions Pvt Ltd\cox_work_order_import\cox.xlsx"
# SHEET_NAME = "MergedData"

def validate_excel(driver, mode, COX_FILE, credentials,  download_folder, log_file):
    print("Starting Excel validation process...")
    def extract_tech_id(value):
        if not isinstance(value, str):
            return None

        match = re.search(r'(\d+)\s*\(', value)
        if match:
            return int(match.group(1))
        return None


    csv_files  = glob.glob(os.path.join(download_folder, "*.csv"))
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

    ### Insert Tech ID column ###
    merged_df.insert(1, 'Tech ID', source_col.apply(extract_tech_id))
    merged_df["Tech ID"] = pd.to_numeric(merged_df["Tech ID"],errors="coerce")
    merged_df.to_excel(COX_FILE, index=False)


    filenames = [COX_FILE]
    # target_word = "Bulk"
    target_word = "Bulk|Default|Bucket|Backfill Bucket"
    for filename in filenames:
        df = pd.read_excel(filename)
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
        ### Remove HST/UHT column ###
        if "HST/UHT" in df.columns:
            df.drop(columns=["HST/UHT"], inplace=True)

        ### Append Date header in start Header###
        if "Start" in df.columns:
            df.rename(columns={"Start": "Date"}, inplace=True)
        
        #### Update Date in Date column ####
        # current_date = datetime.now().strftime("%m/%d/%Y")                                        
        # df["Date"] = current_date
        # df.to_excel(filename, index=False)
        today = datetime.now().date()
        if mode == "previous":
            process_date = (today - timedelta(days=1)).strftime("%m/%d/%Y")
            print(f"Processing previous day's file with date: {process_date}")
        else:
            process_date = today.strftime("%m/%d/%Y")
            print(f"Processing current day's file with date: {process_date}")

        df["Date"] = process_date

        ### Delete the Empty rows ###
        if "Service window" in df.columns:
            df["Service window"] = df["Service window"].astype(str).str.strip()
            # try: 
            print("Check first condition method")
            print("Total rows before:", len(df))
            df = df[
                df["Service window"].notna() &
                (df["Service window"].astype(str).str.strip() != "")
            ]
            print("Total rows after:", len(df))
            df.to_excel(filename, index=False)
        else:
            print("'Service window' column not found in the DataFrame.")

        #### Filter Work Skill #####
        # mask = df["Work Skill"].astype(str).str.contains( "Cox - CB", case=False, regex=False)
        # df.loc[mask, "Manual ICOMS"] = "Commercial"
        # df.loc[~mask, "Manual ICOMS"] = "Residential"
        df["Manual ICOMS"] = np.where(
            df["Work Skill"].fillna("").str.contains("Cox - CB", case=False, regex=False),
            "Commercial",
            "Residential"
        )
        df.to_excel(filename, index=False)
        print(f"Finished processing: {filename}")
        print("---------------------------------------------------")

    cox_bulk_import(driver, mode, COX_FILE, credentials, log_file)