import os
from datetime import datetime
import pandas as pd

# Step 1: Get today's folder path
base_dir = os.getcwd()
today_str = datetime.now().strftime("%Y-%m-%d")
folder_name = f"Patch-{today_str}"
folder_path = os.path.join(base_dir, folder_name)

# Step 2: Initialize the list to collect all Patch IDs
all_ids = []

# Step 3: Check if folder exists
if not os.path.exists(folder_path):
    print(f"Folder not found: {folder_path}")
else:
    # Step 4: Get all Excel files in the folder
    excel_files = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.endswith(".xlsx")
    ])

    # Step 5: Read each Excel file and collect IDs
    for file in excel_files:
        print(f"\nReading from: {file}")
        df = pd.read_excel(file)
        ids = df.iloc[:, 0].dropna().astype(str).tolist()
        all_ids.extend(ids)

    # Step 6: Confirm list and print results
    print("\nType of all_ids:", type(all_ids))
    print("Total Patch IDs collected:", len(all_ids))
    print("\nAll Patch IDs:")
    #for idx, pid in enumerate(all_ids, start=1):
        #print(f"{idx}. {pid}")


