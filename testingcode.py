# title123 = "2025-06 Cumulative Update for Windows 10 Version 22H2 for amdx64-based Systems (KB5063159)"
# arch123 = "64"

# title = title123.upper()
# arch = arch123.upper()

# if "AMD" in title or "AMD" in arch:
#     print("1,2")
# elif "ARM" in title or "ARM" in arch:
#     print("3")
# else:
#     print("1,2,3")



# def get_cpu_arch(soup):
#     cpu_keywords = ["AMD", "ARM", "arm64", "amd64", "ARM64", "AMD64", "arm86", "amd86", "AMD86", "ARM86"]

#     # Get title from soup
#     #title = get_title(soup)

#     # Directly extract architecture content from archDiv
#     arch_div = soup.find("div", {"id": "archDiv"})
#     print(arch_div)
#     #arch = arch_div.text.strip()

#     # Combine title and architecture
#     complete_word = arch_div

#     # Find first matched keyword
#     matched_word = next((word for word in cpu_keywords if word in complete_word), None)

#     if matched_word:
#         if "arm" in matched_word.lower():
#             return "3"
#         elif "amd64" in matched_word.lower():
#             return "1, 2"
#     else:
#         return "1, 2, 3"


# def get_cpu_arch(soup):
#     cpu_keywords = ["AMD", "ARM", "arm64", "amd64", "ARM64", "AMD64", "arm86", "amd86", "AMD86", "ARM86"]

#     title = get_title(soup)
#     arch = get_architecture(soup)
#     complete_word = title + " " + arch

#     matched_word = next((word for word in cpu_keywords if word in complete_word), None)

#     if matched_word:
#         if "arm" in matched_word.lower():
#             return "3"
#         elif "amd64" in matched_word.lower():
#             return "1, 2"
#     else:
#         return "1, 2, 3"


#Child Function 11: get_cpu_arch
# def get_cpu_arch(soup):
#     AMD_List = ["AMD", "amd", "amd64", "AMD64", "amd86", "AMD86"]
#     ARM_List = ["ARM", "arm", "arm64", "ARM64", "arm86", "ARM86"]

#     title = get_title(soup)
#     arch = get_architecture(soup)

#     # Match either title or arch with AMD list
#     if any(item in title for item in AMD_List) or any(item in arch for item in AMD_List):
#         return "1, 2"
    
#     # Match either title or arch with ARM list
#     if any(item in title for item in ARM_List) or any(item in arch for item in ARM_List):
#         return "3"

#     # Default case
#     return "1, 2, 3"


# def get_cpu_arch(soup):
#     AMD_List = ["AMD", "amd", "amd64", "AMD64", "amd86", "AMD86"]
#     ARM_List = ["ARM", "arm", "arm64", "ARM64", "arm86", "ARM86"]

#     title = get_title(soup).replace(" ", "")  # Remove all spaces
#     arch = get_architecture(soup)

#     match_found = False

#     for amd in AMD_List:
#         if amd in title or amd in arch:
#             return "1, 2"

#     for arm in ARM_List:
#         if arm in title or arm in arch:
#             return "3"

#     return "1, 2, 3"

# def get_cpu_arch(soup):
#     #cpu_keywords = ["AMD", "ARM", "INTEL"]
    
#     title = get_title(soup).replace("(", " ").replace(")", " ")  # remove brackets
#     title_parts = title.replace(",", " ").replace("-", " ").split()  # split on space, comma, dash
#     arch = get_architecture(soup)

#     if "AMD64" in title_parts or "AMD64" in arch:
#         return "1, 2"
#     elif "ARM" in title_parts or "ARM" in arch:
#         return "3"
#     else: 
#         return "1, 2, 3"


# Child Function 7: Get Architecture
# def get_architecture(soup):
#     title = get_title(soup)
#     if "x64" in title:
#         return "x64"
#     elif "x86" in title:
#         return "x86"
#     else:
#         return "N/A"

# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# import time
# from bs4 import BeautifulSoup
# import re
# import pandas as pd
# import os
# from datetime import datetime

# # Step 1: Get today's folder path
# base_dir = os.getcwd()
# today_str = datetime.now().strftime("%Y-%m-%d")
# folder_name = f"Patch-{today_str}"
# folder_path = os.path.join(base_dir, folder_name)

# # Step 2: Initialize a set for unique Patch IDs
# unique_ids = set()

# # Step 3: Read Patch IDs from all Excel files in the folder
# if os.path.exists(folder_path):
#     for file in os.listdir(folder_path):
#         if file.endswith(".xlsx"):
#             file_path = os.path.join(folder_path, file)
#             try:
#                 df = pd.read_excel(file_path, usecols=[0])  # Assuming Patch ID is in first column
#                 ids = df.iloc[:, 0].dropna().astype(str).tolist()
#                 unique_ids.update(ids)  # Ensures only unique entries
#             except Exception as e:
#                 print(f"Error reading {file}: {e}")
# else:
#     print(f"Folder not found: {folder_path}")

# # Step 4: Convert set to list (if needed for indexing)
# all_ids = list(unique_ids)

# print(f"\nTotal unique Patch IDs collected: {len(all_ids)}")



import os
import pandas as pd
from datetime import datetime

# Step 1: Get today's folder path
base_dir = os.getcwd()
today_str = datetime.now().strftime("%Y-%m-%d")
folder_name = f"Patch-{today_str}"
folder_path = os.path.join(base_dir, folder_name)

# Step 2: Initialize the list to collect all Patch IDs
all_ids = []

# Step 3: Read patch IDs from all Excel files in the folder
if os.path.exists(folder_path):
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path, usecols=[0])  # Only read the first column
                patch_ids = df.iloc[:, 0].dropna().tolist()  # Drop blanks, convert to list
                all_ids += patch_ids  # Append to main list (no set, allows duplicates)
            except Exception as e:
                print(f"Error reading {file}: {e}")
else:
    print(f"Folder not found: {folder_path}")

# Optional: Cast all values to string if needed
all_ids = [str(patch_id).strip() for patch_id in all_ids]

print(f"\nTotal Patch IDs collected: {len(all_ids)}")
