import os
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import re
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
    #print("\nType of all_ids:", type(all_ids))
    #print("Total Patch IDs collected:", len(all_ids))
    #print("\nAll Patch IDs:")
    #for idx, pid in enumerate(all_ids, start=1):
        #print(f"{idx}. {pid}")


# Child Function 1: Get Title
def get_title(soup):
    title_div = soup.find("div", {"id": "titleDiv"})
    if title_div:
        span = title_div.find("span", {"id": "ScopedViewHandler_titleText"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 2: Get Date
def get_date(soup):
    date_div = soup.find("div", {"id": "dateDiv"})
    if date_div:
        span = date_div.find("span", {"id": "ScopedViewHandler_date"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 3: Get Size
def get_size(soup):
    size_div = soup.find("div", {"id": "sizeDiv"})
    if size_div:
        span = size_div.find("span", {"id": "ScopedViewHandler_size"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 4: Get Description
def get_desc(soup):
    desc_div = soup.find("div", {"id": "descDiv"})
    if desc_div:
        span = desc_div.find("span", {"id": "ScopedViewHandler_desc"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 5: Get Architecture
def get_architecture(soup):
    title = get_title(soup)
    if "x64" in title:
        return "x64"
    elif "x86" in title:
        return "x86"
    else:
        return "N/A"
    
# Child Function 6: Get KB ID
def get_kb_id(soup):
    title = get_title(soup)
    match = re.search(r"\(KB\d+\)", title)
    if match:
        return match.group(0).strip("()")  # remove parentheses
    return "N/A"

# Child Function 7: Get OS
def get_os(soup):
    title = get_title(soup)

    os_list = [
        "Windows 11", "Windows 10 S", "Windows 10", "Windows 8/8.1",
        "Windows 7", "Windows Vista", "Windows XP", "Windows 2000",
        "Windows 98", "Windows 95", "Windows 3.1", "Windows 3.0", "Windows 1.0"
    ]

    for os_name in os_list:
        if os_name in title:
            return os_name

    return "N/A"

# Child Function 7: Get OS_Version
def get_os_version(soup):
    title = get_title(soup)

    version_list = [
        "NT 3.1", "NT 3.5", "NT 3.51", "NT 4.0",
        "2000", "2003", "2003 R2", "2008", "2008 R2", "2012", "2012 R2",
        "2016", "2019", "2022",
        "19H1", "19H2", "20H1", "20H2", "21H1", "21H2", "22H2", "23H2", "24H2",
        "1507", "1511", "1607", "1703", "1803", "1809", "1903", "1909", "2004"
    ]

    for version in version_list:
        if version in title:
            return version

    return "N/A"


# Parent Function to scrape first ID
def scrape_first_patch_details(patch_ids):
    if not patch_ids:
        print("No patch IDs found.")
        return

    first_id = patch_ids[0]
    url = f"https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={first_id}"

    # Setup Selenium
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract all info
        title = get_title(soup)
        date = get_date(soup)
        size = get_size(soup)
        desc = get_desc(soup)
        arch = get_architecture(soup)
        KbId = get_kb_id(soup)
        OS = get_os(soup)
        OS_Version = get_os_version(soup)

        print(f"\nPatch Details for ID: {first_id}")
        print("Title       :", title)
        print("Date        :", date)
        print("Size        :", size)
        print("Description :", desc)
        print("Architecture:", arch)
        print("KB_ID       :", KbId)
        print("OS          :", OS)
        print("OS_Version  :", OS_Version)
    finally:
        driver.quit()

scrape_first_patch_details(all_ids)

