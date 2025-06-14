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

# Child Function 8: Get OS_Version  [AMD=1, Intel=2, ARM=3]
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

# Child Function 9: get_cpu_arch
def get_cpu_arch(soup):
    cpu_keywords = ["AMD", "ARM", "INTEL"]
    
    title = get_title(soup).upper()
    arch = get_architecture(soup).upper()

    if "AMD" in title or "AMD" in arch:
        return "1, 2"
    elif "ARM" in title or "ARM" in arch:
        return "3"
    elif any(cpu in title or cpu in arch for cpu in cpu_keywords):
        # If matched with any specific keyword, already handled above
        pass
    return "1, 2, 3"

# Child Function 10: Get More Information URL
def more_info(soup):
    kb_id = get_kb_id(soup)  # e.g., "KB5022282"
    if kb_id == "N/A":
        return "N/A"
    
    kb_no = kb_id.replace("KB", "")  # e.g., "5022282"

    more_info_div = soup.find("div", {"id": "moreInfoDiv"})
    if more_info_div:
        anchors = more_info_div.find_all("a", href=True)
        for a in anchors:
            href = a['href']
            if kb_no in href:
                return href.strip()

    # Fallback if not found in page
    return f"https://support.microsoft.com/help/{kb_no}"

# Child Function 11: Get support_url 
def support_url(soup):
    kb_id = get_kb_id(soup)  # e.g., "KB5022282"
    if kb_id == "N/A":
        return "N/A"

    kb_no = kb_id.replace("KB", "")  # e.g., "5022282"

    support_div = soup.find("div", {"id": "suportUrlDiv"})
    if support_div:
        anchors = support_div.find_all("a", href=True)
        for a in anchors:
            href = a['href']
            if kb_no in href:
                return href.strip()

    # Fallback URL if no match is found
    return f"https://support.microsoft.com/help/{kb_no}"

# Child Function 12: Get update_type
def update_type(soup):
    classification_div = soup.find("div", {"id": "classificationDiv"})
    if classification_div:
        text = classification_div.text.replace("Classification:", "").strip()
        return text
    return "N/A"

# Child Function 13: Get severity
def get_severity(soup):
    severity_div = soup.find("div", {"id": "msrcSeverityDiv"})
    if severity_div:
        text = severity_div.text.replace("MSRC severity:", "").strip()
        return text
    return "N/A"

# Child Function 14: Get MSRC_number
def MSRC_number(soup):
    bulletin_div = soup.find("div", {"id": "securityBullitenDiv"})
    if bulletin_div:
        text = bulletin_div.text.replace("MSRC Number:", "").strip()
        return text
    return "N/A"

# Child Function 15: Restart_Patch enable/disable [enable=1, disable=0]
def Restart_Patch(soup):
    reboot_div = soup.find("div", {"id": "rebootBehaviorDiv"})
    if reboot_div:
        text = reboot_div.text.strip()
        if "Can request restart" in text:
            return 1
    return 0

# Child Function 16: user_input
def user_input(soup):
    user_input_div = soup.find("div", {"id": "userInputDiv"})
    if user_input_div:
        text = user_input_div.text.strip()
        return text.replace("May request user input:", "").strip()
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
        CPU_Arch = get_cpu_arch(soup)
        info_url  = more_info(soup)
        support_url_value = support_url(soup)
        update_type_value = update_type(soup)
        get_msrc_severity = get_severity(soup)
        get_msrc_number = MSRC_number(soup)
        restart_enable = Restart_Patch(soup)
        may_request_user = user_input(soup)
        


        print(f"\nPatch Details for ID: {first_id}")
        print("Title       :", title)
        print("Date        :", date)
        print("Size        :", size)
        print("Description :", desc)
        print("Architecture:", arch)
        print("KB_ID       :", KbId)
        print("OS          :", OS)
        print("OS_Version  :", OS_Version)
        print("CPU_Arch    :", CPU_Arch)
        print("more_info   :", info_url)
        print("support_URL :", support_url_value)
        print("Update_Type :", update_type_value)
        print("Severity    :", get_msrc_severity)
        print("MSRC_Number :", get_msrc_number)
        print("Restart     :", restart_enable) 
        print("Request_UserInput:", may_request_user)

    finally:
        driver.quit()

scrape_first_patch_details(all_ids)

