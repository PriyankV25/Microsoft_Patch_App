import os
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



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

# Child Function 2: Patch_title_category
def Patch_title_category(soup):
    categories = [
        "CumulativeUpdate", "SafeOSDynamicUpdate", "SetupDynamicUpdate", ".NET", "QualityUpdate", 
        "MonthlyQualityRollup", "CumulativeSecurityUpdate", "ServicingStackUpdate", 
        "DynamicUpdate", "DynamicCumulativeUpdate"
    ]

    title = get_title(soup).replace(" ", "")  # Remove all spaces to normalize
    title_upper = title.upper()  # Optional: to make matching case-insensitive

    for category in categories:
        if category.upper() in title_upper:
            return category
    return "N/A"

# Child Function 3: Patch_for
def Patch_for(soup):
    title = get_title(soup)
    title_lower = title.lower()

    if title_lower.count("for") >= 2:
        # Match text between two 'for' (non-greedy)
        match = re.search(r'for\s+(.*?)\s+for', title_lower, re.IGNORECASE)
        if match:
            return match.group(1).strip().title()
    elif title_lower.count("for") == 1:
        # Match everything after 'for' and before (KB...) if exists
        match = re.search(r'for\s+(.*?)(\s*\(kb.*?\))?$', title, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return "N/A"

# Child Function 4: Get Date
def get_date(soup):
    date_div = soup.find("div", {"id": "dateDiv"})
    if date_div:
        span = date_div.find("span", {"id": "ScopedViewHandler_date"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 5: Get Size
def get_size(soup):
    size_div = soup.find("div", {"id": "sizeDiv"})
    if size_div:
        span = size_div.find("span", {"id": "ScopedViewHandler_size"})
        return span.text.strip() if span else "N/A"
    return "N/A"

# Child Function 6: Get Description
def get_desc(soup):
    desc_div = soup.find("div", {"id": "descDiv"})
    if desc_div:
        span = desc_div.find("span", {"id": "ScopedViewHandler_desc"})
        return span.text.strip() if span else "N/A"
    return "N/A"
    
# Child Function 8: Get KB ID
def get_kb_id(soup):
    title = get_title(soup)
    match = re.search(r"\(KB\d+\)", title)
    if match:
        return match.group(0).strip("()")  # remove parentheses
    return "N/A"

# Child Function 9: Get OS
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

# Child Function 10: Get OS_Version  
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

# Child Function 11: Get_arch_details  [AMD,ARM][64,86]
def Get_arch_details(soup):
    arch_div = soup.find("div", {"id": "archDiv"})
    if arch_div:
        text = arch_div.text.strip().replace("Architecture:", "").strip()
        return text if text else "N/A"
    return "N/A"

#Child Function 12: get_cpu_arch  [AMD=1, Intel=2, ARM=3]
def get_cpu_arch(soup):
    cpu_keywords = ["AMD", "ARM", "INTEL"]
    
    title = get_title(soup).upper()
    arch = Get_arch_details(soup).upper()

    if "AMD" in title or "AMD" in arch:
        return "1, 2"
    elif "ARM" in title or "ARM" in arch:
        return "3"
    else:
        return "1, 2, 3"
    
# Child Function 13: Get get_architecture
def get_architecture(soup):
    title = get_title(soup)
    title_filter = title[:-12]
    archiT = Get_arch_details(soup)

    if "64" in title_filter or "64" in archiT:
        return "x64"
    elif "86" in title_filter or "86" in archiT:
        return "x86"
    else:
        return "N/A"

# Child Function 14: Get More Information URL
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

# Child Function 15: Get support_url 
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

# Child Function 16: Get update_type
def update_type(soup):
    classification_div = soup.find("div", {"id": "classificationDiv"})
    if classification_div:
        text = classification_div.text.replace("Classification:", "").strip()
        return text
    return "N/A"

# Child Function 17: Get severity       Modified
def get_severity(soup):
    severity_div = soup.find("div", {"id": "msrcSeverityDiv"})
    if severity_div:
        text = severity_div.text.replace("MSRC severity:", "").strip()
        if text.lower() == "n/a" or text == "":
            return "N/A"
        return text
    return "N/A"

# Child Function 18: Get MSRC_number    Modified
def MSRC_number(soup):
    bulletin_div = soup.find("div", {"id": "securityBullitenDiv"})
    if bulletin_div:
        text = bulletin_div.text.replace("MSRC Number:", "").strip()
        if text.lower() == "n/a" or text == "":
            return "N/A"
        return text
    return "N/A"

# Child Function 19: Restart_Patch enable/disable [enable=1, disable=0]
def Restart_Patch(soup):
    reboot_div = soup.find("div", {"id": "rebootBehaviorDiv"})
    if reboot_div:
        text = reboot_div.text.strip()
        if "Can request restart" in text:
            return 1
    return 0

# Child Function 18: user_input
def user_input(soup):
    user_input_div = soup.find("div", {"id": "userInputDiv"})
    if user_input_div:
        text = user_input_div.text.strip()
        return text.replace("May request user input:", "").strip()
    return "N/A"

# Child Function 20: Install_impact
def Install_impact(soup):
    impact_div = soup.find("div", {"id": "installationImpactDiv"})
    if impact_div:
        text = impact_div.text.replace("Must be installed exclusively:", "").strip()
        return text if text else "N/A"
    return "N/A"

# Child Function 21: connectivity_requirement
def connectivity_requirement(soup):
    conn_div = soup.find("div", {"id": "connectivityDiv"})
    if conn_div:
        text = conn_div.text.replace("Requires network connectivity:", "").strip()
        return text if text else "N/A"
    return "N/A"

# Child Function 22: Uninstall_patch [enable = 1, diable = 0]
def Uninstall_patch(soup):
    uninstall_div = soup.find("div", {"id": "uninstallNotesDiv"})
    if uninstall_div:
        text = uninstall_div.text.replace("Uninstall Notes:", "").strip().lower()
        if text == "n/a" or text == "":
            return 0
        return 1
    return 0

# Child Function 23: Uninstall_steps
def Uninstall_steps(soup):
    steps_div = soup.find("div", {"id": "uninstallStepsDiv"})
    if steps_div:
        text = steps_div.text.replace("Uninstall Steps:", "").strip()
        if text.lower() == "n/a" or text == "":
            return "N/A"
        return text
    return "N/A"



##############################################################################################


def scrape_patch_details_to_excel(patch_ids):
    if not patch_ids:
        print("No patch IDs found.")
        return

    # Setup Selenium once
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=options)

    all_patch_data = []

    try:
        for idx, patch_id in enumerate(patch_ids, start=1):
            url = f"https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={patch_id}"
            print(f"\n[{idx}/{len(patch_ids)}] Fetching patch details for Update ID: {patch_id}")

            try:
                driver.get(url)
                time.sleep(4)
                soup = BeautifulSoup(driver.page_source, 'html.parser')

                # Append all extracted values into a dictionary
                patch_data = {
                    "Update_ID": patch_id,
                    "Title": get_title(soup),
                    "Patch_title_category": Patch_title_category(soup),
                    "Patch_for": Patch_for(soup),
                    "Date": get_date(soup),
                    "Size": get_size(soup),
                    "Description": get_desc(soup),
                    "KB_ID": get_kb_id(soup),
                    "OS": get_os(soup),
                    "OS_Version": get_os_version(soup),
                    "CPU_Arch": get_cpu_arch(soup),
					"Architecture": get_architecture(soup),
                    "More_Info_URL": more_info(soup),
                    "Support_URL": support_url(soup),
                    "Update_Type": update_type(soup),
                    "Severity": get_severity(soup),
                    "MSRC_Number": MSRC_number(soup),
                    "Restart_Required": Restart_Patch(soup),
                    "Request_UserInput": user_input(soup),
                    "Install_Impact": Install_impact(soup),
                    "Connectivity_Required": connectivity_requirement(soup),
                    "Uninstall_Enabled": Uninstall_patch(soup),
                    "Uninstall_Steps": Uninstall_steps(soup)
                }

                all_patch_data.append(patch_data)

            except Exception as e:
                print(f"Failed to fetch details for ID {patch_id}: {e}")

    finally:
        driver.quit()

    # Convert list of dicts to DataFrame
    df = pd.DataFrame(all_patch_data)

    # Create today's folder path again
    today_str = datetime.now().strftime("%Y-%m-%d")
    folder_name = f"Patch-{today_str}"
    base_dir = os.getcwd()
    folder_path = os.path.join(base_dir, folder_name)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    output_file = os.path.join(folder_path, "patch_details.xlsx")
    df.to_excel(output_file, index=False)
    print(f"\n Patch details saved to: {output_file}")




scrape_patch_details_to_excel(all_ids)






########################################################################################
# # Parent Function to scrape all IDs
# def scrape_patch_details(patch_ids):
#     if not patch_ids:
#         print("No patch IDs found.")
#         return

#     # Setup Selenium outside the loop for better performance
#     service = Service()
#     options = webdriver.ChromeOptions()
#     options.add_argument('--headless')
#     driver = webdriver.Chrome(service=service, options=options)

#     try:
#         for idx, patch_id in enumerate(patch_ids, start=1):
#             url = f"https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={patch_id}"
#             print(f"\n[{idx}/{len(patch_ids)}] Fetching patch details for Update ID: {patch_id}")
            
#             try:
#                 driver.get(url)
#                 time.sleep(4)  # Small delay to let page load
#                 soup = BeautifulSoup(driver.page_source, 'html.parser')

#                 # Extract all info
#                 title = get_title(soup)
#                 Title_Category = Patch_title_category(soup)
#                 Patching_for = Patch_for(soup)
#                 date = get_date(soup)
#                 size = get_size(soup)
#                 desc = get_desc(soup)
#                 arch = get_architecture(soup)
#                 KbId = get_kb_id(soup)
#                 OS = get_os(soup)
#                 OS_Version = get_os_version(soup)
#                 CPU_Arch = get_cpu_arch(soup)
#                 info_url = more_info(soup)
#                 support_url_value = support_url(soup)
#                 update_type_value = update_type(soup)
#                 get_msrc_severity = get_severity(soup)
#                 get_msrc_number = MSRC_number(soup)
#                 restart_enable = Restart_Patch(soup)
#                 may_request_user = user_input(soup)
#                 Installation_Impact = Install_impact(soup)
#                 Connectivity = connectivity_requirement(soup)
#                 Uninstallation_patch = Uninstall_patch(soup)
#                 Uninstallation_steps = Uninstall_steps(soup)

#                 # Print summary for each patch
#                 print("Update_ID                :", patch_id)
#                 print("Title                    :", title)
#                 print("Patch_title_category     :", Title_Category)
#                 print("Patch_for                :", Patching_for)
#                 print("Date                     :", date)
#                 print("Size                     :", size)
#                 print("Description              :", desc)
#                 print("Architecture             :", arch)
#                 print("KB_ID                    :", KbId)
#                 print("OS                       :", OS)
#                 print("OS_Version               :", OS_Version)
#                 print("CPU_Arch                 :", CPU_Arch)
#                 print("more_info                :", info_url)
#                 print("support_URL              :", support_url_value)
#                 print("Update_Type              :", update_type_value)
#                 print("Severity                 :", get_msrc_severity)
#                 print("MSRC_Number              :", get_msrc_number)
#                 print("Restart                  :", restart_enable) 
#                 print("Request_UserInput        :", may_request_user)
#                 print("Install_impact           :", Installation_Impact)
#                 print("Connectivity             :", Connectivity)
#                 print("Uninstall_patch          :", Uninstallation_patch)
#                 print("Uninstall_steps          :", Uninstallation_steps)

#             except Exception as e:
#                 print(f"Failed to fetch details for ID {patch_id}: {e}")

#     finally:
#         driver.quit()


# # Call the updated function
# scrape_patch_details(all_ids)
##################################################################################