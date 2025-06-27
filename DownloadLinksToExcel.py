from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
from datetime import datetime
import sys

# Step 1: Get today's folder path
base_dir = os.getcwd()
today_str = datetime.now().strftime("%Y-%m-%d")
folder_name = f"Patch-{today_str}"
folder_path = os.path.join(base_dir, folder_name)

# Step 2: Initialize the list to collect all Patch IDs
all_ids = []


# Step 3: Read patch_ids.xlsx from today's folder
patch_ids_file = os.path.join(folder_path, "patch_ids.xlsx")
if os.path.exists(patch_ids_file):
    try:
        df_patch_ids = pd.read_excel(patch_ids_file, usecols=[0])
        all_ids = df_patch_ids.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    except Exception as e:
        print(f"Error reading patch_ids.xlsx: {e}")
        sys.exit(1)
else:
    print(f"patch_ids.xlsx not found in {folder_path}")
    sys.exit(1)

# Optional: Cast all values to string if needed
all_ids = [str(patch_id).strip() for patch_id in all_ids]

print(f"\nTotal Patch IDs collected: {len(all_ids)}")


# Step 4: Resume logic based on patch_downloads.xlsx
start_index = 0
output_file = os.path.join(folder_path, "patch_downloads.xlsx")
if os.path.exists(output_file):
    try:
        df_existing = pd.read_excel(output_file, usecols=[0])  # Only read column A (patchid)
        if not df_existing.empty:
            last_patch_id = str(df_existing.iloc[-1, 0]).strip()
            if last_patch_id in all_ids:
                start_index = all_ids.index(last_patch_id) + 1
                print(f"Resuming from index {start_index}: {all_ids[start_index] if start_index < len(all_ids) else 'End of List'}")
            else:
                print(f"Last patch ID '{last_patch_id}' not found in patch_ids.xlsx — starting from beginning.")
    except Exception as e:
        print(f"Error reading existing patch_downloads.xlsx: {e}")
        sys.exit(1)


def open_and_scrape_download_page(patch_id):
    url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={patch_id}"

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=Service(), options=chrome_options)

    data_row = {'patchid': patch_id}
    index = 1  # download link counter

    try:
        # Load page with timeout
        driver.get(url)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@type="button" and @value="Download"]'))
            )
        except TimeoutException:
            print(f"\nTimeout: Page did not load properly for patch ID: {patch_id}")
            driver.quit()
            sys.exit(1)

        # Click the Download button
        try:
            download_button = driver.find_element(By.XPATH, '//input[@type="button" and @value="Download"]')
            download_button.click()
        except Exception as e:
            print(f"\nError clicking Download button for patch ID: {patch_id}: {e}")
            driver.quit()
            sys.exit(1)

        # Switch to download tab
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'downloadFiles')))

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        download_div = soup.find('div', id='downloadFiles')

        if not download_div:
            print(f"\nDownload section not found for {patch_id}")
            return

        hr_tag = download_div.find('hr')
        if not hr_tag:
            print(f"No <hr> tag found for {patch_id}")
            return

        current = hr_tag.find_next_sibling()

        while current:
            if current.name == 'div':
                a_tag = current.find('a')
                if a_tag:
                    PatchDownloadName = a_tag.get('title', 'N/A')
                    PatchDownloadLink = a_tag.get('href', 'N/A')
                    PatchDownloadText = current.get_text(strip=True)

                    if PatchDownloadName in PatchDownloadText:
                        PatchDownloadText = PatchDownloadText.split(PatchDownloadName, 1)[-1].strip()

                    parenthesis_matches = re.findall(r'\(([^)]+)\)', PatchDownloadText)

                    SHA1_code = SHA256_code = "N/A"
                    if len(parenthesis_matches) == 1:
                        code = parenthesis_matches[0]
                        if code.startswith("SHA1:"):
                            SHA1_code = code.replace("SHA1:", "").strip()
                        elif code.startswith("SHA2:") or code.startswith("SHA256:"):
                            SHA256_code = code.replace("SHA2:", "").replace("SHA256:", "").strip()
                    elif len(parenthesis_matches) >= 2:
                        code1, code2 = parenthesis_matches[0], parenthesis_matches[1]
                        if code1.startswith("SHA1:"):
                            SHA1_code = code1.replace("SHA1:", "").strip()
                        elif code1.startswith("SHA2:") or code1.startswith("SHA256:"):
                            SHA256_code = code1.replace("SHA2:", "").replace("SHA256:", "").strip()
                        if code2.startswith("SHA1:"):
                            SHA1_code = code2.replace("SHA1:", "").strip()
                        elif code2.startswith("SHA2:") or code2.startswith("SHA256:"):
                            SHA256_code = code2.replace("SHA2:", "").replace("SHA256:", "").strip()

                    # Save data
                    data_row[f'PatchDownloadName_{index}'] = PatchDownloadName
                    data_row[f'PatchDownloadLink_{index}'] = PatchDownloadLink
                    data_row[f'PatchDownloadText_{index}'] = PatchDownloadText
                    data_row[f'SHA1_code_{index}'] = SHA1_code
                    data_row[f'SHA256_code_{index}'] = SHA256_code

                    index += 1
            current = current.find_next_sibling()

    except WebDriverException as e:
        print(f"\nNetwork or browser error for patch ID {patch_id}: {e}")
        driver.quit()
        sys.exit(1)

    finally:
        driver.quit()

    # Set download link count
    data_row['DownloadLinkCount'] = index - 1  # Total number of download links found

    # Save to Excel
    df = pd.DataFrame([data_row])
    output_file = os.path.join(folder_path, "patch_downloads.xlsx")


    try:
        existing_df = pd.read_excel(output_file)
        final_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        final_df = df

    final_df.to_excel(output_file, index=False)
    print(f"Data for patch ID {patch_id} saved with {index - 1} download links.")





# Step 5: Loop from the appropriate starting index
for pid in all_ids[start_index:]:
    open_and_scrape_download_page(pid)
