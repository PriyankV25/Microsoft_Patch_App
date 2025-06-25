from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
from datetime import datetime


# Step 1: Get today's folder path
base_dir = os.getcwd()
today_str = datetime.now().strftime("%Y-%m-%d")
folder_name = f"Patch-{today_str}"
folder_path = os.path.join(base_dir, folder_name)

# Step 2: Initialize the list to collect all Patch IDs
all_ids = []




def open_and_scrape_download_page(patch_id):
    url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={patch_id}"

    # Setup Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=Service(), options=chrome_options)

    data_row = {'patchid': patch_id}
    try:
        driver.get(url)
        time.sleep(5)

        # Click Download button
        download_button = driver.find_element(By.XPATH, '//input[@type="button" and @value="Download"]')
        download_button.click()
        time.sleep(5)

        # Switch to new tab
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(3)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        download_div = soup.find('div', id='downloadFiles')

        if not download_div:
            print("\nDownload section not found.")
            return

        hr_tag = download_div.find('hr')
        if not hr_tag:
            print("No <hr> tag found in downloadFiles section.")
            return

        current = hr_tag.find_next_sibling()
        index = 1  # Counter for multi-link indexing

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
                        code1 = parenthesis_matches[0]
                        code2 = parenthesis_matches[1]
                        if code1.startswith("SHA1:"):
                            SHA1_code = code1.replace("SHA1:", "").strip()
                        elif code1.startswith("SHA2:") or code1.startswith("SHA256:"):
                            SHA256_code = code1.replace("SHA2:", "").replace("SHA256:", "").strip()
                        if code2.startswith("SHA1:"):
                            SHA1_code = code2.replace("SHA1:", "").strip()
                        elif code2.startswith("SHA2:") or code2.startswith("SHA256:"):
                            SHA256_code = code2.replace("SHA2:", "").replace("SHA256:", "").strip()

                    # Add data to dictionary with index
                    data_row[f'PatchDownloadName_{index}'] = PatchDownloadName
                    data_row[f'PatchDownloadLink_{index}'] = PatchDownloadLink
                    data_row[f'PatchDownloadText_{index}'] = PatchDownloadText
                    data_row[f'SHA1_code_{index}'] = SHA1_code
                    data_row[f'SHA256_code_{index}'] = SHA256_code

                    index += 1

            current = current.find_next_sibling()

    except Exception as e:
        print(f"\nError: {e}")
    finally:
        driver.quit()

    # Save the row to Excel
    df = pd.DataFrame([data_row])
    output_file = "patch_downloads.xlsx"

    try:
        existing_df = pd.read_excel(output_file)
        final_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        final_df = df

    final_df.to_excel(output_file, index=False)
    print(f"\nData for patch ID {patch_id} saved successfully.")

# Example usage
patch_id = "bba5eb25-10ea-497d-9c58-cbf797f5d0e9"
#patch_id = "c7ce6b24-00a1-4018-aeeb-13ae57b15b51"
open_and_scrape_download_page(patch_id)
