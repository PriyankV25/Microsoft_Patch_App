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

# Step 3: Read patch IDs from all Excel files in the folder
if os.path.exists(folder_path):
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path, usecols=[0])  # Assuming Patch ID is in the first column
                all_ids.extend(df.iloc[:, 0].dropna().tolist())
            except Exception as e:
                print(f"Error reading {file}: {e}")
else:
    print(f"Folder not found: {folder_path}")

print(f"\nTotal Patch IDs collected: {len(all_ids)}")

# Step 4: Define the scraper function
def open_and_scrape_download_page(patch_id):
    url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={patch_id}"

    # Setup Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=Service(), options=chrome_options)

    data_row = {'patchid': patch_id}
    try:
        driver.get(url)
        time.sleep(5)

        download_button = driver.find_element(By.XPATH, '//input[@type="button" and @value="Download"]')
        download_button.click()
        time.sleep(5)

        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(3)

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
        index = 1

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

                    data_row[f'PatchDownloadName_{index}'] = PatchDownloadName
                    data_row[f'PatchDownloadLink_{index}'] = PatchDownloadLink
                    data_row[f'PatchDownloadText_{index}'] = PatchDownloadText
                    data_row[f'SHA1_code_{index}'] = SHA1_code
                    data_row[f'SHA256_code_{index}'] = SHA256_code

                    index += 1

            current = current.find_next_sibling()

    except Exception as e:
        print(f"\nError while processing {patch_id}: {e}")
    finally:
        driver.quit()

    df = pd.DataFrame([data_row])
    output_file = "patch_downloads.xlsx"

    try:
        existing_df = pd.read_excel(output_file)
        final_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        final_df = df

    final_df.to_excel(output_file, index=False)
    print(f"Data for patch ID {patch_id} saved successfully.")

# Step 5: Process all collected patch IDs
for pid in all_ids:
    open_and_scrape_download_page(pid)
