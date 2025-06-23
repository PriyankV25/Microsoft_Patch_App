from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup
import re

def open_and_scrape_download_page(patch_id):
    url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={patch_id}"

    # Setup Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(), options=chrome_options)

    try:
        # Step 1: Open Microsoft Update Catalog search page
        driver.get(url)
        time.sleep(5)

        # Step 2: Click on the Download button
        download_button = driver.find_element(By.XPATH, '//input[@type="button" and @value="Download"]')
        download_button.click()
        time.sleep(5)

        # Step 3: Switch to the new tab
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(3)

        # Step 4: Parse the download dialog page
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        download_div = soup.find('div', id='downloadFiles')

        if download_div:
            a_tag = download_div.find('a')

            # Extract PatchDownloadName and PatchDownloadLink
            PatchDownloadName = a_tag.get('title', 'N/A') if a_tag else 'N/A'
            PatchDownloadLink = a_tag.get('href', 'N/A') if a_tag else 'N/A'

            # Extract PatchSecurityCode
            PatchSecurityCode = 'N/A'
            for div in download_div.find_all('div'):
                match = re.search(r'\((SHA[^\)]+)\)', div.get_text(strip=True))
                if match:
                    PatchSecurityCode = match.group(1)
                    break

            # Print Results
            print(f"\nPatchDownloadName: {PatchDownloadName}")
            print(f"PatchDownloadLink: {PatchDownloadLink}")
            print(f"PatchSecurityCode: {PatchSecurityCode}")

        else:
            print("\nDownload section not found.")

    except Exception as e:
        print(f"\nError: {e}")
    finally:
        driver.quit()

# Run with example patch ID
patch_id = "c7ce6b24-00a1-4018-aeeb-13ae57b15b51"
open_and_scrape_download_page(patch_id)
