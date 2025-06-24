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

        if not download_div:
            print("\nDownload section not found.")
            return

        # Step 5: Find <hr> and then process divs after that
        hr_tag = download_div.find('hr')
        if not hr_tag:
            print("No <hr> tag found in downloadFiles section.")
            return

        # Step 6: Loop through siblings after <hr>
        current = hr_tag.find_next_sibling()
        while current:
            if current.name == 'div':
                a_tag = current.find('a')
                if a_tag:
                    PatchDownloadName = a_tag.get('title', 'N/A')
                    PatchDownloadLink = a_tag.get('href', 'N/A')
                    PatchDownloadText = current.get_text(strip=True)

                    # Remove the PatchDownloadName portion from the text (if it exists)
                    if PatchDownloadName in PatchDownloadText:
                        PatchDownloadText = PatchDownloadText.split(PatchDownloadName, 1)[-1].strip()

                    parenthesis_matches = re.findall(r'\(([^)]+)\)', PatchDownloadText)

                    # Initialize variables
                    SHA1_code = SHA2_code = "N/A"

                    if len(parenthesis_matches) == 1:
                        code = parenthesis_matches[0]
                        if code.startswith("SHA1:"):
                            SHA1_code = code.replace("SHA1:", "").strip()
                        elif code.startswith("SHA2:") or code.startswith("SHA256:"):
                            SHA2_code = code.replace("SHA2:", "").replace("SHA256:", "").strip()

                    elif len(parenthesis_matches) >= 2:
                        code1 = parenthesis_matches[0]
                        code2 = parenthesis_matches[1]

                        if code1.startswith("SHA1:"):
                            SHA1_code = code1.replace("SHA1:", "").strip()
                        elif code1.startswith("SHA2:") or code1.startswith("SHA256:"):
                            SHA2_code = code1.replace("SHA2:", "").replace("SHA256:", "").strip()

                        if code2.startswith("SHA1:"):
                            SHA1_code = code2.replace("SHA1:", "").strip()
                        elif code2.startswith("SHA2:") or code2.startswith("SHA256:"):
                            SHA2_code = code2.replace("SHA2:", "").replace("SHA256:", "").strip()

                    # Print final results
                    print(f"\nPatchDownloadName: {PatchDownloadName}")
                    print(f"PatchDownloadLink: {PatchDownloadLink}")
                    print(f"PatchDownloadText: {PatchDownloadText}")
                    print(f"SHA1_code: {SHA1_code}")
                    print(f"SHA2_code: {SHA2_code}")

            current = current.find_next_sibling()

    except Exception as e:
        print(f"\nError: {e}")
    finally:
        driver.quit()

# Run with example patch ID
patch_id = "73028972-4c54-4987-8cf5-4d10ce609560"
#patch_id = "c7ce6b24-00a1-4018-aeeb-13ae57b15b51"
open_and_scrape_download_page(patch_id)
