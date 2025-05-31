from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import os
from datetime import date
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import pandas as pd
from datetime import datetime

# 1. Function to build the query URL
def build_search_url(query: str) -> str:
    base_url = "https://www.catalog.update.microsoft.com/Search.aspx?q="
    return f"{base_url}{query}"

#2. function to get the total page
def get_update_summary_info(url: str) -> tuple:

    # Set up Selenium WebDriver
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        number_div = soup.find("div", {"id": "numberOfUpdates"})
        if number_div:
            span = number_div.find("span", {"id": "ctl00_catalogBody_searchDuration"})
            if span:
                summary_text = span.get_text(strip=True)
                
                # Extract page count if ends with ')'
                page_count = None
                if summary_text.endswith(")"):
                    parts = summary_text.split()
                    if parts:
                        last_part = parts[-1]       # get last word, e.g., "5)"
                        page_count = last_part.rstrip(')')
                
                return summary_text, page_count
            else:
                return "Search duration span not found.", None
        else:
            return "Number of updates div not found.", None

    finally:
        driver.quit()



# 2. Function to perform the web scraping and return list of <tr> IDs under #tableContainer
def scrape_tr_ids(url: str) -> list:

    # Set up Selenium WebDriver
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        time.sleep(5)  # Wait for JavaScript to render content
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        table_container = soup.find("div", {"id": "tableContainer"})
        if not table_container:
            return []

        tr_tags = table_container.find_all("tr", id=True)

        # Skip the first <tr>, and clean the rest
        cleaned_ids = []
        for tr in tr_tags[1:]:  # skip first tr
            raw_id = tr['id']
            cleaned_id = raw_id.split('_R')[0]  # Remove everything from "_R" onward
            cleaned_ids.append(cleaned_id)
        
        return cleaned_ids

    finally:
        driver.quit()


# 3. Function to print the output
def print_patch_summary(summary_text: str, page_count: str):
    print("\nPatch Summary Info:")
    print(summary_text)
    if page_count:
        print(f"Total Pages: {page_count}")
    else:
        print("Page count not found.")



# 4. Function to print the output
def save_ids_to_excel(ids: list[str], base_dir: str = "."):
    # Create folder "Patch-YYYY-MM-DD"
    date_str = datetime.now().strftime("%Y-%m-%d")
    folder_name = f"Patch-{date_str}"
    folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    # Save to Excel
    df = pd.DataFrame(ids, columns=["Patch ID"])
    excel_path = os.path.join(folder_path, "patch_ids.xlsx")
    df.to_excel(excel_path, index=False)
    print(f"Saved {len(ids)} IDs to: {excel_path}")


def scrape_all_pages(query: str, total_pages: int) -> list:
    all_tr_ids = []

    for i in range(total_pages):
        paged_url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={query}&p={i}"
        print(f"\nScraping page {i + 1} of {total_pages}...")
        tr_ids = scrape_tr_ids(paged_url)
        all_tr_ids.extend(tr_ids)
          # Print IDs for this page immediately
        for tr_id in tr_ids:
            print(tr_id)

        print()  # Add one blank line after each page's output
    return all_tr_ids


# --- Main Execution ---
if __name__ == "__main__":
    query = "2025-05"
    url = build_search_url(query)

    summary, page_count = get_update_summary_info(url)
    print_patch_summary(summary, page_count)

    if page_count:
        total_pages = int(page_count)
        all_tr_ids = scrape_all_pages(query, total_pages)
        save_ids_to_excel(all_tr_ids)
        
    else:
        print("Unable to fetch page count. Exiting.")

    
