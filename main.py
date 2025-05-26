import requests
from bs4 import BeautifulSoup

url = "https://www.catalog.update.microsoft.com/Search.aspx?q=2025"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    print(soup.prettify()[:3000])  # print part of the HTML
else:
    print("Failed to load page:", response.status_code)
