import gspread
import os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import requests
import time

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))

# Use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(credentials, scopes=scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
sheet = client.open("WriteWave2.0").worksheet('Overview')

# Extract all the values
all_values = sheet.get_all_values()

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
keyword_col_index = header_row.index('Keyword')
datasheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

# Headers for requests
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
    'Accept-Encoding': 'identity'
}
print(f'>>> START fetch_serp_data.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    keyword = row[keyword_col_index]
    datasheet_link = row[datasheet_col_index]

    # Check if the 'Datasheet' cell is empty
    if not datasheet_link:
        print(f"No Google Sheet defined for keyword, '{keyword}'")
        continue

    # Open the Google Sheet by URL
    datasheet = client.open_by_url(datasheet_link)

    # Select the 'Page Structure' worksheet
    page_structure_sheet = datasheet.worksheet('Page Structure')

    # Define the factors
    factors = ['Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count', 'H5 Tag Count',
               'H6 Tag Count',
               'p Tag Count', 'a Tag Count', 'a Internal Tag Count', 'a External Tag Count', 'img Tag Count',
               'alt Tag Count',
               'b Tag Count', 'strong Tag Count', 'i Tag Count', 'em Tag Count', 'u Tag Count', 'ol Tag Count',
               'ol Item Count', 'ul Tag Count', 'ul Item Count', 'table Tag Count', 'form Tag Count',
               'iframe Tag Count']

    # For each result column, scrape the web page and update the sheet
    for j in range(2, 12):  # Assuming there are 10 result columns starting from column 2
        url = page_structure_sheet.cell(2, j).value  # URLs are in the second row

        # Check if the URL is not empty
        if url:
            response = requests.get(url, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')

            for k, factor in enumerate(factors, start=1):
                if factor == 'Word Count':
                    count = len(soup.get_text().split())
                elif factor == 'a Internal Tag Count':
                    count = len([a for a in soup.find_all('a') if a.get('href') and a.get('href').startswith('/')])
                elif factor == 'a External Tag Count':
                    count = len([a for a in soup.find_all('a') if a.get('href') and a.get('href').startswith('http')])
                elif factor == 'alt Tag Count':
                    count = len([img for img in soup.find_all('img') if img.get('alt')])
                elif factor in ['ol Item Count', 'ul Item Count']:
                    tag = factor.split(' ')[0].lower()
                    count = sum(len(ol.find_all('li')) for ol in soup.find_all(tag))
                else:
                    tag = factor.split(' ')[0].lower()
                    count = len(soup.find_all(tag))

                # Find the row with the factor title and output the count next to it
                cell = page_structure_sheet.find(factor)
                page_structure_sheet.update_cell(cell.row, j, count)
            print(f"Page structure updated for the following url: {url}")
            time.sleep(15)

    print(f"Google Sheet Page Structure sheet updated for keyword, '{keyword}'")

print(f'\n>>> COMPLETE scrape_page_structure.py <<<')