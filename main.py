import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from sheet_creation.create_google_sheet import *
from sheet_creation.create_worksheet_serp_data import *
from sheet_creation.create_worksheet_entities import *
from sheet_creation.create_worksheet_variations import *
from sheet_creation.create_worksheet_page_structure import *
from scrape_page_structure import *
from fetch_serp_data import *
from find_entities import *

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))
serpapi_api_key = os.getenv('SERPAPI_API_KEY')
textrazor_api_key = os.getenv('TEXTRAZOR_API_KEY_TWO')

# Use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/documents']
creds = Credentials.from_service_account_file(credentials, scopes=scope)
client = gspread.authorize(creds)

# Build the Drive v3 API service
drive_service = build('drive', 'v3', credentials=creds)

# The ID of the 'Data' folder
data_folder_id = '1-hE7be9CdevebuhCeBtAICprqTEgZdHV'

# Find a workbook by name and open the first sheet
main_sheet = client.open("WriteWave2.0").sheet1

# Extract all the values
all_values = main_sheet.get_all_values()

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
keyword_col_index = header_row.index('Keyword')
google_sheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    keyword = row[keyword_col_index]
    keyword_sheet_link = row[google_sheet_col_index]
    print(f"\n>>> START keyword, '{keyword}'\n")

    # Check if the 'Datasheet' cell is empty
    if keyword_sheet_link:
        print(f"Google Sheet already defined for keyword, '{keyword}' [create_google_sheet.py]")
        keyword_sheet_url = keyword_sheet_link
        keyword_sheet = client.open_by_url(keyword_sheet_url)
    else:
        # Create new Google Sheet from create_google_sheet.py
        print(f"Creating Google Sheet [create_google_sheet.py]")
        keyword_sheet_url = create_google_sheet(client, keyword, data_folder_id, main_sheet, google_sheet_col_index, i)
        keyword_sheet = client.open_by_url(keyword_sheet_url)

        # Create 'SERP Data' worksheet from create_worksheet_serp_data.py
        print(f"Creating SERP Data worksheet [create_worksheet_serp_data.py]")
        create_worksheet_serp_data(keyword_sheet)

        # Create 'Entities' worksheet from create_worksheet_entities.py
        print(f"Creating Entities worksheet [create_worksheet_entities.py]")
        create_worksheet_entities(keyword_sheet)

        # Create 'Variations' worksheet from create_worksheet_variations.py
        print(f"Creating Variations worksheet [create_worksheet_variations.py]")
        create_worksheet_variations(keyword_sheet)

        # Create 'Page Structure' worksheet from create_worksheet_page_structure.py
        print(f"Creating Page Structure worksheet [create_worksheet_page_structure.py]")
        create_worksheet_page_structure(keyword_sheet)

    # Fetch SERP data
    print(f"Fetching SERP data [fetch_serp_data.py]")
    # fetch_serp_data(keyword_sheet, keyword, serpapi_api_key)

    # Find Entities
    print(f"Scanning for entities [find_entities.py]")
    # find_entities(keyword_sheet, textrazor_api_key)

    # Scrape URLs
    print(f"Scraping web pages for structure [scrape_page_structure.py")
    scrape_page_structure(keyword_sheet)

    print(f"\n>>> END keyword, '{keyword}'\n")
