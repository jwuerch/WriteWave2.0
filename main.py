import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from sheet_creation.create_google_sheet import *
from sheet_creation.create_worksheet_serp_data import *

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))

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
    print(f"\n>>> START keyword, '{keyword}'")

    # Check if the 'Datasheet' cell is empty
    if keyword_sheet_link:
        print(f"Google Sheet already defined for keyword, '{keyword}'")
        continue

    # Create new Google Sheet from create_google_sheet.py
    print(f"Creating Google Sheet [create_google_sheet.py]")
    keyword_sheet_url = create_google_sheet(client, keyword, data_folder_id, main_sheet, google_sheet_col_index, i)

    # Create 'SERP Data' worksheet from create_worksheet_serp_data.py
    print(f"Creating SERP Data worksheet")
    create_worksheet_serp_data(client, keyword_sheet_url)
