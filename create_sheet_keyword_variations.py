import gspread
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv

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

# Find a workbook by name and open the first sheet
sheet = client.open("WriteWave2.0").sheet1

# Extract all the values
all_values = sheet.get_all_values()

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
datasheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

print(f'>>> START update_keyword_variations.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    datasheet_link = row[datasheet_col_index]

    # Check if the 'Datasheet' cell is empty
    if not datasheet_link:
        print(f"No Google Sheet defined for row {i+1}")
        continue

    # Open the Google Sheet and get the 'Keyword Variations' worksheet
    datasheet = client.open_by_url(datasheet_link)
    keyword_variations = datasheet.worksheet('Keyword Variations')

    # Write '# used' in cell C1
    keyword_variations.update('C1', '# used')

    print(f"Updated 'Keyword Variations' for row {i+1}")

print(f'\n>>> COMPLETE update_keyword_variations.py <<<')