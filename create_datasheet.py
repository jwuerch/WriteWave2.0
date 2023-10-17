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

# The ID of the 'Data' folder
data_folder_id = '1-hE7be9CdevebuhCeBtAICprqTEgZdHV'

# Find a workbook by name and open the first sheet
sheet = client.open("WriteWave2.0").sheet1

# Extract all the values
all_values = sheet.get_all_values()

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
keyword_col_index = header_row.index('Keyword')
datasheet_col_index = header_row.index('Datasheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

print(f'>>> START create_datasheet.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    keyword = row[keyword_col_index]
    datasheet_link = row[datasheet_col_index]

    # Check if the 'Datasheet' cell is empty
    if datasheet_link:
        print(f"Google Sheet already defined for keyword, '{keyword}'")
        continue

    # Create a new Google Sheet with the keyword as the name in the 'Data' folder
    new_sheet = client.create(keyword, folder_id=data_folder_id)

    # Get the URL of the new Google Sheet
    new_sheet_url = f"https://docs.google.com/spreadsheets/d/{new_sheet.id}"

    # Update the 'Datasheet' column in the Google Sheet
    sheet.update_cell(i + 1, datasheet_col_index + 1, new_sheet_url)
    print(f"Google Sheet created for keyword, '{keyword}'")

print(f'\n>>> COMPLETE create_datasheet.py <<<')