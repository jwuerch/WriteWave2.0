import gspread
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from gspread_formatting import *

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
header_row = all_values[1]
keyword_col_index = header_row.index('Keyword')

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
datasheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

print(f'>>> START create_sheet_keyword_variations.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    datasheet_link = row[datasheet_col_index]
    keyword = row[keyword_col_index]

    # Check if the 'Datasheet' cell is empty
    if not datasheet_link:
        print(f"No Google Sheet defined for row {i+1}")
        continue

    # Open the Google Sheet and get the 'Keyword Variations' worksheet
    datasheet = client.open_by_url(datasheet_link)
    entities_worksheet = datasheet.worksheet('Entities')

    # Define the data to be pasted
    data = [
        ['Result 1'],
        ['Result 2'],
        ['Result 3'],
        ['Result 4'],
        ['Result 5'],
        ['Result 6'],
        ['Result 7'],
        ['Result 8'],
        ['Result 9'],
        ['Result 10'],
        ['Page 1 Average'],
        ['Page 1 Maximum']
    ]

    # Create a list of updates
    updates = [
        {
            'range': 'C2',
            'values': [['# used']]
        },
        {
            'range': f'A3:A{len(data) + 3}',
            'values': data
        }
    ]

    # Apply all updates in a single batch
    entities_worksheet.batch_update(updates)

    # Get the total number of rows and columns in the worksheet
    num_rows = entities_worksheet.row_count
    num_cols = entities_worksheet.col_count

    # Create a CellFormat object with bold text and text wrapping set to 'CLIP'
    fmt = cellFormat(
        textFormat=textFormat(bold=True),
        wrapStrategy='CLIP'
    )

    # Apply the formatting to all cells in the worksheet
    format_cell_range(entities_worksheet, f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt)

    # Expand the width of column A
    set_column_width(entities_worksheet, 'A:A', 150)  # 100 is the standard width

    print(f"Updated 'Keyword Variations' for row {i+1}")

    print(f"Create Entities worksheet for keyword, '{keyword}'")

print(f'\n>>> COMPLETE create_sheet_keyword_variations.py <<<')