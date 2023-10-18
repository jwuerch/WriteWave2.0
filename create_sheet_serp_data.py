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

# The ID of the 'Data' folder
data_folder_id = '1-hE7be9CdevebuhCeBtAICprqTEgZdHV'

# Find a workbook by name and open the first sheet
sheet = client.open("WriteWave2.0").sheet1

# Extract all the values
all_values = sheet.get_all_values()

# Get the index of 'Keyword' and 'Datasheet' columns
header_row = all_values[1]
keyword_col_index = header_row.index('Keyword')
datasheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

print(f'>>> START create_sheet_serp_data.py <<<\n')
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

    # Rename Sheet1 to 'SERP Overview'
    worksheet = new_sheet.get_worksheet(0)
    worksheet.update_title('SERP Data')

    # Create a new sheet titled 'Page Structure'
    new_sheet.add_worksheet(title='Page Structure', rows="1000", cols="30")
    new_sheet.add_worksheet(title='Keyword Variations', rows="1000", cols="200")
    new_sheet.add_worksheet(title='Entities', rows="1000", cols="500")
    new_sheet.add_worksheet(title='Entity Grabber', rows="1000", cols="500")
    # Update titles in the second row of 'SERP Data'
    worksheet.update('A2', '#')
    worksheet.update('B2', 'Search Result')
    worksheet.update('C2', 'SEO Title')
    worksheet.update('D2', 'Meta Description')
    worksheet.update('E2', 'People Also Ask')
    worksheet.update('F2', 'Answer')
    worksheet.update('G2', 'Video')
    worksheet.update('H2', 'Featured Snippet')
    worksheet.update('I2', 'Keyword Variations')

    # Format the second row
    fmt = cellFormat(
        backgroundColor=color(0.05, 0.33, 0.58),
        textFormat=textFormat(bold=True, fontSize=13, foregroundColor=color(1, 1, 1)),
        horizontalAlignment='CENTER',
        verticalAlignment='MIDDLE'
    )
    format_cell_range(worksheet, 'A2:I2', fmt)

    # Set column widths with some padding
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": {
                    "pixelSize": len(title) * 9 + 40  # adjust the multiplier and padding as needed
                },
                "fields": "pixelSize"
            }
        } for i, title in enumerate(['#', 'Search Result', 'SEO Title', 'Meta Description', 'People Also Ask', 'Answer', 'Video', 'Featured Snippet', 'Keyword Variations'])
    ]
    new_sheet.batch_update({"requests": requests})

    # Double the row height using Google Sheets API
    requests = [{
        "updateDimensionProperties": {
            "range": {
                "sheetId": worksheet.id,
                "dimension": "ROWS",
                "startIndex": 1,
                "endIndex": 2
            },
            "properties": {
                "pixelSize": 50
            },
            "fields": "pixelSize"
        }
    }]
    new_sheet.batch_update({"requests": requests})
    # Get the total number of rows and columns in the worksheet
    num_rows = worksheet.row_count
    num_cols = worksheet.col_count

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt = cellFormat(
        wrapStrategy='CLIP'
    )

    # Apply the formatting to all cells in the worksheet
    format_cell_range(worksheet, f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt)

    # Get the URL of the new Google Sheet
    new_sheet_url = f"https://docs.google.com/spreadsheets/d/{new_sheet.id}"

    # Update the 'Datasheet' column in the Google Sheet
    sheet.update_cell(i + 1, datasheet_col_index + 1, new_sheet_url)
    print(f"Google Sheet created for keyword, '{keyword}'")

print(f'\n>>> COMPLETE create_sheet_serp_data.py <<<')