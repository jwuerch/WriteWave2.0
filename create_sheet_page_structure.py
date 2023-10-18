import gspread
import os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from gspread_formatting import *


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

print(f'>>> START update_datasheet_page_structure.py <<<\n')
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

    # Select the 'SERP Data' worksheet and get all search results starting from row 3
    serp_data_sheet = datasheet.worksheet('SERP Data')
    search_results = serp_data_sheet.col_values(serp_data_sheet.find('Search Result').col)[2:]

    # Select the 'Page Structure' worksheet
    page_structure_sheet = datasheet.worksheet('Page Structure')

    # Define the values to be written
    values = ['Factor', 'Result 1', 'Result 2', 'Result 3', 'Result 4', 'Result 5', 'Result 6', 'Result 7', 'Result 8',
              'Result 9', 'Result 10', 'Average']

    factors = ['Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count', 'H5 Tag Count',
               'H6 Tag Count',
               'p Tag Count', 'a Tag Count', 'a Internal Tag Count', 'a External Tag Count', 'img Tag Count',
               'alt Tag Count',
               'b Tag Count', 'strong Tag Count', 'i Tag Count', 'em Tag Count', 'u Tag Count', 'ol Tag Count',
               'ol Item Count', 'ul Tag Count', 'ul Item Count', 'table Tag Count', 'form Tag Count',
               'iframe Tag Count',
               'FAQs Count', 'TOC Count']

    # Write the factors starting from row 4
    for j, factor in enumerate(factors, start=1):
        page_structure_sheet.update_cell(j + 3, 1, factor)

    # Write the values starting from row 3
    for j, value in enumerate(values, start=1):
        page_structure_sheet.update_cell(3, j, value)

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt = CellFormat(wrapStrategy='CLIP')

    # Get the number of rows and columns in the worksheet
    num_rows = len(page_structure_sheet.get_all_values())
    num_cols = len(page_structure_sheet.get_all_values()[0])

    # Apply the formatting to all cells in the worksheet
    format_cell_range(page_structure_sheet, f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt)

    # Create a CellFormat object for bold text
    bold_fmt = CellFormat(textFormat=TextFormat(bold=True))

    # Apply bold formatting to row 3
    format_cell_range(page_structure_sheet, f'A3:{gspread.utils.rowcol_to_a1(3, num_cols)}', bold_fmt)

    # Double the width of column A
    set_column_width(page_structure_sheet, 'A:A', 150)  # 100 is the standard width


    print(f"Google Sheet Page Structure sheet updated for keyword, '{keyword}'")

print(f'\n>>> COMPLETE update_datasheet_page_structure.py <<<')