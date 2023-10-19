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

print(f'>>> START create_sheet_page_structure.py <<<\n')
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
              'Result 9', 'Result 10', 'Page 1 Average', 'Page 1 Maximum']

    factors = ['Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count', 'H5 Tag Count',
               'H6 Tag Count',
               'p Tag Count', 'a Tag Count', 'a Internal Tag Count', 'a External Tag Count', 'img Tag Count',
               'alt Tag Count',
               'b Tag Count', 'strong Tag Count', 'i Tag Count', 'em Tag Count', 'u Tag Count', 'ol Tag Count',
               'ol Item Count', 'ul Tag Count', 'ul Item Count', 'table Tag Count', 'form Tag Count',
               'iframe Tag Count',
               'FAQs Count', 'TOC Count']

    # Write the factors starting from row 4
    factors_range = 'A4:A' + str(len(factors) + 3)
    page_structure_sheet.update(factors_range, [[factor] for factor in factors])

    # Write the values starting from row 3
    values_range = 'A3:' + gspread.utils.rowcol_to_a1(3, len(values))
    page_structure_sheet.update(values_range, [values])

    # Get the number of rows and columns in the worksheet
    all_values = page_structure_sheet.get_all_values()
    num_rows = len(all_values)
    num_cols = len(all_values[0])

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt = CellFormat(wrapStrategy='CLIP')

    # Create a CellFormat object for bold text
    bold_fmt = CellFormat(textFormat=TextFormat(bold=True))

    # Apply the formatting to all cells in the worksheet and bold formatting to row 3
    ranges_fmt = [
        (f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt),
        (f'A3:{gspread.utils.rowcol_to_a1(3, num_cols)}', bold_fmt)
    ]
    format_cell_ranges(page_structure_sheet, ranges_fmt)

    # Double the width of column A
    set_column_width(page_structure_sheet, 'A:A', 150)  # 100 is the standard width

    print(f"Google Sheet Page Structure sheet updated for keyword, '{keyword}'")

print(f'\n>>> COMPLETE create_sheet_page_structure.py <<<')