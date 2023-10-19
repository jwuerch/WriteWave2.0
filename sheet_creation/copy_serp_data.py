import gspread
import os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

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
datasheet_col_index = header_row.index('Google Sheet')

# Adjust the range to 0-indexed
start_range -= 1
end_range -= 1

def get_column_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

print(f'>>> START copy_keyword_variations.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    datasheet_link = row[datasheet_col_index]

    # Check if the 'Datasheet' cell is empty
    if not datasheet_link:
        print(f"No Google Sheet defined for row {i+1}")
        continue

    # Open the Google Sheet from the 'Datasheet' link
    datasheet_id = datasheet_link.split('/')[5]
    datasheet = client.open_by_key(datasheet_id)

    # Open the 'SERP Data' and 'Keyword Variations' worksheets
    serp_worksheet = datasheet.worksheet('SERP Data')
    keyword_variations_worksheet = datasheet.worksheet('Keyword Variations')

    # Get the index of 'Keyword Variations' column
    serp_header_row = serp_worksheet.row_values(2)
    keyword_variations_col_index = serp_header_row.index('Keyword Variations')

    # Get all the keyword variations
    keyword_variations = serp_worksheet.col_values(keyword_variations_col_index + 1)[2:]  # Skip the header row and the empty row

    # Prepare the data for the update
    # This will create a list of lists where each inner list is a row
    keyword_variations_data = [keyword_variations]

    # Calculate the end column letter based on the number of keyword variations
    end_column_letter = get_column_letter(len(keyword_variations) + 3)  # +3 because we start from column D (index 4)

    # Update the cells in the 'Keyword Variations' worksheet starting from D2
    keyword_variations_worksheet.update(f'D2:{end_column_letter}2', keyword_variations_data)

    print(f"Copied all keyword variations for row {i + 1}")

print(f'\n>>> COMPLETE copy_keyword_variations.py <<<')