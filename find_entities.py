import gspread
import os
import textrazor
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))
textrazor.api_key = os.getenv('TEXTRAZOR_API_KEY')

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


def get_column_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

print(f'>>> START find_entities.py <<<\n')
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

    # Select the 'SERP Data' worksheet
    serp_data_sheet = datasheet.worksheet('SERP Data')

    # Get the top 10 URLs from the 'Search Result' column
    search_result_col = serp_data_sheet.find('Search Result').col
    urls = serp_data_sheet.col_values(search_result_col)[2:12]  # Skip the header row and get the next 10 rows

    # Process each URL with TextRazor
    entities_dict = {}  # To store unique entities
    for url in urls:
        # Add 'http://' prefix if not present
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url

        try:
            response = textrazor.TextRazor(extractors=["entities"]).analyze_url(url)

            for entity in response.entities():
                entities_dict[entity.id] = entity  # Add entity to dict (automatically handles duplicates)

        except textrazor.TextRazorAnalysisException as e:
            print(f"Failed to analyze URL '{url}': {e}")
    print(entities_dict)
    print(len(entities_dict))

    # Select the 'Entities' worksheet
    entities_sheet = datasheet.worksheet('Entities')

    # Append URLs to 'Entities' worksheet starting from B3
    for i, url in enumerate(urls, start=3):
        entities_sheet.update_cell(i, 2, url)  # Column B is index 2

    # Convert the entities dictionary to a list of entity IDs
    entities_list = [entity.id for entity in entities_dict.values()]

    # Prepare the data for the update
    # This will create a list of lists where each inner list is a row
    # Since we want to print the entities horizontally, we only need one row
    data = [entities_list]

    # Calculate the end column letter based on the number of entities
    end_column_letter = get_column_letter(len(entities_list) + 3)  # +3 because we start from column D (index 4)

    # Update the cells in the 'Entities' worksheet starting from D2
    entities_sheet.update(f'D2:{end_column_letter}2', data)


print(f'\n>>> COMPLETE find_entities.py <<<')