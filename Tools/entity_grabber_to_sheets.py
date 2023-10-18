import gspread
import os
import textrazor
import time
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from gspread_formatting import *

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

    # Select the 'SERP Data' and 'Entities' worksheets
    serp_data_sheet = datasheet.worksheet('SERP Data')
    entities_sheet = datasheet.worksheet('Entities')

    # Get the top 10 URLs from the 'Search Result' column
    search_result_col = serp_data_sheet.find('Search Result').col
    urls = serp_data_sheet.col_values(search_result_col)[2:12]  # Skip the header row and get the next 10 rows

    entity_count = 0

    # Process each URL with TextRazor
    for url in urls:
        # Add 'http://' prefix if not present
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url

        try:
            response = textrazor.TextRazor(extractors=["entities"]).analyze_url(url)

            # Create a dictionary to store entities and their details
            entity_dict = {}

            for entity in response.entities():
                # If entity is already in dictionary, increment count
                if entity.id in entity_dict:
                    entity_dict[entity.id][2] += 1
                # If entity is not in dictionary, add it with count 1
                else:
                    entity_dict[entity.id] = [entity.relevance_score, entity.confidence_score, 1,
                                              ', '.join(entity.freebase_types), entity.wikidata_id,
                                              entity.wikipedia_link,
                                              f'<a href="{entity.wikipedia_link}" title="{entity.id}" rel="nofollow">{entity.id}</a>',
                                              entity.freebase_id]

            # Sort entities by relevance score and append them to the Google Sheet
            for entity_id, details in sorted(entity_dict.items(), key=lambda item: item[1][0], reverse=True):
                entities_sheet.append_row([
                    url,  # Website
                    entity_id,  # Entity
                    details[2],  # Count
                    details[7],  # Freebase ID
                    details[0],  # Relevance
                    details[1],  # Confidence
                    details[3],  # Type
                    details[4],  # Wikidata ID
                    details[5],  # Wiki Link
                    details[6],  # Wiki Link Tag
                    ''  # Freebase ID (not available in TextRazor)
                ])

                entity_count += 1

                if entity_count % 60 == 0:
                    time.sleep(60)
                    print(f'Waiting 60 seconds...')

        except textrazor.TextRazorAnalysisException as e:
            print(f"Failed to analyze URL '{url}': {e}")


    print(f"Google Sheet Entities sheet updated for keyword, '{keyword}'")

print(f'\n>>> COMPLETE find_entities.py <<<')