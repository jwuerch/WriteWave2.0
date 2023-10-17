import gspread
import os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from serpapi import GoogleSearch

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))
serpapi_api_key = os.getenv('SERPAPI_API_KEY')

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

print(f'>>> START fetch_serp_data.py <<<\n')
# Process only rows within the defined range
for i, row in enumerate(all_values[start_range:end_range + 1], start=start_range):
    keyword = row[keyword_col_index]
    datasheet_link = row[datasheet_col_index]

    # Check if the 'Datasheet' cell is empty
    if not datasheet_link:
        print(f"No Google Sheet defined for keyword, '{keyword}'")
        continue

    # Open the Google Sheet from the 'Datasheet' link
    datasheet_id = datasheet_link.split('/')[5]
    datasheet = client.open_by_key(datasheet_id)

    # Open the 'SERP Data' worksheet
    serp_worksheet = datasheet.worksheet('SERP Data')

    # Get the index of 'Search Result' column
    serp_header_row = serp_worksheet.row_values(2)
    search_result_col_index = serp_header_row.index('Search Result')

    # Use SerpAPI to get the top 10 search results for the keyword
    params = {
        "engine": "google",
        "q": keyword,
        "api_key": serpapi_api_key,
        "device": "desktop",
        "location": "United States",
        "num": 10
    }
    search = GoogleSearch(params)
    results = search.get_dict()
    print(results)

    # Check if 'organic_results' is in the results
    if 'organic_results' in results:
        # Extract the URLs, titles, and meta descriptions of the top 10 results
        urls = [result['link'] for result in results['organic_results'][:10]]
        titles = [result['title'] for result in results['organic_results'][:10]]
        descriptions = [result.get('snippet', '') for result in results['organic_results'][:10]]

        # Get the index of 'SEO Title' and 'Meta Description' columns
        seo_title_col_index = serp_header_row.index('SEO Title')
        meta_description_col_index = serp_header_row.index('Meta Description')

        # Update the 'Search Result', 'SEO Title', and 'Meta Description' columns with the URLs, titles, and descriptions
        for j, (url, title, description) in enumerate(zip(urls, titles, descriptions), start=3):
            serp_worksheet.update_cell(j, search_result_col_index + 1, url)
            serp_worksheet.update_cell(j, seo_title_col_index + 1, title)
            serp_worksheet.update_cell(j, meta_description_col_index + 1, description)

        # Check if 'people_also_ask' is in the results
        if 'related_questions' in results:
            # Extract the People Also Ask questions
            pas_questions = [pas['question'] for pas in results['related_questions']]

            # Get the index of 'People Also Ask' column
            people_also_ask_col_index = serp_header_row.index('People Also Ask')

            # Update the 'People Also Ask' column with the questions
            for j, question in enumerate(pas_questions, start=3):
                serp_worksheet.update_cell(j, people_also_ask_col_index + 1, question)

        # Check if 'answer_box' is in the results
        if 'answer_box' in results:
            # Extract the answer from 'answer_box'
            featured_snippet = results['answer_box']['snippet']

            # Get the index of 'Featured Snippet' column
            featured_snippet_col_index = serp_header_row.index('Featured Snippet')

            # Update the 'Featured Snippet' column with the Featured Snippet
            serp_worksheet.update_cell(3, featured_snippet_col_index + 1, featured_snippet)

        # Check if 'inline_videos' is in the results
        if 'inline_videos' in results:
            # Extract the URLs of the top 10 video results
            video_urls = [video['link'] for video in results['inline_videos'][:10]]

            # Get the index of 'Videos' column
            video_col_index = serp_header_row.index('Video')

            # Write each URL to its own cell in the "Videos" column
            for j, url in enumerate(video_urls, start=3):
                serp_worksheet.update_cell(j, video_col_index + 1, url)
    else:
        print(f"No organic results found for keyword, '{keyword}'")

    print(f"Updated Google Sheet for keyword, '{keyword}'")

print(f'\n>>> COMPLETE fetch_serp_data.py <<<')