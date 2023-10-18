import gspread
import os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from serpapi import GoogleSearch
import time
from collections import defaultdict

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

    # Initialize a dictionary to store the rankings
    rankings = defaultdict(list)

    # Initialize a set to store the unique keyword variations
    keyword_variations = set()

    # Perform 5 searches and calculate the average ranking
    for _ in range(5):
        # Use SerpAPI to get the top 100 search results for the keyword
        params = {
            "engine": "google",
            "q": keyword,
            "api_key": serpapi_api_key,
            "device": "desktop",
            "location": "United States",
            "num": 100
        }
        search = GoogleSearch(params)
        results = search.get_dict()

        # Check if 'organic_results' is in the results
        if 'organic_results' in results:
            # Extract the URLs and their rankings
            for rank, result in enumerate(results['organic_results'][:100], start=1):
                # Add the 'snippet_highlighted_words' to the keyword variations set
                if 'snippet_highlighted_words' in result:
                    keyword_variations.update(result['snippet_highlighted_words'])

                # Only process the other data for the top 10 results
                if rank <= 10:
                    url = result['link']
                    rankings[url].append(rank)

        # Add a 30-second delay between each search
        print(f'Successfully completed a search.')
        time.sleep(10)

    # Calculate the average ranking for each URL
    avg_rankings = {url: sum(ranks) / len(ranks) for url, ranks in rankings.items()}

    # Get the index of '#' column
    rank_col_index = serp_header_row.index('#')

    # Update the '#' column with the average rankings
    for j, (url, avg_rank) in enumerate(avg_rankings.items(), start=3):
        serp_worksheet.update_cell(j, rank_col_index + 1, avg_rank)

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
        # Extract the People Also Ask questions and answers
        pas_questions = [pas['question'] for pas in results['related_questions']]
        pas_answers = [pas.get('snippet') for pas in results['related_questions']]  # Use get method here

        # Get the index of 'People Also Ask' and 'Answer' columns
        people_also_ask_col_index = serp_header_row.index('People Also Ask')
        answer_col_index = serp_header_row.index('Answer')

        # Update the 'People Also Ask' and 'Answer' columns with the questions and answers
        for j, (question, answer) in enumerate(zip(pas_questions, pas_answers), start=3):
            serp_worksheet.update_cell(j, people_also_ask_col_index + 1, question)
            serp_worksheet.update_cell(j, answer_col_index + 1, answer)

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
        print(f"No organic video results found for keyword, '{keyword}'")

    # Get the index of 'Keyword Variations' column
    keyword_variations_col_index = serp_header_row.index('Keyword Variations')
    time.sleep(50)
    # Update the 'Keyword Variations' column with the unique keyword variations
    for j, variation in enumerate(keyword_variations, start=3):
        serp_worksheet.update_cell(j, keyword_variations_col_index + 1, variation.lower())

        # Add a delay every 90 variations
        if (j - 2) % 90 == 0:  # Subtract 2 from j because j starts from 3
            time.sleep(61)

    print(f"Fetched the following SERP data for the keyword, '{keyword}':\n-Ranking\n-Top 10 Search Results\n-SEO Titles\n-Meta Descriptions\n-People Also Ask Quetions\n-People Also Ask Answers\n-Videos\n-Featured Snippet\n-Keyword Variations")
    time.sleep(5)

print(f'\n>>> COMPLETE fetch_serp_data.py <<<')