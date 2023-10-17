import gspread
from google.oauth2.service_account import Credentials
from serpapi import GoogleSearch
import time
from dotenv import load_dotenv
import os

# Load the .env file
load_dotenv()

credentials = os.getenv('CREDENTIALS')
start_range = int(os.getenv('START_RANGE'))
end_range = int(os.getenv('END_RANGE'))
serpapi_api_key = os.getenv('SERPAPI_API_KEY')

# Use gspread to open the Google Sheet
gc = gspread.service_account(filename=credentials)
sh = gc.open('WriteWave')
worksheet = sh.get_worksheet(1)  # 0 for the first sheet, 1 for the second sheet etc.

# Find the "Keyword", "Search Results", "SEO Titles", "Meta Descriptions", "PAS", "Featured Snippet" and "Videos" columns
header_row = worksheet.row_values(2)
keyword_col = header_row.index('Keyword') + 1
results_col = header_row.index('Search Results') + 1
titles_col = header_row.index('SEO Titles') + 1
desc_col = header_row.index('Meta Descriptions') + 1
pas_col = header_row.index('PAS') + 1
snippet_col = header_row.index('Featured Snippet') + 1
video_col = header_row.index('Videos') + 1

# Loop through the range
for i in range(start_range, end_range + 1):
    # Get the keyword from the "Keyword" column
    keyword = worksheet.cell(i, keyword_col).value

    # Use SerpAPI to get the top 10 search results
    params = {
        "engine": "google",
        "q": keyword,
        "api_key": serpapi_api_key,
        "device": "mobile",
        "location": "United States",
        "num": 10
    }

    search = GoogleSearch(params)
    results = search.get_dict()

    # Check if 'organic_results' is in the results
    if 'organic_results' in results:
        # Extract the URLs, titles, and meta descriptions of the top 10 results
        urls = [result['link'] for result in results['organic_results'][:10]]
        titles = [result['title'] for result in results['organic_results'][:10]]
        descriptions = [result.get('snippet', '') for result in results['organic_results'][:10]]

        # Write the URLs to the "Search Results" column, the titles to the "SEO Titles" column, and the descriptions to the "Meta Descriptions" column
        worksheet.update_cell(i, results_col, '\n'.join(urls))
        worksheet.update_cell(i, titles_col, '\n'.join(titles))
        worksheet.update_cell(i, desc_col, '\n'.join(descriptions))

    # Check if 'related_questions' is in the results
    if 'related_questions' in results:
        # Extract the questions from 'related_questions'
        pas_questions = [pas['question'] for pas in results['related_questions']]

        # Write the questions to the "PAS" column
        worksheet.update_cell(i, pas_col, '\n'.join(pas_questions))

    # Check if 'answer_box' is in the results
    if 'answer_box' in results:
        # Extract the answer from 'answer_box'
        featured_snippet = results['answer_box']['snippet']

        # Write the answer to the "Featured Snippet" column
        worksheet.update_cell(i, snippet_col, featured_snippet)

    # Check if 'video_results' is in the results
    if 'video_results' in results:
        # Extract the URLs of the top 5 video results
        video_urls = [video['link'] for video in results['video_results'][:5]]

        # Write the URLs to the "Videos" column
        worksheet.update_cell(i, video_col, '\n'.join(video_urls))

    # To prevent hitting the rate limit
    time.sleep(1)
    print(f"SERP Data grabbed for keyword, '{keyword}'")