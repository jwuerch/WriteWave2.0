from serpapi import GoogleSearch
import time
from collections import defaultdict
import random
import time


def fetch_serp_data(client, keyword_sheet_url, keyword, serpapi_api_key):

    # Open the Google Sheet from the 'Datasheet' link
    datasheet_id = keyword_sheet_url.split('/')[5]
    datasheet = client.open_by_key(datasheet_id)

    # Open the 'SERP Data' worksheet
    serp_worksheet = datasheet.worksheet('SERP Data')

    # Initialize a dictionary to store the rankings
    rankings = defaultdict(list)

    # Initialize a set to store the unique keyword variations
    keyword_variations = set()

    # Initialize variables to store the results
    pas_questions = []
    pas_answers = []
    featured_snippet = None
    video_urls = []

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

        # Check if 'people_also_ask' is in the results
        if 'related_questions' in results and not pas_questions:
            # Extract the People Also Ask questions and answers
            pas_questions = [pas['question'] for pas in results['related_questions']]
            pas_answers = [pas.get('snippet') for pas in results['related_questions']]  # Use get method here

        # Check if 'answer_box' is in the results
        if 'answer_box' in results and not featured_snippet:
            # Extract the answer from 'answer_box'
            featured_snippet = results['answer_box']['snippet']

        # Check if 'inline_videos' is in the results
        if 'inline_videos' in results and not video_urls:
            # Extract the URLs of the top 10 video results
            video_urls = [video['link'] for video in results['inline_videos'][:10]]

        # Add a 30-second delay between each search
        print(f'Successfully completed a search.')

        # Generate a random number between 5 and 10
        sleep_time = random.uniform(5, 7)
        time.sleep(sleep_time)

    # Calculate the average ranking for each URL
    avg_rankings = {url: sum(ranks) / len(ranks) for url, ranks in rankings.items()}

    # Prepare data for batch update
    updates = []

    updates.append({
        'range': f'A3:A{len(avg_rankings)+2}',
        'values': [[avg_rank] for avg_rank in avg_rankings.values()]
    })

    # Extract the URLs, titles, and meta descriptions of the top 10 results
    urls = [result['link'] for result in results['organic_results'][:10]]
    titles = [result['title'] for result in results['organic_results'][:10]]
    descriptions = [result.get('snippet', '') for result in results['organic_results'][:10]]

    updates.append({
        'range': f'B3:B{len(urls)+2}',
        'values': [[url] for url in urls]
    })
    updates.append({
        'range': f'C3:C{len(titles)+2}',
        'values': [[title] for title in titles]
    })
    updates.append({
        'range': f'D3:D{len(descriptions)+2}',
        'values': [[description] for description in descriptions]
    })

    if 'related_questions' in results:
        updates.append({
            'range': f'E3:E{len(pas_questions)+2}',
            'values': [[question] for question in pas_questions]
        })
        updates.append({
            'range': f'F3:F{len(pas_answers)+2}',
            'values': [[answer] for answer in pas_answers]
        })

    if 'answer_box' in results:
        updates.append({
            'range': 'H3:H4',
            'values': [[featured_snippet]]
        })

    if 'inline_videos' in results:
        updates.append({
            'range': f'G3:G{len(video_urls)+2}',
            'values': [[url] for url in video_urls]
        })
    else:
        print(f"No organic video results found for keyword, '{keyword}'")

    updates.append({
        'range': f'I3:I{len(keyword_variations) + 2}',
        'values': [[variation.lower()] for variation in keyword_variations]
    })

    # Update the worksheet with all the data at once
    serp_worksheet.batch_update(updates)