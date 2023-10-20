from serpapi import GoogleSearch
import time
import random
from collections import defaultdict


def get_column_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def fetch_serp_data(keyword_sheet, keyword, serpapi_api_key):

    # Open the 'SERP Data' worksheet
    serp_data_worksheet = keyword_sheet.worksheet('SERP Data')
    variations_worksheet = keyword_sheet.worksheet('Variations')

    # Initialize a dictionary to store the rankings and related questions
    rankings = defaultdict(list)
    related_questions = {}
    video_links = {}
    keyword_variations = set()
    featured_snippet = None

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
            for rank, result in enumerate(results['organic_results'], start=1):
                url = result['link']
                rankings[url].append(rank)
                # Check if 'snippet_highlighted_words' is in the result
                if 'snippet_highlighted_words' in result:
                    # Add the keyword variations to the set
                    keyword_variations.update([variation.lower() for variation in result['snippet_highlighted_words']])

                # print(f"Rank: {rank}, URL: {url}")  # Print to console

        # Check if 'related_questions' is in the results
        if 'related_questions' in results:
            # Extract the questions and their answers
            for question in results['related_questions']:
                q = question['question']
                a = question.get('snippet', '')
                if q not in related_questions:
                    related_questions[q] = a

        # Check if 'inline_videos' is in the results
        if 'inline_videos' in results:
            # Extract the video links
            for video in results['inline_videos']:
                link = video['link']
                if link not in video_links:
                    video_links[link] = True

        # Check if 'answer_box' is in the results and featured_snippet is not yet set
        if 'answer_box' in results and not featured_snippet:
            # Extract the featured snippet
            featured_snippet = results['answer_box']['snippet']

        # Add a 30-second delay between each search
        random_number = random.uniform(7, 12)
        print(f"Successfully completed a search. Waiting {random_number} seconds...")
        time.sleep(random_number)

    # Calculate the average ranking for each URL
    avg_rankings = {url: sum(ranks) / len(ranks) for url, ranks in rankings.items()}

    # Prepare data for batch update
    updates = []

    updates.append({
        'range': f'A3:A{len(avg_rankings)+2}',
        'values': [[avg_rank] for avg_rank in avg_rankings.values()]
    })

    # Extract the URLs of all results
    urls = [result['link'] for result in results['organic_results']]

    updates.append({
        'range': f'B3:B{len(urls) + 2}',
        'values': [[url] for url in urls]
    })

    # Extract the SEO Titles and Meta Descriptions of all results
    titles = [result['title'] for result in results['organic_results']]
    descriptions = [result.get('snippet', '') for result in results['organic_results']]

    updates.append({
        'range': f'C3:C{len(titles) + 2}',
        'values': [[title] for title in titles]
    })
    updates.append({
        'range': f'D3:D{len(descriptions) + 2}',
        'values': [[description] for description in descriptions]
    })

    # Add related questions and their answers
    questions = list(related_questions.keys())
    answers = list(related_questions.values())

    updates.append({
        'range': f'E3:E{len(questions) + 2}',
        'values': [[question] for question in questions]
    })
    updates.append({
        'range': f'F3:F{len(answers) + 2}',
        'values': [[answer] for answer in answers]
    })

    # Add video links
    links = list(video_links.keys())

    updates.append({
        'range': f'G3:G{len(links) + 2}',
        'values': [[link] for link in links]
    })

    # Add featured snippet
    updates.append({
        'range': 'H3',
        'values': [[featured_snippet]]
    })

    # Add keyword variations
    variations = list(keyword_variations)

    updates.append({
        'range': f'I3:I{len(variations) + 2}',
        'values': [[variation] for variation in variations]
    })

    # Update the worksheet with all the data at once
    serp_data_worksheet.batch_update(updates)

    # Get all the keyword variations
    keyword_variations = serp_data_worksheet.col_values(serp_data_worksheet.row_values(2).index('Keyword Variations') + 1)[2:]

    # Prepare the data for the update. This will create a list of lists where each inner list is a row
    keyword_variations_data = [keyword_variations]

    # Calculate the end column letter based on the number of keyword variations
    end_column_letter = get_column_letter(len(keyword_variations) + 3)  # +3 because we start from column D (index 4)

    # Update the cells in the 'Keyword Variations' worksheet starting from D2
    variations_worksheet.update(f'D2:{end_column_letter}2', keyword_variations_data)

