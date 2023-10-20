import gspread
import textrazor

def get_column_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def clean_data_for_sheet(data_list):
    return [item if item is not None else "" for item in data_list]

def find_entities(keyword_sheet, textrazor_api_key):

    textrazor.api_key = textrazor_api_key

    # Select the 'SERP Data' worksheet
    serp_data_sheet = keyword_sheet.worksheet('SERP Data')

    # Get the top 10 URLs from the 'Search Result' column
    search_result_col = serp_data_sheet.find('Search Result').col
    urls = serp_data_sheet.col_values(search_result_col)[2:12]  # Skip the header row and get the next 10 rows

    # Process each URL with TextRazor
    entities_dict = {}  # To store unique entities with Wikipedia links
    confidence_scores_dict = {}  # To store confidence scores for each entity
    relevance_scores_dict = {}  # To store relevance scores for each entity
    entity_counts_dict = {}  # To store entity counts for each URL

    total_counts_dict = {}  # To store total counts for each entity
    count_occurrences_dict = {}  # To store count occurrences for each entity

    for url in urls:
        # Add 'http://' prefix if not present
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url

        try:
            response = textrazor.TextRazor(extractors=["entities"]).analyze_url(url)

            for entity in response.entities():
                # Check if the entity has a Wikipedia link before adding it to the dict
                if entity.wikipedia_link:
                    entities_dict[entity.id] = entity  # Add entity to dict (automatically handles duplicates)
                    # Add confidence and relevance scores to their respective dicts
                    if entity.id in confidence_scores_dict:
                        confidence_scores_dict[entity.id].append(entity.confidence_score)
                    else:
                        confidence_scores_dict[entity.id] = [entity.confidence_score]
                    if entity.id in relevance_scores_dict:
                        relevance_scores_dict[entity.id].append(entity.relevance_score)
                    else:
                        relevance_scores_dict[entity.id] = [entity.relevance_score]
                    # Count the entity for the current URL
                    if url in entity_counts_dict:
                        if entity.id in entity_counts_dict[url]:
                            entity_counts_dict[url][entity.id] += 1
                        else:
                            entity_counts_dict[url][entity.id] = 1
                    else:
                        entity_counts_dict[url] = {entity.id: 1}

                    # Add the count to the total counts dict
                    if entity.id in total_counts_dict:
                        total_counts_dict[entity.id] += 1
                    else:
                        total_counts_dict[entity.id] = 1

                    # Add the count to the count occurrences dict
                    if entity.id in count_occurrences_dict:
                        count_occurrences_dict[entity.id] += 1
                    else:
                        count_occurrences_dict[entity.id] = 1

        except textrazor.TextRazorAnalysisException as e:
            print(f"Failed to analyze URL '{url}': {e}")

    # Calculate average confidence and relevance scores for each entity
    average_confidence_scores = [sum(scores) / len(scores) for scores in confidence_scores_dict.values()]
    average_relevance_scores = [sum(scores) / len(scores) for scores in relevance_scores_dict.values()]


    # Select the 'Entities' worksheet
    entities_worksheet = keyword_sheet.worksheet('Entities')

    # Append URLs to 'Entities' worksheet starting from B3
    for i, url in enumerate(urls, start=3):
        entities_worksheet.update_cell(i, 2, url)  # Column B is index 2

    # Initialize an empty list for updates
    updates = []

    # Extract the details for each entity and prepare for Google Sheets
    entities_list = [entity.id for entity in entities_dict.values()]
    types = [", ".join(entity.freebase_types) for entity in entities_dict.values()]  # Join multiple types into a single string
    freebase_ids = [entity.freebase_id for entity in entities_dict.values()]
    wikidata_ids = [entity.wikidata_id for entity in entities_dict.values()]
    wikilinks = [entity.wikipedia_link for entity in entities_dict.values()]
    wikilink_tags = [f'<a href="{wikilink}" title="{entity.id}" rel="nofollow">{entity.id}</a>' for entity, wikilink in zip(entities_dict.values(), wikilinks)]

    # Calculate average and maximum counts for each entity
    average_counts = [round(total_counts_dict[entity_id] / len(urls)) for entity_id in entities_list]
    max_counts = [max([entity_counts_dict.get(url, {}).get(entity_id, 0) for url in urls]) for entity_id in
                  entities_list]
    average_counts_first_3_urls = [
        round(sum([entity_counts_dict.get(url, {}).get(entity_id, 0) for url in urls[:3]]) / 3) for entity_id in
        entities_list]
    max_counts_first_3_urls = [max([entity_counts_dict.get(url, {}).get(entity_id, 0) for url in urls[:3]]) for
                               entity_id in entities_list]

    # Calculate the end column letter based on the number of entities
    end_column_letter = get_column_letter(len(entities_list) + 3)  # +3 because we start from column D (index 4)

    # Prepare the updates
    updates.append({'range': f'D2:{end_column_letter}2', 'values': [clean_data_for_sheet(entities_list)]})
    updates.append({'range': f'D17:{end_column_letter}17', 'values': [clean_data_for_sheet(average_relevance_scores)]})
    updates.append({'range': f'D18:{end_column_letter}18', 'values': [clean_data_for_sheet(average_confidence_scores)]})
    updates.append({'range': f'D19:{end_column_letter}19', 'values': [clean_data_for_sheet(types)]})
    updates.append({'range': f'D20:{end_column_letter}20', 'values': [clean_data_for_sheet(freebase_ids)]})
    updates.append({'range': f'D21:{end_column_letter}21', 'values': [clean_data_for_sheet(wikidata_ids)]})
    updates.append({'range': f'D22:{end_column_letter}22', 'values': [clean_data_for_sheet(wikilinks)]})
    updates.append({'range': f'D23:{end_column_letter}23', 'values': [clean_data_for_sheet(wikilink_tags)]})

    # Add average and maximum counts to the updates
    updates.append({'range': f'D13:{end_column_letter}13', 'values': [clean_data_for_sheet(average_counts)]})
    updates.append({'range': f'D14:{end_column_letter}14', 'values': [clean_data_for_sheet(max_counts)]})
    updates.append({'range': f'D15:{end_column_letter}15', 'values': [clean_data_for_sheet(average_counts_first_3_urls)]})
    updates.append({'range': f'D16:{end_column_letter}16', 'values': [clean_data_for_sheet(max_counts_first_3_urls)]})

    # Add entity counts to the updates
    for i, url in enumerate(urls, start=3):
        url_entity_counts = entity_counts_dict.get(url, {})
        entity_counts = [url_entity_counts.get(entity_id, 0) for entity_id in entities_list]
        updates.append({'range': f'D{i}:{end_column_letter}{i}', 'values': [clean_data_for_sheet(entity_counts)]})

    # Apply the updates to the 'Entities' worksheet
    entities_worksheet.batch_update(updates)