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

    # Select the 'Entities' worksheet
    entities_sheet = keyword_sheet.worksheet('Entities')

    # Append URLs to 'Entities' worksheet starting from B3
    for i, url in enumerate(urls, start=3):
        entities_sheet.update_cell(i, 2, url)  # Column B is index 2

    # Extract the details for each entity and prepare for Google Sheets
    entities_list = [entity.id for entity in entities_dict.values()]
    confidence_scores = [entity.confidence_score for entity in entities_dict.values()]
    relevance_scores = [entity.relevance_score for entity in entities_dict.values()]
    types = [", ".join(entity.freebase_types) for entity in entities_dict.values()]  # Join multiple types into a single string
    freebase_ids = [entity.freebase_id for entity in entities_dict.values()]
    wikidata_ids = [entity.wikidata_id for entity in entities_dict.values()]
    wikilinks = [entity.wikipedia_link for entity in entities_dict.values()]
    wikilink_tags = [f'<a href="{wikilink}" title="{entity.id}" rel="nofollow">{entity.id}</a>'
                     for entity, wikilink in zip(entities_dict.values(), wikilinks) if wikilink]

    # Calculate the end column letter based on the number of entities
    end_column_letter = get_column_letter(len(entities_list) + 3)  # +3 because we start from column D (index 4)

    # Update the cells in the 'Entities' worksheet with the entities and their details
    entities_sheet.update(values=[clean_data_for_sheet(entities_list)], range_name=f'D2:{end_column_letter}2')
    entities_sheet.update(values=[clean_data_for_sheet(confidence_scores)], range_name=f'D13:{end_column_letter}13')
    entities_sheet.update(values=[clean_data_for_sheet(relevance_scores)], range_name=f'D14:{end_column_letter}14')
    entities_sheet.update(values=[clean_data_for_sheet(types)], range_name=f'D15:{end_column_letter}15')
    entities_sheet.update(values=[clean_data_for_sheet(freebase_ids)], range_name=f'D16:{end_column_letter}16')
    entities_sheet.update(values=[clean_data_for_sheet(wikidata_ids)], range_name=f'D17:{end_column_letter}17')
    entities_sheet.update(values=[clean_data_for_sheet(wikilinks)], range_name=f'D18:{end_column_letter}18')
    entities_sheet.update(values=[clean_data_for_sheet(wikilink_tags)], range_name=f'D19:{end_column_letter}19')