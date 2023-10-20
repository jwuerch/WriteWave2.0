import gspread
import textrazor

def get_column_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

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
