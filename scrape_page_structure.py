from bs4 import BeautifulSoup
import requests
import time
import gspread

# Headers for requests
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
    'Accept-Encoding': 'identity'
}

def scrape_page_structure(keyword_sheet):
    serp_data_worksheet = keyword_sheet.worksheet('SERP Data')
    search_result_col = serp_data_worksheet.find('Search Result').col
    page_structure_worksheet = keyword_sheet.worksheet('Page Structure')
    urls = serp_data_worksheet.col_values(search_result_col)[2:12]  # Skip the header row and get the next 10 rows

    # Prepare a list of lists to hold the data to be written to the sheet
    data_to_write = []

    # Write URLs to the sheet in row 2, starting from column B
    for i, url in enumerate(urls, start=2):  # Start from column B
        data_to_write.append({
            'range': f'{gspread.utils.rowcol_to_a1(2, i)}',
            'values': [[url]]
        })

    # Write the URLs to the sheet in one batch
    page_structure_worksheet.batch_update(data_to_write)

    # Define the factors
    factors = ['Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count', 'H5 Tag Count', 'H6 Tag Count']

    # Clear the list for the next batch update
    data_to_write.clear()

    # For each URL, scrape the web page and update the sheet
    for url in urls:
        if url:
            response = requests.get(url, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')

            for i, factor in enumerate(factors, start=4):  # Start from row 4
                if factor == 'Word Count':
                    count = len(soup.get_text().split())
                else:
                    tag = factor.split(' ')[0].lower()
                    count = len(soup.find_all(tag))

                # Add the count to the data to be written
                url_col = urls.index(url) + 2  # +2 because URLs start from column B
                data_to_write.append({
                    'range': f'{gspread.utils.rowcol_to_a1(i, url_col)}',
                    'values': [[count]]
                })

            print(f"Page structure updated for the following url: {url}")
            time.sleep(2)

    # Write the data to the sheet in one batch
    page_structure_worksheet.batch_update(data_to_write)