import gspread
from gspread_formatting import *

def create_worksheet_page_structure(client, keyword_sheet_url):
    # Open the Google Sheet
    keyword_sheet = client.open_by_url(keyword_sheet_url)

    # Check if 'Page Structure' worksheet exists and create if not
    try:
        page_structure_worksheet = keyword_sheet.worksheet('Page Structure')  # attempt to get the worksheet
    except gspread.WorksheetNotFound:
        # if worksheet does not exist, create a new one
        page_structure_worksheet = keyword_sheet.add_worksheet(title='Page Structure', rows="1000", cols="30")

    # Define the values to be written
    values = ['Factor', 'Result 1', 'Result 2', 'Result 3', 'Result 4', 'Result 5', 'Result 6', 'Result 7', 'Result 8',
              'Result 9', 'Result 10', 'Page 1 Average', 'Page 1 Maximum']

    factors = ['Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count', 'H5 Tag Count',
               'H6 Tag Count', 'p Tag Count', 'a Tag Count', 'a Internal Tag Count', 'a External Tag Count',
               'img Tag Count', 'alt Tag Count', 'b Tag Count', 'strong Tag Count', 'i Tag Count', 'em Tag Count',
               'u Tag Count', 'ol Tag Count', 'ol Item Count', 'ul Tag Count', 'ul Item Count', 'table Tag Count',
               'form Tag Count', 'iframe Tag Count', 'FAQs Count', 'TOC Count']

    # Write the factors starting from row 4
    factors_range = 'A4:A' + str(len(factors) + 3)
    page_structure_worksheet.update(factors_range, [[factor] for factor in factors])

    # Write the values starting from row 3
    values_range = 'A3:' + gspread.utils.rowcol_to_a1(3, len(values))
    page_structure_worksheet.update(values_range, [values])

    # Get the number of rows and columns in the worksheet
    all_values = page_structure_worksheet.get_all_values()
    num_rows = len(all_values)
    num_cols = len(all_values[0])

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt = CellFormat(wrapStrategy='CLIP')

    # Create a CellFormat object for bold text
    bold_fmt = CellFormat(textFormat=TextFormat(bold=True))

    # Apply the formatting to all cells in the worksheet and bold formatting to row 3
    ranges_fmt = [
        (f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt),
        (f'A3:{gspread.utils.rowcol_to_a1(3, num_cols)}', bold_fmt)
    ]
    format_cell_ranges(page_structure_worksheet, ranges_fmt)

    # Double the width of column A
    set_column_width(page_structure_worksheet, 'A:A', 150)  # 100 is the standard width
