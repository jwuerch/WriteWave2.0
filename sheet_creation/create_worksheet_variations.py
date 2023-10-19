import gspread
from gspread_formatting import *
def create_worksheet_variations(client, keyword_sheet_url):
    # Open the Google Sheet and get the 'Keyword Variations' worksheet
    keyword_sheet = client.open_by_url(keyword_sheet_url)

    # Create a new sheet titled 'Page Structure'
    variations_worksheet = keyword_sheet.add_worksheet(title='Variations', rows="1000", cols="30")

    # Define the data to be pasted
    data = [
        ['Result 1'],
        ['Result 2'],
        ['Result 3'],
        ['Result 4'],
        ['Result 5'],
        ['Result 6'],
        ['Result 7'],
        ['Result 8'],
        ['Result 9'],
        ['Result 10'],
        ['Page 1 Average'],
        ['Page 1 Maximum']
    ]

    # Create a list of updates
    updates = [
        {
            'range': 'C2',
            'values': [['# used']]
        },
        {
            'range': f'A3:A{len(data) + 3}',
            'values': data
        }
    ]

    # Apply all updates in a single batch
    variations_worksheet.batch_update(updates)

    # Get the total number of rows and columns in the worksheet
    num_rows = variations_worksheet.row_count
    num_cols = variations_worksheet.col_count

    # Create a CellFormat object with bold text and text wrapping set to 'CLIP'
    fmt = cellFormat(
        textFormat=textFormat(bold=True),
        wrapStrategy='CLIP'
    )

    # Apply the formatting to all cells in the worksheet
    format_cell_range(variations_worksheet, f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt)

    # Expand the width of column A
    set_column_width(variations_worksheet, 'A:A', 150)  # 100 is the standard width

