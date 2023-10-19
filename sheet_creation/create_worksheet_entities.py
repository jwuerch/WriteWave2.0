import gspread
from gspread_formatting import *
def create_worksheet_entities(client, keyword_sheet_url):
    # Open the Google Sheet
    keyword_sheet = client.open_by_url(keyword_sheet_url)

    # Check if 'Entities' worksheet exists and create if not
    try:
        entities_worksheet = keyword_sheet.worksheet('Entities')  # attempt to get the worksheet
    except gspread.WorksheetNotFound:
        # if worksheet does not exist, create a new one
        entities_worksheet = keyword_sheet.add_worksheet(title='Entities', rows="1000", cols="30")

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

    # Prepare the requests for the batch update
    requests = [
        {
            'updateCells': {
                'range': {
                    'sheetId': entities_worksheet.id,  # You need the ID of the sheet here.
                    'startRowIndex': 1,  # Because it's zero-based index, C2 will be row 1
                    'startColumnIndex': 2,  # C is the third column, but zero-based index so it's 2
                },
                'rows': [
                    {
                        'values': [{'userEnteredValue': {'stringValue': '# used'}}]
                    }
                ],
                'fields': 'userEnteredValue',  # We're updating the value only.
            }
        },
        {
            'updateCells': {
                'range': {
                    'sheetId': entities_worksheet.id,  # Again, you need the ID of the sheet here.
                    'startRowIndex': 2,  # This is A3, so it's the third row (zero-based index)
                    'startColumnIndex': 0,  # A is the first column
                },
                'rows': [{'values': [{'userEnteredValue': {'stringValue': cell[0]}}]} for cell in data],
                'fields': 'userEnteredValue',  # We're updating the value only.
            }
        }
    ]

    # Execute the batch update
    keyword_sheet.batch_update({'requests': requests})

    # Get the total number of rows and columns in the worksheet
    num_rows = entities_worksheet.row_count
    num_cols = entities_worksheet.col_count

    # Create a CellFormat object with bold text and text wrapping set to 'CLIP'
    fmt = cellFormat(
        textFormat=textFormat(bold=True),
        wrapStrategy='CLIP'
    )

    # Apply the formatting to all cells in the worksheet
    format_cell_range(entities_worksheet, f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}', fmt)

    # Expand the width of column A
    set_column_width(entities_worksheet, 'A:A', 150)  # 100 is the standard width

