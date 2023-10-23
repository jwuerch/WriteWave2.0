import gspread
from gspread_formatting import *
def create_worksheet_overview(keyword_sheet):

    # Check if 'Overview' worksheet exists and create if not
    try:
        overview_worksheet = keyword_sheet.worksheet('Overview')  # attempt to get the worksheet
    except gspread.WorksheetNotFound:
        # if worksheet does not exist, create a new one
        overview_worksheet = keyword_sheet.add_worksheet(title='Overview', rows="1000", cols="30")

    titles = ['Factor ID', 'Factor Name', 'Page 1 Average', 'Result 1', 'Result 2', 'Result 3', 'Average', 'Maximum', 'Goal']

    # Update titles in the second row of 'SERP Data'
    cell_list = overview_worksheet.range('A2:I2')
    for j, cell in enumerate(cell_list):
        cell.value = titles[j]
    overview_worksheet.update_cells(cell_list)

    # Format the second row
    fmt = cellFormat(
        backgroundColor=color(0.05, 0.33, 0.58),
        textFormat=textFormat(bold=True, fontSize=13, foregroundColor=color(1, 1, 1)),
        horizontalAlignment='CENTER',
        verticalAlignment='MIDDLE'
    )
    format_cell_range(overview_worksheet, 'A2:I2', fmt)

    # Set column widths with some padding
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": overview_worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": {
                    "pixelSize": len(title) * 9 + 40  # adjust the multiplier and padding as needed
                },
                "fields": "pixelSize"
            }
        } for i, title in enumerate(titles)
    ]

    # Double the row height using Google Sheets API
    requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": overview_worksheet.id,
                "dimension": "ROWS",
                "startIndex": 1,
                "endIndex": 2
            },
            "properties": {
                "pixelSize": 50
            },
            "fields": "pixelSize"
        }
    })

    # Batch update all requests
    keyword_sheet.batch_update({"requests": requests})

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt = cellFormat(
        wrapStrategy='CLIP'
    )

    # Apply the formatting to all cells in the worksheet
    format_cell_range(overview_worksheet, f'A1:{gspread.utils.rowcol_to_a1(overview_worksheet.row_count, overview_worksheet.col_count)}', fmt)

