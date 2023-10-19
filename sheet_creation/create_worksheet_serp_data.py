import gspread
from gspread_formatting import *
def create_worksheet_serp_data(client, keyword_sheet_url):
    # Open the Google Sheet and get the 'Keyword Variations' worksheet
    keyword_sheet = client.open_by_url(keyword_sheet_url)

    # Rename Sheet1 to 'SERP Data'
    worksheet = keyword_sheet.get_worksheet(0)
    worksheet.update_title('SERP Data')

    titles = ['#', 'Search Result', 'SEO Title', 'Meta Description', 'People Also Ask', 'Answer', 'Video',
              'Featured Snippet', 'Keyword Variations']

    # Update titles in the second row of 'SERP Data'
    cell_list = worksheet.range('A2:I2')
    for j, cell in enumerate(cell_list):
        cell.value = titles[j]
    worksheet.update_cells(cell_list)

    # Format the second row
    fmt = cellFormat(
        backgroundColor=color(0.05, 0.33, 0.58),
        textFormat=textFormat(bold=True, fontSize=13, foregroundColor=color(1, 1, 1)),
        horizontalAlignment='CENTER',
        verticalAlignment='MIDDLE'
    )
    format_cell_range(worksheet, 'A2:I2', fmt)

    # Set column widths with some padding
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
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
                "sheetId": worksheet.id,
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
    format_cell_range(worksheet, f'A1:{gspread.utils.rowcol_to_a1(worksheet.row_count, worksheet.col_count)}', fmt)

