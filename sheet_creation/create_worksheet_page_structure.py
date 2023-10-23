import gspread
from gspread_formatting import *

def create_worksheet_page_structure(keyword_sheet):
    # Check if 'Page Structure' worksheet exists and create if not
    try:
        page_structure_worksheet = keyword_sheet.worksheet('Page Structure')  # attempt to get the worksheet
    except gspread.WorksheetNotFound:
        # if worksheet does not exist, create a new one
        page_structure_worksheet = keyword_sheet.add_worksheet(title='Page Structure', rows="1000", cols="30")

    # Define the values to be written
    values = [
        'Factor', 'Result 1', 'Result 2', 'Result 3', 'Result 4', 'Result 5',
        'Result 6', 'Result 7', 'Result 8', 'Result 9', 'Result 10',
        'Page 1 Average', 'Page 1 Maximum'
    ]

    factors = [
        'Word Count', 'H1 Tag Count', 'H2 Tag Count', 'H3 Tag Count', 'H4 Tag Count',
        'H5 Tag Count', 'H6 Tag Count', 'p Tag Count', 'a Tag Count', 'a Internal Tag Count',
        'a External Tag Count', 'img Tag Count', 'alt Tag Count', 'b Tag Count', 'strong Tag Count',
        'i Tag Count', 'em Tag Count', 'u Tag Count', 'ol Tag Count', 'ol Item Count',
        'ul Tag Count', 'ul Item Count', 'table Tag Count', 'form Tag Count', 'iframe Tag Count',
        'video Tag Count', 'FAQs Count', 'TOC Count', 'SEO Title', 'Meta Description'
    ]

    # Write the values starting from row 3
    values_range = 'A3:' + gspread.utils.rowcol_to_a1(3, len(values))
    page_structure_worksheet.update(values_range, [values])

    # Write the factors starting from row 4
    factors_range = 'A4:A' + str(len(factors) + 3)
    page_structure_worksheet.update(factors_range, [[factor] for factor in factors])

    # Refresh the worksheet data after writing
    all_values = page_structure_worksheet.get_all_values()
    num_rows = len(all_values)
    num_cols = len(all_values[0])

    # Create a CellFormat object with text wrapping set to 'CLIP'
    fmt_clip = CellFormat(wrapStrategy='CLIP')

    # Apply 'CLIP' text wrapping to all cells in the worksheet
    all_cells_range = f'A1:{gspread.utils.rowcol_to_a1(num_rows, num_cols)}'
    format_cell_range(page_structure_worksheet, all_cells_range, fmt_clip)

    # Create a CellFormat object for bold text
    bold_fmt = CellFormat(textFormat=TextFormat(bold=True))

    # Apply bold formatting to row 3
    bold_range = f'A3:{gspread.utils.rowcol_to_a1(3, num_cols)}'
    format_cell_range(page_structure_worksheet, bold_range, bold_fmt)

    # Double the width of column A
    set_column_width(page_structure_worksheet, 'A', 150)  # 100 is the standard width
