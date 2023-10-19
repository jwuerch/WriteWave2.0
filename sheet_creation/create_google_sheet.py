
def create_google_sheet(client, keyword, data_folder_id, writewave_google_sheet, google_sheet_col_index, i):
    # Create a new Google Sheet with the keyword as the name in the 'Data' folder
    keyword_sheet = client.create(keyword, folder_id=data_folder_id)

    # Get the URL of the new Google Sheet
    keyword_sheet_url = f"https://docs.google.com/spreadsheets/d/{keyword_sheet.id}"

    # Update the 'Datasheet' column in the Google Sheet
    writewave_google_sheet.update_cell(i + 1, google_sheet_col_index + 1, keyword_sheet_url)
    return keyword_sheet_url

