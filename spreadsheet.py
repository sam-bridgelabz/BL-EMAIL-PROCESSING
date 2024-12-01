from googleapiclient.errors import HttpError


def create_spreadsheet_in_folder(drive_service, folder_id_dict, content_dict, spreadsheet_title="Spreadsheet Data"):
    """
    Creates a new Google Spreadsheet inside a specified folder.
    If a spreadsheet with the same title already exists, it returns the existing spreadsheet's ID.

    Parameters:
    - folder_id: The ID of the folder where the spreadsheet will be created.
    - spreadsheet_title: The title of the spreadsheet to be created (default is "New Spreadsheet").

    Returns:
    - The created or existing spreadsheet's ID.
    """
    try:
        sheet_folder_id_key = content_dict['sheet_folder_id_key']
        folder_id = folder_id_dict.get(sheet_folder_id_key)
        if folder_id is not None:
            # Search for an existing spreadsheet with the given title in the specified folder
            query = f"mimeType='application/vnd.google-apps.spreadsheet' and name='{
                spreadsheet_title}' and '{folder_id}' in parents"
            results = drive_service.files().list(q=query, fields="files(id, name)").execute()
            existing_files = results.get('files', [])

            if existing_files:
                # If an existing file is found, return its ID
                spreadsheet_id = existing_files[0]['id']
                print(f"Spreadsheet with title '{
                      spreadsheet_title}' already exists. Using existing ID: {spreadsheet_id}")
            else:
                # If no existing file is found, create a new one
                file_metadata = {
                    'name': spreadsheet_title,
                    'mimeType': 'application/vnd.google-apps.spreadsheet',
                    # Place the file in the specified folder
                    'parents': [folder_id]
                }
                # Create the spreadsheet in Google Drive
                file = drive_service.files().create(body=file_metadata, fields='id').execute()
                spreadsheet_id = file.get('id')
                print(f"Spreadsheet created with ID: {
                      spreadsheet_id} inside folder ID: {folder_id}")

            return spreadsheet_id
        else:
            return {'info': 'No sheet data available for this month'}

    except HttpError as e:
        print(f"Error creating or fetching spreadsheet: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None


def write_data_to_spreadsheet(sheet_service, spreadsheet_id, data, link, range_name="Sheet1!A1"):
    """
    Appends specific fields from a single dictionary to a Google Spreadsheet, adding headers only if the sheet is empty.

    Parameters:
    - spreadsheet_id: The ID of the spreadsheet where the data will be written.
    - data: A dictionary to write to the spreadsheet.
    - link: An external link to add alongside the data.
    - range_name: The range in the spreadsheet to start writing (default is "Sheet1!A1").

    Returns:
    - The updated range where the data was written.
    """
    try:
        # Specify the keys to include from the dictionary and set up headers
        keys_to_include = ['from', 'to', 'date']
        headers = keys_to_include + ['link']

        # Check if the spreadsheet already has data
        existing_data = sheet_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()

        # Prepare the values to append
        row_values = [[data.get(key, '')
                       for key in keys_to_include] + [link or '']]

        # If the sheet is empty, prepend headers
        if 'values' not in existing_data or not existing_data['values']:
            values = [headers] + row_values
        else:
            values = row_values  # No headers if data already exists

        # Specify the request body with the values to append
        body = {
            'values': values
        }

        # Use the Sheets API to append the data to the spreadsheet
        request = sheet_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body
        )
        response = request.execute()

        # Print the updated range if available
        updated_range = response.get('updates', {}).get('updatedRange', None)
        if updated_range:
            print(f"Data appended to range: {updated_range}")
        else:
            print("No updatedRange returned in the response.")
        return updated_range

    except HttpError as err:
        print(f"An error occurred: {err}")
        return None
