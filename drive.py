import datetime
import os
import json
from datetime import datetime, timedelta
from googleapiclient.errors import HttpError
from spreadsheet import create_spreadsheet_in_folder, write_data_to_spreadsheet
from helper import get_spreadsheet_file_name

# # creates folder structure in drive EmailData/year/current_date


def create_and_write_doc_in_folder(sheet_service, drive_service, docs_service, folder_dict, spreadsheet_folder_id_dict, doc_title="New Document", content_dict=None):
    """
    Creates a new Google Docs document inside a specified folder and writes data into it from a dictionary.

    Parameters:
    - creds: google.oauth2.credentials.Credentials object.
    - folder_id: The ID of the folder where the document will be created.
    - doc_title: The title of the document to be created (default is "New Document").
    - content_dict: A dictionary where the keys are sections or headings, and the values are the content for those sections.

    Returns:
    - The ID of the created document.
    """
    # gmail_service, service, drive_service, docs_service = get_authenticated_services()
    if content_dict is None:
        content_dict = {}  # Default to an empty dictionary if none provided

    try:
        doc_folder_id_key = content_dict['docs_folder_id_key']
        folder_id = folder_dict[doc_folder_id_key]

        # Create the Google Docs document in the specified folder
        file_metadata = {
            'name': doc_title,
            'mimeType': 'application/vnd.google-apps.document',
            'parents': [folder_id]  # Specify the parent folder ID
        }

        # Create the document using Drive API
        file = drive_service.files().create(body=file_metadata, fields='id').execute()
        document_id = file.get('id')

        # Prepare requests to insert content into the document
        requests = []
        content = f"""
        From: {content_dict['from']}
        To: {content_dict['to']}
        Date: {content_dict['date']}
        Subject: {content_dict['subject']}

        Message:
        {content_dict['message']}
        """

        requests = [
            {'insertText': {'location': {'index': 1}, 'text': content}}]

        # to add spreadsheet from here
        spreadsheet_id = create_spreadsheet_in_folder(
            drive_service, spreadsheet_folder_id_dict, content_dict, get_spreadsheet_file_name())

        if isinstance(spreadsheet_id, dict):
            return spreadsheet_id
        else:
            link = f'https://docs.google.com/document/d/{document_id}/edit'
            write_data_to_spreadsheet(
                sheet_service, spreadsheet_id, content_dict, link, "Sheet1!A1")

            # Execute the batch update to insert all the content
            docs_service.documents().batchUpdate(documentId=document_id,
                                                 body={'requests': requests}).execute()
            return document_id

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def create_folder_in_drive(service, path, type):
    folder_id = None
    current_year = str(datetime.now().year)

    # Folder dictionary to store folder IDs based on type and date structure
    folder_dict = {}

    # Step 1: Create or find the 'Email_data@Bdlaz' folder
    query = "mimeType='application/vnd.google-apps.folder' and name='Email_data@Bdlaz'"
    try:
        results = service.files().list(q=query, spaces='drive',
                                       fields="files(id, name)").execute()
        folders = results.get('files', [])

        if folders:
            folder_id = folders[0]['id']
        else:
            file_metadata = {
                'name': 'Email_data@Bdlaz',
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = service.files().create(body=file_metadata, fields='id').execute()
            folder_id = folder.get('id')

    except Exception as e:
        raise e

    # Step 2: Create Year Folder (e.g., 2024)
    year_folder_name = current_year
    year_folder_id = create_subfolder(service, year_folder_name, folder_id)

    # Step 3: Loop through each month to create folders for the specified type
    for month in range(1, 13):
        month_str = str(month).zfill(2)  # Format month as a 2-digit string
        month_folder_name = f"{month_str}_{current_year}"

        # Create the month folder inside the year folder
        month_folder_id = create_subfolder(
            service, month_folder_name, year_folder_id)

        # Only create folders for the specified `type` (either "docs" or "spreadsheet")
        if type == "docs":
            docs_folder_id = create_subfolder(service, "docs", month_folder_id)
            docs_day_folders = {}
            start_date = datetime.now().replace(month=month, day=1)
            end_date = (start_date.replace(month=start_date.month %
                        # Last day of the month
                                           12 + 1, day=1) - timedelta(days=1)).day
            for day in range(1, end_date + 1):
                day_folder_name = f"{str(day).zfill(2)}_{
                    month_str}_{current_year}"
                day_folder_id = create_subfolder(
                    service, day_folder_name, docs_folder_id)
                docs_day_folders[f"docs_day_{day}"] = day_folder_id

            # Update the folder_dict with day-wise folder IDs for docs
            folder_dict.update(docs_day_folders)

        elif type == "spreadsheet":
            # Only add the month folder ID for spreadsheets without day-wise subfolders
            spreadsheet_folder_id = create_subfolder(
                service, "spreadsheet", month_folder_id)
            folder_dict[f"spreadsheet_{month_str}_{
                current_year}"] = spreadsheet_folder_id

    return folder_dict


# Helper function to create a subfolder
def create_subfolder(service, folder_name, parent_folder_id):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{
        folder_name}'"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"

    try:
        results = service.files().list(q=query, spaces='drive',
                                       fields="files(id, name)").execute()
    except Exception as e:
        print(f"Error executing query: {query}")
        raise e

    existing_folders = results.get('files', [])

    if existing_folders:
        return existing_folders[0]['id']  # Return existing folder ID

    # Create the folder if it doesn't exist
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id] if parent_folder_id else []
    }
    folder = service.files().create(body=file_metadata, fields='id').execute()
    return folder.get('id')

# maybe needed afterwards
# def create_and_move_doc(drive_service, doc_service, folder_id, title, msg_text):
#     document = {
#     'title': title
#     }
#     response = doc_service.documents().create(body=document).execute()
#     document_id = response.get('documentId')
#     print("document_id at create_and_move_doc",document_id)
#     drive_service.files().update(
#         fileId=document_id,
#         addParents=folder_id).execute()
