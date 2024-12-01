from helper import get_text_between_tags
import base64
import json
from datetime import datetime
from helper import get_next_day_date, get_current_date, join_non_empty_strings, extract_date_time_components
from logger import logger

#  proper functions to get and retreive details from email


def get_emails_in_date_range(service, start_date=get_current_date(), end_date=get_next_day_date()):
    """
    Retrieves emails within a specified date range and extracts their details.

    Parameters:
    - service: The Gmail API service instance.
    - start_date (str): Start date in 'YYYY/MM/DD' format.
    - end_date (str): End date in 'YYYY/MM/DD' format.

    Returns:
    - List of dictionaries, each containing details of an email.
    """
    # Convert dates to Gmail's query format (epoch timestamps in seconds)
    start_epoch = int(datetime.strptime(start_date, "%Y/%m/%d").timestamp())
    end_epoch = int(datetime.strptime(end_date, "%Y/%m/%d").timestamp())
    with open('last_created_data.json') as file:
        json_data = json.load(file)

    if not json_data or json_data.get('last_epoch', float('-inf')) < end_epoch:
        # Construct the query string for the date range
        query = f"after:{start_epoch} before:{end_epoch}"

        logger.info(f"Started fetching email data for {
                    start_date} to {end_date} ")
        # Get list of emails in the specified date range
        results = service.users().messages().list(userId="me", q=query).execute()
        logger.info(f"Finished fetching email data for {
                    start_date} to {end_date} ")
        email_details_list = []
        logger.info(f"Starting to process fetched email data")
        # If messages are found
        if 'messages' in results:
            message_ids = results['messages']

            # Retrieve details for each email by calling extract_email_details
            for msg in message_ids:
                email_data = extract_email_details(service, msg['id'])
                email_details_list.append(email_data)

        json_data['last_epoch'] = end_epoch
        # Write the updated data back to the JSON file
        with open('last_created_data.json', 'w') as file:
            json.dump(json_data, file, indent=4)
        return email_details_list
    else:
        return {'info': f'Data till {end_date} already exists'}


def extract_email_details(service, message_id):
    """
    Extracts specific details (from, to, date, subject, message) from an email.

    Parameters:
    - service: The Gmail API service instance.
    - message_id (str): The ID of the email message to retrieve.

    Returns:
    - A dictionary containing the email's from, to, date, subject, and message body.
    """
    email_data = service.users().messages().get(
        userId="me", id=message_id, format='full').execute()
    headers = email_data['payload']['headers']

    # Extract the required headers
    email_info = {
        'from': join_non_empty_strings(next((header['value'] for header in headers if header['name'] == 'From')), ""),
        'date': next((header['value'] for header in headers if header['name'] == 'Date'), ""),
        'subject': next((header['value'] for header in headers if header['name'] == 'Subject'), "")
    }
    to = next((header['value']
              for header in headers if header['name'] == 'To'), "")
    cc = next((header['value']
              for header in headers if header['name'] == 'Cc'), "")
    bcc = next((header['value']
               for header in headers if header['name'] == 'Bcc'), "")
    email_info['to'] = join_non_empty_strings(to, cc, bcc)
    email_info['email_link'] = f"https://mail.google.com/mail/u/0/#inbox/{
        message_id}"

    date_string = next(
        (header['value'] for header in headers if header['name'] == 'Date'), "")
    (day, month, year, hour, minute, second) = extract_date_time_components(date_string)
    email_info['date'] = day+"/"+month+"/"+year+"; "+hour+":"+minute+":"+second
    email_info['docs_folder_id_key'] = "docs_day_"+day
    email_info['sheet_folder_id_key'] = "spreadsheet_"+month+"_"+year

    # Get the full email message (plain text or HTML)
    parts = email_data['payload'].get('parts', [])
    message_body = ""

    # Check for plain text and HTML parts
    for part in parts:
        if part['mimeType'] == 'text/plain':
            # Decode and extract plain text message
            message_body = base64.urlsafe_b64decode(
                part['body'].get('data', '')).decode('utf-8')
            break  # Stop after the first plain text part is found

        elif part['mimeType'] == 'text/html':
            # Decode and extract HTML message
            message_body = base64.urlsafe_b64decode(
                part['body'].get('data', '')).decode('utf-8')
            break  # Stop after the first HTML part is found

        # Check for nested parts if there are any
        elif 'parts' in part:
            nested_parts = part['parts']
            for nested_part in nested_parts:
                if nested_part['mimeType'] == 'text/plain':
                    message_body = base64.urlsafe_b64decode(
                        nested_part['body'].get('data', '')).decode('utf-8')
                    break
                elif nested_part['mimeType'] == 'text/html':
                    message_body = base64.urlsafe_b64decode(
                        nested_part['body'].get('data', '')).decode('utf-8')
                    break

    # If no plain text or HTML was found, you may want to check for other mime types
    if not message_body:
        return None
    email_info['message'] = message_body
    return email_info
