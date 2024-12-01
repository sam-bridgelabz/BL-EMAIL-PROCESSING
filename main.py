from auth import get_authenticated_services
from emails import get_emails_in_date_range
from datetime import datetime
from drive import create_folder_in_drive, create_and_write_doc_in_folder
from helper import get_main_path
from logger import logger

# def get_date_string():
#     today = datetime.date.today()
#     midnight = datetime.datetime.combine(today, datetime.time(0, 0, 0))

#     # Format the date and time as a string with seconds
#     date_string_with_seconds = midnight.strftime("%Y-%m-%d %H:%M:%S")

#     return date_string_with_seconds  # Output: 2023-11-08 00:00:00


if __name__ == "__main__":
    gmail_service, sheet_service, drive_service, doc_service = get_authenticated_services()
    data = get_emails_in_date_range(gmail_service, "2024/11/14", "2024/11/22")
    logger.info(f"Finished processing fetched email data")
    if isinstance(data, list):
        spreadsheet_folder_id_dict = create_folder_in_drive(
            drive_service, get_main_path(), "spreadsheet")
        folder_dict = create_folder_in_drive(
            drive_service, get_main_path(), "docs")
        logger.info(f"Started writing data to docs and spreadsheet")
        for msg in data:
            date = msg['date']
            subject = msg['subject']
            msg_text = msg['message']
            title = "Email_"+date+"_full_message"
            create_and_write_doc_in_folder(
                sheet_service, drive_service, doc_service, folder_dict, spreadsheet_folder_id_dict, title, msg)
        logger.info(f"Finished writing data to docs and spreadsheet")
    else:
        logger.info(f"Data already present")
        raise Exception(data)
