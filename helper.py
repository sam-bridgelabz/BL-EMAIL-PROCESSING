from datetime import datetime, timedelta
import time
import re


def get_text_between_tags(text):
    start_index = text.find("<")
    end_index = text.find(">")

    if start_index != -1 and end_index > start_index:
        # Extract text between tags (excluding brackets)
        return text[start_index + 1:end_index]
    else:
        return None


def get_spreadsheet_file_name():
    month = str(datetime.now().month)
    file_name = "Month_"+month+"_Spreadsheet_Data"
    return file_name


def get_main_path():
    return "Email_data@Bdlaz"


def get_current_date():
    # Get current date in the format YYYY/MM/DD
    current_date = datetime.now().strftime('%Y/%m/%d')
    return current_date


def get_next_day_date():
    # Get the current date and add one day using timedelta
    next_day = datetime.now() + timedelta(days=1)
    # Return the next day's date in the format YYYY/MM/DD
    return next_day.strftime('%Y/%m/%d')


def join_non_empty_strings(*args):
    # Filter out empty strings and join remaining with commas
    return ", ".join([re.search(r"<(.*?)>", arg).group(1) for arg in args if re.search(r"<(.*?)>", arg)])


def extract_date_time_components(date_string):
    try:
        # Define the format of the input date string
        date_format = "%a, %d %b %Y %H:%M:%S %z"

        # Parse the date string to a datetime object
        date_obj = datetime.strptime(date_string, date_format)

        # Extract and convert components to strings
        day = str(date_obj.day)
        # month = date_obj.strftime("%B")  # Full month name as a string
        month = str(date_obj.month)
        year = str(date_obj.year)
        hour = str(date_obj.hour)
        minute = str(date_obj.minute)
        second = str(date_obj.second)

        return day, month, year, hour, minute, second

    except ValueError:
        # Return an empty tuple if the date string does not match the format
        return 'D', 'M', 'Y', 'H', 'M', 'S'
