"""This module contains function to perform operations involving dates and time."""
from datetime import datetime

def pdf_date_to_format(pdf_date: str, format_: str) -> str:
    """
    Converts a PDF date to a given format

    Args:
        pdf_date (str): The PDF date.
        format (str): The format to convert the PDF date to.
            Ex: "%d.&m.%Y", "%Y-%m-%d %H:%M:%S"

    Returns:
        str: The PDF date in the given format.
    """
    try:
        return datetime.strptime(pdf_date[2:16], "%Y%m%d%H%M%S").strftime(format_)
    except Exception as err:
        # See which exceptions can occur then add them here.
        print(err)
        return None
    return None
