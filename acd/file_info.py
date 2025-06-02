"""This module provides functions to retrieve information about a file.
"""

from os.path import getsize

def get_file_size(file: str) -> int:
    """
    This function takes a file and returns the size of the file in MB.
    """
    return getsize(file) / 1024 / 1024