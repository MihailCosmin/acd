"""This module provides functions to add, update or delete the version of a filename.
"""
from os import rename
from os import remove
from os.path import basename

from shutil import copy

from re import sub
from re import search

from tqdm import tqdm

from .filelist import list_files

VER_REGEX_1 = r"_v\d\.\d\.\d"

def add_filename_version(directory: str, version: str, extension: str) -> None:
    """This function adds a version to the filename of all files in a directory.
    It adds it before the given extension.

    Args:
        directory (str): The directory to search in.
        version (str): The version to add.
        extension (str): The extension to add the version before.
    """
    for file_path in tqdm(list_files(directory, True)):
        if file_path.endswith(extension):
            new_file_path = file_path.replace(extension, version + extension)
            rename(file_path, new_file_path)

def update_filename_version(directory: str, version: str, extension: str, reg_ex: str) -> None:
    """This function updates a version in the filename of all files in a directory.

    Args:
        directory (str): The directory to search in.
        version (str): The version to update.
        extension (str): The extension to update the version after.
    """
    for file_path in tqdm(list_files(directory, True)):
        if file_path.endswith(extension):
            new_file_path = sub(reg_ex, version, file_path)
            rename(file_path, new_file_path)

def delete_filename_version(directory: str, extension: str, reg_ex: str) -> None:
    """This function deletes a version from the filename of all files in a directory.
    It deletes it after the given extension.

    Args:
        directory (str): The directory to search in.
        extension (str): The extension to delete the version after.
        reg_ex (str): The regular expression to match the version.
    """
    for file_path in tqdm(list_files(directory, True)):
        if file_path.endswith(extension):
            new_file_path = sub(reg_ex, "", file_path)
            rename(file_path, new_file_path)

def increase_filename_version(file: str, increase_unit: int = 1, reg_ex: str = None, back_up: bool = False, debug: bool = False) -> int:
    """This function increases the version of a file.

    Args:
        file (str): The file to increase the version of.
        increase_unit (int, optional): The unit to increase the version by. Defaults to 1.
        reg_ex (str, optional): The regular expression to match the version. Defaults to None.
        back_up (bool, optional): If True, the function will create a back-up of the file. Defaults to False.
        debug (bool, optional): If True, the function will print debug messages. Defaults to False.

    Returns:
        int: 0 if successful, 1 if not
    """
    extension = file.split(".")[-1]

    if debug:
        print(f'Increasing version of "{file}" by {increase_unit}')

    try:
        if not reg_ex:
            reg_ex = r"(\d+)"
            if search(reg_ex, file):
                version = search(reg_ex, file).group(1)
                if back_up:
                    copy(file, file.replace(extension, f"_backup{extension}"))
                new_filename = sub(version, str(int(version) + 1).zfill(len(version)), file)
                rename(file, new_filename)
            elif debug:
                print(f'No version found in "{file}"')
        else:
            print("No case for this yet. Please add it.")
            return 1

    except Exception as err:
        print(f"Error, could not increase version: {err}")
        return 1
    return 0


if __name__ == "__main__":
    pass
    increase_filename_version(r"D:\Liebherr Automation\TV_DAD06.docx")
    # add_filename_version(r"C:\wamp64\www\seat-configurator\resources\backgrounds\bridge", "_v0.0.1", ".png")
    # update_filename_version(r"C:\wamp64\www\seat-configurator\resources\backgrounds\bridge", "_v0.0.1", ".png", VER_REGEX_1)
    # delete_filename_version(r"C:\wamp64\www\seat-configurator\resources\backgrounds\bridge", ".png", VER_REGEX_1)
