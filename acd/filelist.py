"""This module provides functions to list all/specific files in a directory.
"""
from os import walk
from os.path import join
from os.path import splitext

from re import search

from zipfile import ZipFile
from zipfile import BadZipFile

from tqdm import tqdm

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal

def list_files3(
        directory: str,
        path: bool = True,
        ext_list: list = None,
        reg_ex: str = None,
        include_tqdm: bool = False,
        search_archives: bool = False):
    """This function lists all files in a directory and optionally inside archives.

    Args:
        directory (str): The directory to search in.
        path (bool): Whether to return the path or just the filename.
        ext_list (list): The list of extensions to search for.
        reg_ex (str): The regular expression to match the filename.
        include_tqdm (bool): Whether to include a progress bar.
        search_archives (bool): Whether to search inside archives.

    Yields:
        str: A file path in the directory (and optionally inside archives) with complete path.
    """
    def file_matches(file_, ext_list, reg_ex):
        if ext_list is not None and not any(file_.endswith(ext) for ext in ext_list):
            return False
        if reg_ex is not None and not search(reg_ex, file_):
            return False
        return True

    def search_archives_recursive(file_path):
        try:
            with ZipFile(file_path) as archive:
                for inner_file in archive.namelist():
                    if file_matches(inner_file, ext_list, reg_ex):
                        yield inner_file if not path else join(file_path, inner_file)
        except BadZipFile:
            pass

    if include_tqdm:
        file_list = []
        for root, _, files in walk(directory):
            for file_ in files:
                if file_matches(file_, ext_list, reg_ex):
                    file_list.append(join(root, file_) if path else file_)

            if search_archives:
                for file_ in files:
                    if file_.endswith('.zip') or file_.endswith('.rar'):
                        file_list.extend(search_archives_recursive(join(root, file_)))

        for file_path in tqdm(file_list):
            yield file_path
    else:
        for root, _, files in walk(directory):
            for file_ in files:
                if file_matches(file_, ext_list, reg_ex):
                    yield join(root, file_) if path else file_

            if search_archives:
                for file_ in files:
                    if file_.endswith('.zip') or file_.endswith('.rar'):
                        yield from search_archives_recursive(join(root, file_))

def list_files2(
        directory: str,
        path: bool = True,
        ext_list: list = None,
        reg_ex: str = None,
        include_tqdm: bool = False,
        search_archives: bool = False) -> list:
    """This function lists all files in a directory and optionally inside archives.

    Args:
        directory (str): The directory to search in.
        path (bool): Whether to return the path or just the filename.
        ext_list (list): The list of extensions to search for.
        reg_ex (str): The regular expression to match the filename.
        include_tqdm (bool): Whether to include a progress bar.
        search_archives (bool): Whether to search inside archives.

    Returns:
        list: A list of all files in the directory (and optionally inside archives) with complete path.
    """
    file_list = []

    # Function to check if a file matches the extension list or pattern
    def file_matches(file_, ext_list, reg_ex):
        if ext_list is not None and not any(file_.endswith(ext) for ext in ext_list):
            return False
        if reg_ex is not None and not search(reg_ex, file_):
            return False
        return True

    # Function to search inside archives
    def search_archives_recursive(file_path):
        try:
            with ZipFile(file_path) as archive:
                for inner_file in archive.namelist():
                    if file_matches(inner_file, ext_list, reg_ex):
                        file_list.append(inner_file if not path else join(file_path, inner_file))
        except BadZipFile:
            pass  # Skip invalid ZIP files

    for root, _, files in walk(directory):
        for file_ in files:
            if file_matches(file_, ext_list, reg_ex):
                file_list.append(join(root, file_) if path else file_)

        if search_archives:
            for file_ in files:
                if file_.endswith('.zip') or file_.endswith('.rar'):
                    search_archives_recursive(join(root, file_))

    if include_tqdm:
        return tqdm(file_list)
    return file_list

def list_files(
        directory: str,
        path: bool = True,
        ext_list: list = None,
        reg_ex: str = None,
        include_tqdm: bool = False,
        debug: bool = False,
        progress_start: int = 0,
        multiplier: int = 1,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> list:
    """This function lists all files in a directory.

    Args:
        directory (str): The directory to search in.
        path (bool): Whether to return the path or just the filename.
        ext_list (list): The list of extensions to search for.
        reg_ex (str): The regular expression to match the filename.
        include_tqdm (bool): Whether to include a progress bar.
        debug (bool): Whether to print debug messages.
        qt_window (QMainWindow): The main window.
        progress (Signal): The progress signal.
        console (Signal): The console signal.

    Returns:
        list: A list of all files in the directory with complete path.
    """
    file_list = []
    ind = 1
    for root, _, files in walk(directory):
        for file_ in files:
            if reg_ex is not None:
                if ext_list is not None:
                    if path and any(file_.endswith(ext) for ext in ext_list) and search(reg_ex, file_) is not None:
                        file_list += [join(root, file_)]
                        ind += 1
                        if qt_window:
                            progress.emit((ind % 2) * multiplier + progress_start)
                    elif any(file_.endswith(ext) for ext in ext_list) and search(reg_ex, file_) is not None:
                        file_list += [file_]
                        ind += 1
                        if qt_window:
                            progress.emit((ind % 2) * multiplier + progress_start)
                elif search(reg_ex, file_) is not None:
                    file_list += [join(root, file_)] if path else [file_]
                    ind += 1
                    if qt_window:
                        progress.emit((ind % 2) * multiplier + progress_start)
            else:
                if ext_list is not None:
                    if path and any(file_.endswith(ext) for ext in ext_list):
                        file_list += [join(root, file_)]
                        ind += 1
                        if qt_window:
                            progress.emit((ind % 2) * multiplier + progress_start)
                    elif any(file_.endswith(ext) for ext in ext_list):
                        file_list += [file_]
                        ind += 1
                        if qt_window:
                            progress.emit((ind % 2) * multiplier + progress_start)
                else:
                    file_list += [join(root, file_)] if path else [file_]
                    ind += 1
                    if qt_window:
                        progress.emit((ind % 2) * multiplier + progress_start)
    if include_tqdm:
        return tqdm(file_list)
    if qt_window:
        progress.emit(multiplier + progress_start)
    return file_list

def get_extensions(directory: str, include_tqdm: bool = False) -> dict:
    """This function gets all the extensions in a directory.

    Args:
        directory (str): The directory to search in.
        include_tqdm (bool): Whether to include a progress bar.

    Returns:
        dict: A dictionary of all extensions in the directory and their count.
    """
    exts = {}
    for file_ in list_files(directory, include_tqdm=include_tqdm):
        if splitext(file_)[1] not in exts:
            exts[splitext(file_)[1]] = 1
        else:
            exts[splitext(file_)[1]] += 1
    return exts
