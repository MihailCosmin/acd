from os import walk
from os.path import sep
from os.path import join
from os.path import isdir
from os.path import dirname
from os.path import relpath
from os.path import basename

from shutil import rmtree

from subprocess import call

from traceback import format_exc

from zipfile import ZipFile
from zipfile import ZIP_DEFLATED
from zipfile import BadZipFile

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal


def unarchive_file(
        word_file: str,
        debug: bool = False,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")):
    """Unarchive a given word file

    Args:
        word_file (str): path + filename of the word file
    """
    # filename = word_file.split(sep)[-1]  # this was failing even on windows, the sep was / instead of \
    filename = basename(word_file)  # we should probably always use basename and join

    # print(f"filename: {filename}")
    ext = "." + filename.split(".")[-1]
    # print(f"ext: {ext}")
    # print(f"word archive: {dirname(word_file)}{sep}{filename.replace(ext, '').strip()}")
    unarchived_folder = join(dirname(word_file), filename.replace(ext, '').strip())
    message = ""
    try:
        with ZipFile(word_file, 'r') as zip_ref:
            if isdir(unarchived_folder):
                rmtree(unarchived_folder)
            if debug:
                print(unarchived_folder)
            zip_ref.extractall(unarchived_folder)
            message = f"Unarchived successfully: {filename} with zipfile."
    except BadZipFile:
        if isdir(unarchived_folder):
            rmtree(unarchived_folder)
        try:
            seven_unzip(word_file)
            message = f"BadZipFile, unarchived successfully: {filename} with 7zip."
        except Exception:
            message = f'Bad Word File: {filename.strip()}. Please check it.\n'
            console.emit(message)
            if debug:
                print(message)
            if isdir(unarchived_folder):
                rmtree(unarchived_folder)
    except PermissionError:
        message = f'Permission Error: {filename.strip()}\n'
        if debug:
            print(message)
        console.emit(message)
        rmtree(unarchived_folder)
    except Exception as err:
        try:
            seven_unzip(word_file)
            message = f"Unarchived successfully: {filename} with 7zip."
        except Exception:
            if debug:
                print(message)
            if isdir(unarchived_folder):
                rmtree(unarchived_folder)
            message = f"Other error: {err}\n{format_exc()}\n"
            if debug:
                print(message)
            console.emit(message)

    if not isdir(unarchived_folder):
        message = f"File not unarchived correctly {filename}.\n{format_exc()}\n"
        if debug:
            print(message)
        if qt_window is not None:
            console.emit(message)
    return message


def zip_word_folder(
        folder: str,
        strict_timestamps: bool = False):
    """Adds the sub-folders of an word folder to a zip file

    Args:
        folder (str): The Word Folder
    """
    with ZipFile(f"{folder}.zip", 'w', ZIP_DEFLATED, strict_timestamps=strict_timestamps) as zipf:
        zipdir(f"{folder}{sep}_rels", zipf)
        zipdir(f"{folder}{sep}docProps", zipf)
        zipdir(f"{folder}{sep}word", zipf)
        if isdir(f"{folder}{sep}customXml"):
            zipdir(f"{folder}{sep}customXml", zipf)
        zipf.write(f"{folder}{sep}[Content_Types].xml", "[Content_Types].xml")

def zip_excel_folder(
        folder: str,
        strict_timestamps: bool = False):
    """Adds the sub-folders of an word folder to a zip file

    Args:
        folder (str): The Word Folder
    """
    with ZipFile(f"{folder}.zip", 'w', ZIP_DEFLATED, strict_timestamps=strict_timestamps) as zipf:
        zipdir(f"{folder}{sep}_rels", zipf)
        zipdir(f"{folder}{sep}docProps", zipf)
        zipdir(f"{folder}{sep}xl", zipf)
        if isdir(f"{folder}{sep}customXml"):
            zipdir(f"{folder}{sep}customXml", zipf)
        zipf.write(f"{folder}{sep}[Content_Types].xml", "[Content_Types].xml")

def zip_folder(folder: str, strict_timestamps: bool = False):
    """Zip a folder

    Args:
        folder (str): _description_
        strict_timestamps (bool, optional): _description_. Defaults to False.
    """
    with ZipFile(f"{folder}.zip", 'w', ZIP_DEFLATED, strict_timestamps=strict_timestamps) as zipf:
        for root, _, files in walk(folder):
            for file in files:
                if file != f"{folder}.zip":
                    zipf.write(join(root, file),
                               relpath(join(root, file),
                                       join(folder, '..')))
        zipf.close()


def zipdir(path: str, ziph: ZipFile, debug: bool = False):
    """Adds all content of a directory to a zip file

    Args:
        path (str): path of the folder to be added to zip file
        ziph (ZipFile): ZipFile where to add the directory
    """
    # ziph is zipfile handle
    for root, _, files in walk(path):
        for file in files:
            if debug:
                print(f"Adding {file} to zip file")
            ziph.write(join(root, file),
                       relpath(join(root, file),
                               join(path, '..')))


def seven_unzip(word: str):
    """Use 7zip to unzip

    Args:
        word (str): word file to unzip
    """
    try:
        command = [
            r'C:\Program Files\7-Zip\7z.exe',
            'x',
            f'{word}',
            f'-o{word.replace(".docx", "")}',
            '-aoa']
        call(command)
    except Exception:
        try:
            command = [
                r'C:\Program Files (x86)\7-Zip\7z.exe',
                'x',
                f'{word}',
                f'-o{word.replace(".docx", "")}',
                '-aoa']
            call(command)
        except Exception:
            try:
                command = [
                    '7z',
                    'x',
                    f'{word}',
                    f'-o{word.replace(".docx", "")}',
                    '-aoa']
                call(command)
            except Exception:
                print(f"Could not unzip {word} with 7zip.")

if __name__ == "__main__":
    file = r"C:\Users\munteanu\Downloads\21-26-01-021\A320A321_IPC_GM_212601_021_0_R00_AUG_01_19.docx"
    filename = basename(file)
    ext = "." + filename.split(".")[-1]
    unarchived_folder = join(dirname(file), filename.replace(ext, '').strip())
    print(unarchived_folder)
