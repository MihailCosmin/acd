from os import utime
from os.path import sep
from os.path import join
from os.path import getmtime

from datetime import datetime

from time import mktime

from PIL import Image
Image.MAX_IMAGE_PIXELS = None
# https://stackoverflow.com/questions/51152059/pillow-in-python-wont-let-me-open-image-exceeds-limit

from traceback import format_exc

from re import search
from json import dump

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal

from .filelist import list_files
from .txt import get_textfile_content

def illu_date_check(
        folder: str,
        export_json: bool = False,
        fix_dates: bool = False,
        latest_date: bool = True,
        specific_date: any = int,
        debug: bool = False,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> dict:
    """Checks if the date of the illustration is correct.

    Args:
        folder (str): The folder to check.

    Returns:
        dict: The result of the check.
    """
    specific_date = None if specific_date == 946681200 else specific_date

    result = {}
    folder_files = {}
    for cgm in list_files(folder):
        extension = cgm.split(sep)[-1].split(".")[-1]  # Cosmin: in case the filename has multiple dots
        filename_without_extension = cgm.split(sep)[-1].split(".")[0]
        if filename_without_extension not in folder_files:
            folder_files[filename_without_extension] = []
        if extension not in folder_files[filename_without_extension]:
            folder_files[filename_without_extension].append(extension)
    for cgm in list_files(folder, True, ["cgm"]):
        check = "OK"
        filename_without_extension = cgm.split(sep)[-1].split(".")[0]
        extension = cgm.split(sep)[-1].split(".")[-1]
        cgm_date = str(datetime.fromtimestamp(getmtime(cgm)))[:16]
        result[filename_without_extension] = {extension: cgm_date}
        if len(folder_files[filename_without_extension]) > 1:
            for ext in folder_files[filename_without_extension]:
                if ext != extension and ext not in result[filename_without_extension]:
                    adt_illustration = filename_without_extension + "." + ext
                    adt_illustration = join(folder, adt_illustration)
                    adt_date = str(datetime.fromtimestamp(getmtime(adt_illustration)))[:16]
                    if adt_date != cgm_date:
                        check = "NOT OK"
                    result[filename_without_extension][ext] = adt_date

            result[filename_without_extension]["check"] = check
    if fix_dates:
        try:
            max_progress = len(result)
            for ind, filename in enumerate(result):
                if specific_date is None:
                    newest_date = datetime(1970, 1, 1) if latest_date else datetime(2130, 1, 1)
                    for ext in result[filename]:
                        if ext != "check":
                            adt_date = result[filename][ext]
                            adt_date = datetime.strptime(adt_date, "%Y-%m-%d %H:%M")
                            if latest_date:
                                if adt_date > newest_date:
                                    newest_date = adt_date
                            elif adt_date < newest_date:
                                newest_date = adt_date
                else:
                    newest_date = datetime.fromtimestamp(specific_date)
                for ext in result[filename]:
                    if ext != "check":
                        adt_illustration = filename + "." + ext
                        adt_illustration = join(folder, adt_illustration)
                        utime(adt_illustration, (mktime(newest_date.timetuple()), mktime(newest_date.timetuple())))
                if qt_window:
                    progress.emit(int(ind / max_progress * 100))
                    console.emit(f"Set the date for {filename} to {newest_date}")
        except Exception as err:
            if qt_window:
                progress.emit(100)
                console.emit(f"Finished with error: {err}\n{format_exc()}")
            return 1
        if qt_window:
            progress.emit(100)
            console.emit(f"Finished setting the dates!")
        return 0

    if export_json:
        with open(join(folder, "illu_date_check.json"), "w", encoding="utf-8") as out:
            dump(result, out, indent=4)
    return result

def check_cgm_details(folder: str, check_strings: list, export_json: bool = False) -> dict:
    result = {}
    for cgm in list_files(folder, True, ["cgm"]):
        # with open(cgm, "r", encoding="utf-8") as cgm_in:
        #     cgm_data = cgm_in.read()
        cgm_data = get_textfile_content(cgm)

        if any(string not in cgm_data for string in check_strings):
            result[cgm] = "NOT OK"
        else:
            result[cgm] = "OK"

    if export_json:
        with open(join(folder, "cgm_details_check.json"), "w", encoding="utf-8") as out:
            dump(result, out, indent=4)
    return result

def check_tif_details(folder: str, export_json: bool = False) -> dict:
    result = {}
    for tif in list_files(folder, True, ["tif", "tiff"]):
        image = Image.open(tif)
        width, height = image.size
        dpi_x, _ = image.info.get('dpi', (None, None))

        result[tif] = {
            "Resolution": str(width) + "x" + str(height),
            "DPI": str(int(dpi_x))
        }

    if export_json:
        with open(join(folder, "tif_details_check.json"), "w", encoding="utf-8") as out:
            dump(result, out, indent=4)
    return result

if __name__ == "__main__":
    illu_date_check(r"C:\Users\munteanu\Downloads\TEST", True)
