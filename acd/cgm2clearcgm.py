from os.path import sep
from os.path import join
from os.path import isdir
from os.path import isfile
from os.path import dirname
from os.path import abspath
from os.path import basename

from subprocess import check_output

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal


from .filelist import list_files

CGM2CLEARCGM = join(dirname(abspath(__file__)), "cgm2cleartxt_bin", "cgm2cleartext.exe")

def cgm2svgclear(
        item: str,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> int:
    """Converts a cgm file to svg using cgm2svg.exe

    Args:
        item (str): path to a cgm file or a directory containing cgm files
    """
    try:
        if isfile(item):
            item = item.replace("/", sep).replace("\\", sep)
            res = check_output(f'"{CGM2CLEARCGM}" "{item}"', shell=True)
            if qt_window:
                console.emit(res)
        elif isdir(item):
            scope = list_files(item, True, ["cgm", "CGM"])
            max_progress = len(scope)
            for ind, cgm in enumerate(scope):
                cgm = cgm.replace("/", sep).replace("\\", sep)
                if qt_window is None:
                    print(f"Converting {cgm}")
                if "cleartext" in cgm.lower():
                    continue
                res = check_output(f'"{CGM2CLEARCGM}" "{cgm}"', shell=True)
                if qt_window is not None:
                    progress.emit(int(ind / max_progress * 100))
                    console.emit(f"Converted {cgm}")
        else:
            print("Please enter a valid file or directory path")
    except Exception as err:
        if qt_window is not None:
            progress.emit(100)
            console.emit(f"Finished with error: {err}")
        else:
            print(f"Finished with error: {err}")
        return 1
    if qt_window is not None:
        progress.emit(100)
        console.emit("Finished!")
    return 0

if __name__ == "__main__":
    cgm2svg(r"D:\IT\althomcodebase\althomcodebase\automation\convertors\cgm2svg2")