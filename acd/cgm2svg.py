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

CGM2SVG = join(dirname(abspath(__file__)), "cgm2svg.exe")

def cgm2svg(
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
            out_fname = basename(item).split(".")[0] + ".svg"
            out_fname = join(dirname(item), out_fname)
            out_fname = out_fname.replace("/", sep).replace("\\", sep)
            item = item.replace("/", sep).replace("\\", sep)
            res = check_output(f'"{CGM2SVG}" -i "{item}" -o "{out_fname}"', shell=True)
            if qt_window:
                console.emit(res)
        elif isdir(item):
            scope = list_files(item, True, ["cgm", "CGM"])
            max_progress = len(scope)
            for ind, cgm in enumerate(scope):
                out_fname = basename(cgm).split(".")[0] + ".svg"
                out_fname = join(item, out_fname)
                out_fname = out_fname.replace("/", sep).replace("\\", sep)
                cgm = cgm.replace("/", sep).replace("\\", sep)
                if qt_window is None:
                    print(f"Converting {cgm} to {out_fname}")
                res = check_output(f'"{CGM2SVG}" -i "{cgm}" -o "{out_fname}"', shell=True)
                if qt_window is not None:
                    progress.emit(int(ind / max_progress * 100))
                    console.emit(f"Converted {cgm} to {out_fname}")
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
    cgm2svg(r"C:\Users\munteanu\Desktop\LHT - Wiring")
