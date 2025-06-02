import sys

from os import remove
from os import rename
from os.path import join
from os.path import isfile
from os.path import dirname
from os.path import basename

from shutil import rmtree

from re import findall
from re import sub

from traceback import format_exc

from time import sleep

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal


from .archive import unarchive_file
from .archive import zip_folder
from .ataispec2200 import ipl_to_dict


def add_iplnom_to_stp(
        stp: str,
        ipl1: str,
        ipl2: str = None,
        debug: bool = False,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> int:
    """Add the iplnom to the smg file

    Args:
        stp (str): path + filename of the stp file
        ipl1 (str): path + filename of the ipl1 file
        ipl2 (str, optional): path + filename of the ipl2 file. Defaults to None.
        debug (bool, optional): print debug messages. Defaults to False.

    """

    try:
        with open(stp, "r", encoding="utf-8") as stp_file:
            stp_content = stp_file.read()

        ipl_dict = ipl_to_dict(ipl1)
        if ipl2 is not None:
            ipl_dict.update(ipl_to_dict(ipl2))
        for part in findall(r"(PRODUCT\(')(.*?)(')", stp_content):  # <Actor.Name Value="MS21902J4 |  | 81343 | DMU"/>
            pnr = part[1].split(" ")[0].strip()
            if debug:
                print(f"pnr: {pnr}")
            if pnr in ipl_dict:
                stp_content = stp_content.replace("".join(part), f"PRODUCT('{pnr + ' | ' + str(ipl_dict[pnr]['Itemnbr']) + ' ' + ipl_dict[pnr]['Nomenclature']}'")
            elif pnr.replace("-", "") in ipl_dict:
                stp_content = stp_content.replace("".join(part), f"PRODUCT('{pnr + ' | ' + str(ipl_dict[pnr.replace('-', '')]['Itemnbr']) + ' ' + ipl_dict[pnr.replace('-', '')]['Nomenclature']}'")

        with open(stp, "w", encoding="utf-8") as stp_file:
            stp_file.write(stp_content)
    except Exception as err:
        if debug:
            print(f"Error: {err}\n{format_exc()}")
        if qt_window is not None:
            progress.emit(100)
            console.emit(f"Finished with error: {err}\n{format_exc()}")
        return 1

    if qt_window is not None:
        progress.emit(100)
    return 0

if __name__ == "__main__":
    add_iplnom_to_stp(
        r"D:\CMM Automation\REWORK\ILLU\3D file and list for EFW\D252R1258-004-00_A-APPROVED.stp",
        r"D:\CMM Automation\REWORK\ILLU\3D file and list for EFW\A320A321_IPC_GM_253101_098X_R12_AUG_15_23.xlsx",
        debug=True
    )
