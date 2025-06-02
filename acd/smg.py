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



def add_iplnom_to_smg(
        smg: str,
        ipl1: str,
        ipl2: str = None,
        debug: bool = False,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> int:
    """Add the iplnom to the smg file

    Args:
        smg (str): path + filename of the smg file
        ipl1 (str): path + filename of the ipl1 file
        ipl2 (str, optional): path + filename of the ipl2 file. Defaults to None.
        debug (bool, optional): print debug messages. Defaults to False.

    """
    unarchive_file(smg)

    smg_dir = smg.replace(".smg", "")
    filename = basename(smg).replace(".smg", "")

    try:
        try:
            with open(join(smg_dir, "product.smgXml"), "r", encoding="utf-8") as smg_xxml:
                smg_content = smg_xxml.read()
        except FileNotFoundError:
            try:
                with open(join(smg_dir, filename, "product.smgxml"), "r", encoding="utf-8") as smg_xxml:
                    smg_content = smg_xxml.read()
            except FileNotFoundError:
                if qt_window is not None:
                    progress.emit(100)
                    console.emit(f"Finished with error: {FileNotFoundError}\n{format_exc()}")
                return 1

        ipl_dict = ipl_to_dict(ipl1)
        if ipl2 is not None:
            ipl_dict.update(ipl_to_dict(ipl2))
        new_sgm_content = smg_content
        for part in findall(r'(<Actor.Name Value=")(.*?)(")', smg_content):  # <Actor.Name Value="MS21902J4 |  | 81343 | DMU"/>
            pnr = part[1].split(" ")[0].strip()
            if debug:
                print(f"pnr: {pnr}")
            if pnr in ipl_dict:
                # new_sgm_content = sub(r"", r"", new_sgm_content)
                new_sgm_content = new_sgm_content.replace("".join(part), f'<Actor.Name Value="{pnr + " | " + str(ipl_dict[pnr]["Itemnbr"]) + " " + ipl_dict[pnr]["Nomenclature"]}"')
            elif pnr.replace('-', '') in ipl_dict:
                new_sgm_content = new_sgm_content.replace("".join(part), f'<Actor.Name Value="{pnr + " | " + str(ipl_dict[pnr.replace("-", "")]["Itemnbr"]) + " " + ipl_dict[pnr.replace("-", "")]["Nomenclature"]}"')

        if isfile(join(smg_dir, "product.smgxml")):
            with open(join(smg_dir, "product.smgXml"), "w", encoding="utf-8") as smg_xxml:
                smg_xxml.write(new_sgm_content)
        elif isfile(join(smg_dir, filename, "product.smgXml")):
            with open(join(smg_dir, filename, "product.smgXml"), "w", encoding="utf-8") as smg_xxml:
                smg_xxml.write(new_sgm_content)
        else:
            if qt_window is not None:
                progress.emit(100)
                console.emit("There was an error while trying to find the product.smgXml file.")
            return 1
        remove(smg)
        zip_folder(smg_dir)
        deleted = False
        while not deleted:
            try:
                rmtree(smg_dir)
                deleted = True
            except PermissionError:
                sleep(0.1)
        rename(smg_dir + ".zip", smg)
    except Exception as err:
        if qt_window is not None:
            progress.emit(100)
            console.emit(f"Finished with error: {err}\n{format_exc()}")
        return 1
    if qt_window is not None:
        progress.emit(100)
    return 0

if __name__ == "__main__":
    add_iplnom_to_smg(
        r"D:\Automation\Illu Automation\3D-files\6235A0000-03_rev02.smg",
        r"D:\Automation\Illu Automation\CMM\322131DPLIST_UK.XML",
        r"D:\Automation\Illu Automation\CRM\322131DPLIST_UK.XML"
    )