from os import mkdir
from os.path import join
from os.path import isdir
from os.path import isfile
from shutil import copyfile

from .filelist import list_files

def copy_files(directory: str, destination: str, ext_list: list = None, reg_ex: str = None, overwrite: bool = False) -> None:
    """This function copies all files in a directory.

    Args:
        directory (str): The directory to search in.
        destination (str): The directory to copy the files to.
        ext_list (list): The list of extensions to search for.
        reg_ex (str): The regular expression to match the filename.
        overwrite (bool): Whether to overwrite files in the destination directory.

    Returns:
        None
    """
    if not isdir(destination):
        mkdir(destination)
    for file_ in list_files(directory, True, ext_list, reg_ex, include_tqdm=True):
        if isfile(file_):
            if not isfile(join(destination, file_.split("\\")[-1])):
                copyfile(file_, join(destination, file_.split("\\")[-1]))
            elif overwrite:
                copyfile(file_, join(destination, file_.split("\\")[-1]))
        else:
            print(f"Could not copy: {file_}")

if __name__ == '__main__':
    copy_files(r"C:\Users\munteanu\Desktop\SRM R00_R01 FILES", r"C:\Users\munteanu\Desktop\copied2", ['.tiff', '.tif'])
