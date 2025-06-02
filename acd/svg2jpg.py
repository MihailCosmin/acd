from os.path import isdir
from os.path import isfile

from PIL import Image
Image.MAX_IMAGE_PIXELS = None
# https://stackoverflow.com/questions/51152059/pillow-in-python-wont-let-me-open-image-exceeds-limit

import xml.etree.ElementTree as et

from .filelist import list_files
from .xml_processing import delete_first_line

def svg2jpg(item: str):
    """Converts a svg file to jpg using cairosvg

    Args:
        item (str): path to a svg file or a directory containing svg files

    TODO: No satisfactory results yet
    Use SVG 2 PDF and then PDF 2 JPG
    """
    if isfile(item):
        pass

    elif isdir(item):
        pass

    else:
        print("Please enter a valid file or directory path")

if __name__ == "__main__":
    svg2jpg(r"C:\Users\munteanu\Downloads\_CGM_TIFF\31-17-12")