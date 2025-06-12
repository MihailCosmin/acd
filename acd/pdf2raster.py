from os.path import join
from os.path import isdir
from os.path import isfile
from os.path import dirname
from os.path import abspath

from subprocess import check_output

from pdf2image import convert_from_path

from .filelist import list_files

POPPLER_PATH = join(dirname(abspath(__file__)), "bin")
# Update here: https://github.com/oschwartz10612/poppler-windows/releases/

def pdf2raster(
        item: str,
        extension: str = "jpg",
        overwrite: bool = False,
        pages: list = None):
    """Converts a pdf file to raster using pdf2image

    Args:
        item (str): path to a pdf file or a directory containing pdf files
    """
    if isfile(item):
        out_fname = item.split(".")[0] + f".{extension}"
        images = convert_from_path(item, poppler_path=POPPLER_PATH)
        if not pages:
            pages = [i for i in range(1, len(images) + 1)]
        for i in range(len(images)):
            if i + 1 in pages:
                images[i].save(out_fname.replace(f".{extension}", f"_page_{i + 1}") + f".{extension}", 'JPEG')
    elif isdir(item):
        for pdf in list_files(item, True, ["pdf"]):
            out_fname = pdf.split(".")[0] + f".{extension}"
            images = convert_from_path(pdf, poppler_path=POPPLER_PATH)
            if not pages:
                pages = [i for i in range(1, len(images) + 1)]
            for i in range(len(images)):
                if i + 1 in pages:
                    images[i].save(out_fname.replace(f".{extension}", f"_page_{i + 1}") + f".{extension}", 'JPEG')
    else:
        print("Please enter a valid file or directory path")

if __name__ == "__main__":  
    pdf2raster(r"D:\HIGHLIGHTS_TEST_FOLDERS_SRM_A320\A320_P2F_SRM_GM_R03_AUG_15_23\ATA53\53-41\53-41-14_PB001_C3\A320_SRM_GM_534114_PB001_C3_IS_R03_AUG_15_23.pdf")