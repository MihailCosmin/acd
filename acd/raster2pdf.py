from os.path import join
from os.path import isdir
from os.path import isfile
from os.path import dirname
from os.path import abspath

from PIL import Image
Image.MAX_IMAGE_PIXELS = None
# https://stackoverflow.com/questions/51152059/pillow-in-python-wont-let-me-open-image-exceeds-limit

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from subprocess import check_output

from .filelist import list_files

def convert_image_to_pdf(input_image_path, output_pdf_path):
    # Open the image using PIL (Pillow)
    image = Image.open(input_image_path)

    # Get the dimensions of the image
    img_width, img_height = image.size

    # Create a PDF canvas with the same dimensions as the image
    c = canvas.Canvas(output_pdf_path, pagesize=(img_width, img_height))

    # Draw the image onto the PDF canvas
    c.drawImage(input_image_path, 0, 0, width=img_width, height=img_height)

    # Save the PDF file
    c.save()

def raster2pdf(item: str):
    """Converts a raster file to pdf using reportlab
    Args:
        item (str): path to a raster file or a directory containing raster files
    """
    if isfile(item):
        out_fname = item.split(".")[0] + ".pdf"
        convert_image_to_pdf(item, out_fname)
    elif isdir(item):
        for cgm in list_files(item, True, ["webp", "tif", "tiff", "jpg", "jpeg", "png", "bmp", "gif"]):
            out_fname = cgm.split(".")[0] + ".pdf"
            convert_image_to_pdf(item, out_fname)
    else:
        print("Please enter a valid file or directory path")

if __name__ == "__main__":
    # raster2pdf(r"D:\TIF2PDF\A321_SRM_INTRO_P001_S01_R00.tiff")
    pass