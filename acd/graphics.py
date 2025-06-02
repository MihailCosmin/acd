"""This module provides function to process images.
Sometimes one quality or compression is not enough.
Compressing after resizing sometimes results in a larger file. Or the white becomes gray.
For best results try out different combinations of functions.
"""
from os import remove

from tifffile import imread
from tifffile import imwrite

from PIL import Image, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None
# https://stackoverflow.com/questions/51152059/pillow-in-python-wont-let-me-open-image-exceeds-limit

from tqdm import tqdm

from .txt import get_textfile_content  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position
from .file_info import get_file_size  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position

from .constants import IMG_EXT  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position
from .constants import TIFF_COMPRESSION  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position

# Quality and compression
def to_256(img: str) -> None:
    """
    This function takes an image and returns a 256 color version of the image.

    img: str, the image to convert

    Returns: None
    """
    filename = img.split('\\')[-1]
    opened_img = Image.open(img)
    opened_img = opened_img.convert('P')
    try:
        opened_img.save(img.replace(filename, '256_' + filename))
    except OSError:
        print(f"Could not convert {img} to 256 colors.")
        return


def compress_img(img: str, tiff_compression: str = "deflate") -> None:
    """
    This function takes an image and returns a compressed version of the image.

    img: str, the image to compress
    tiff_compression: str, the compression to use for the tiff file.
    tiff_compression types: zlib, jpeg, deflate, none

    Returns: None
    """
    assert tiff_compression in TIFF_COMPRESSION, "Compression type must be one of the following: zlib, jpeg, deflate, none."
    filename = img.split('\\')[-1]
    if not img.endswith('.tiff') and not img.endswith('.tif'):
        opened_img = Image.open(img)
        opened_img.save(img.replace(filename, 'compressed_' + filename), optimize=True, quality=66)
    else:
        if "IsoDraw" not in get_textfile_content(img) or get_file_size(img) > 10:
            try:
                opened_img = imread(img)
            except ValueError:
                print(f"Could not read {img}.")
                return
            try:
                imwrite(img.replace(filename, 'compressed_' + filename), opened_img, compression=tiff_compression)
            except NotImplementedError:
                print(f"Could not compress {img}.")
                return

def resize_img(img: str, percentage: float) -> None:
    """
    This function takes an image and returns a resized version of the image.

    img: str, the image to resize
    percentage: float, the percentage to resize the image by

    Returns: None
    """
    assert isinstance(percentage, float) and 0 < percentage < 1, "Percentage must be a float between 0 and 1."
    filename = img.split('\\')[-1]
    opened_img = Image.open(img)
    opened_img = opened_img.resize((int(opened_img.size[0] * percentage), int(opened_img.size[1] * percentage)))
    opened_img.save(img.replace(filename, 'resized_' + filename))

# Style and colors
def negative(img: str) -> None:
    """
    This function takes an image and returns a negative version of the image.

    img: str, the image to convert

    Returns: None
    """
    filename = img.split('\\')[-1]
    opened_img = Image.open(img)
    opened_img = Image.open(img).convert('L')
    opened_img = Image.open(img).convert('L').point(lambda x: 255 - x)
    opened_img.save(img.replace(filename, 'negative_' + filename))

def blueprint(img: str) -> None:
    """
    This function takes an image and replaces all black pixels with blue.

    img: str, the image to convert

    Returns: None
    """
    negative(img)
    img = img.replace(img.split('\\')[-1], 'negative_' + img.split('\\')[-1])
    filename = img.split('\\')[-1]
    opened_img = Image.open(img)
    opened_img = opened_img.convert('RGB')
    for pixel_row in tqdm(range(opened_img.size[0])):
        for pixel_column in range(opened_img.size[1]):
            pixel = opened_img.getpixel((pixel_row, pixel_column))
            if pixel == (0, 0, 0):
                opened_img.putpixel((pixel_row, pixel_column), (0, 79, 190))
    opened_img.save(img.replace(filename, 'blueprint_' + filename.replace("negative_", "")))
    remove(img)

def crop_image(
        img: str,
        left: int = None,
        top: int = None,
        right: int = None,
        bottom: int = None,
        width_percentage: float = None,
        height_percentage: float = None,
        overwrite: bool = False) -> None:
    """
    This function takes an image and returns a cropped version of the image.
    NOTE: All arguments have to be passed as keyword arguments.

    img: str, the image to crop
    left: int, the left pixel to start cropping from
    top: int, the top pixel to start cropping from
    right: int, the right pixel to end cropping at
    bottom: int, the bottom pixel to end cropping at
    width_percentage: float, the percentage of the width to keep
    height_percentage: float, the percentage of the height to keep

    Returns: None
    """

    filename = img.split('\\')[-1]
    ext = filename.split('.')[-1]
    opened_img = Image.open(img)
    if left and top and right and bottom:
        opened_img = opened_img.crop((left, top, right, bottom))
    elif width_percentage and height_percentage:
        opened_img = opened_img.crop((0, 0, int(opened_img.size[0] * width_percentage), int(opened_img.size[1] * height_percentage)))
    if overwrite:
        opened_img.save(img)
    else:
        opened_img.save(img.replace(filename, filename.replace(f".{ext}", f"_cropped.{ext}")))

def get_average_color(img: str) -> tuple:
    """
    This function takes an image and returns the average color of the image.

    img: str, the image to get the average color of

    Returns: tuple, the average color of the image
    """
    opened_img = Image.open(img)
    width, height = opened_img.size

    r_total = 0
    g_total = 0
    b_total = 0

    for row in range(width):
        for column in range(height):
            r, g, b = opened_img.getpixel((row, column))
            r_total += r
            g_total += g
            b_total += b

    total_pixels = width * height
    r_average = r_total / total_pixels
    g_average = g_total / total_pixels
    b_average = b_total / total_pixels

    return (r_average, g_average, b_average)

if __name__ == "__main__":
    crop_image(
        r"D:\HIGHLIGHTS_TEST_FOLDERS_SRM_A320\A320_P2F_SRM_GM_R03_AUG_15_23\ATA53\53-41\53-41-14_PB001_C3\A320_SRM_GM_534114_PB001_C3_IS_R03_AUG_15_23_page_1.jpg",
        width_percentage=0.1,
        height_percentage=1)
    print(
        get_average_color(
            r"D:\HIGHLIGHTS_TEST_FOLDERS_SRM_A320\A320_P2F_SRM_GM_R03_AUG_15_23\ATA53\53-41\53-41-14_PB001_C3\cropped_A320_SRM_GM_534114_PB001_C3_IS_R03_AUG_15_23_page_2.jpg"
        )
    )
    print(
        get_average_color(
            r"D:\HIGHLIGHTS_TEST_FOLDERS_SRM_A320\A320_P2F_SRM_GM_R03_AUG_15_23\ATA53\53-41\53-41-14_PB001_C3\cropped_A320_SRM_GM_534114_PB001_C3_IS_R03_AUG_15_23_page_1.jpg"
        )
    )