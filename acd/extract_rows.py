"""This module will try to extract all rows as separate images.
For this we will make subimages of the original image with the complete width and 1px height.
We will check if the height average color is white (or very close to white) and if so we will add only 
coordinates of the last white row, before the next non-white row.

PROBLEM: some rows have no perfectly white line between them.
This happens in some cases when there are tall characters or above some characters that are too long.
Or for scanned documents where the characters are a bit fuzzy and the lines are not perfectly aligned.

INFO: Only as example, not imported in the module.
"""
from os import mkdir
from os.path import sep
from os.path import join
from os.path import isdir
from os.path import isfile
from os.path import expanduser

from time import sleep

from math import sqrt

from itertools import chain

from PIL import Image
Image.MAX_IMAGE_PIXELS = None
# https://stackoverflow.com/questions/51152059/pillow-in-python-wont-let-me-open-image-exceeds-limit

from numpy import mean
from tqdm import tqdm

from pyautogui import click
from pyautogui import press
from pyautogui import dragTo
from pyautogui import locateAll
from pyautogui import screenshot
from pyautogui import locateAllOnScreen

from keyboard import press
from keyboard import release
from keyboard import press_and_release as keys

from clipboard import paste
from clipboard import copy as clipboardcopy

from pytesseract import image_to_string

DELAY_DICT = {
    (0, 16): 0.5,
    (16, 27): 0.55,
    (27, 38): 0.6,
    (38, 49): 0.65,
    (49, 50): 0.7,
    (50, 61): 0.75,
    (61, 72): 0.8,
    (72, 83): 0.85,
    (83, 94): 0.9,
    (94, 105): 0.95,
    (105, 116): 1,
    (116, 127): 1.05,
    (127, 138): 1.1,
    (138, 149): 1.15,
    (149, 160): 1.2,
    (160, 171): 1.25,
    (171, 182): 1.3,
    (182, 193): 1.35,
    (193, 204): 1.4,
    (204, 215): 1.45,
    (215, 226): 1.5,
    (226, 237): 1.55,
    (237, 248): 1.6,
    (248, 259): 1.65,
    (259, 270): 1.7,
    (270, 281): 1.75,
    (281, 292): 1.8,
    (292, 303): 1.85,
    (303, 314): 1.9,
    (314, 325): 1.95,
    (325, 336): 2,
    (336, 347): 2.05,
    (347, 358): 2.1,
    (358, 369): 2.15,
    (369, 380): 2.2,
    (380, 391): 2.25,
    (391, 402): 2.3,
    (402, 413): 2.35,
    (413, 424): 2.4,
    (424, 435): 2.45,
    (435, 446): 2.5,
    (446, 457): 2.55,
    (457, 468): 2.6,
    (468, 479): 2.65,
    (479, 490): 2.7,
    (490, 501): 2.75,
    (501, 512): 2.8,
    (512, 523): 2.85,
    (523, 534): 2.9,
    (534, 545): 2.95,
    (545, 556): 3,
    (556, 567): 3.05,
    (567, 578): 3.1,
    (578, 589): 3.15,
    (589, 600): 3.2,
    (600, 611): 3.25,
    (611, 622): 3.3,
    (622, 633): 3.35,
    (633, 644): 3.4,
    (644, 655): 3.45,
    (655, 666): 3.5,
    (666, 677): 3.55,
    (677, 688): 3.6,
    (688, 699): 3.65,
    (699, 710): 3.7,
    (710, 721): 3.75,
    (721, 732): 3.8,
    (732, 743): 3.85,
    (743, 754): 3.9,
    (754, 765): 3.95,
    (765, 776): 4,
    (776, 787): 4.05,
    (787, 798): 4.1,
    (798, 809): 4.15,
    (809, 820): 4.2,
    (820, 831): 4.25,
    (831, 842): 4.3,
    (842, 853): 4.35,
    (853, 864): 4.4,
    (864, 875): 4.45,
    (875, 886): 4.5,
    (886, 897): 4.55,
    (897, 908): 4.6,
    (908, 919): 4.65,
    (919, 930): 4.7,
    (930, 941): 4.75,
    (941, 952): 4.8,
    (952, 963): 4.85,
    (963, 974): 4.9,
    (974, 985): 4.95,
    (985, 996): 5,
    (996, 1007): 5.03,
    (1007, 1018): 5.05,
    (1018, 1029): 5.08,
    (1029, 1040): 5.1,
    (1040, 1051): 5.13,
    (1051, 1062): 5.15,
    (1062, 1073): 5.18,
    (1073, 1084): 5.2,
    (1084, 1095): 5.23,
    (1095, 1106): 5.25,
    (1106, 1117): 5.28,
    (1117, 1128): 5.3,
    (1128, 1139): 5.33,
    (1139, 1150): 5.35,
    (1150, 1161): 5.38,
    (1161, 1172): 5.4,
    (1172, 1183): 5.43,
    (1183, 1194): 5.45,
    (1194, 1205): 5.48,
    (1205, 1216): 5.5,
    (1216, 1227): 5.53,
    (1227, 1238): 5.55,
    (1238, 1249): 5.58,
    (1249, 1260): 5.6,
    (1260, 1271): 5.63,
    (1271, 1282): 5.65,
    (1282, 1293): 5.68,
    (1293, 1304): 5.7,
    (1304, 1315): 5.73,
    (1315, 1326): 5.75,
    (1326, 1337): 5.78,
    (1337, 1348): 5.8,
    (1348, 1359): 5.83,
    (1359, 1370): 5.85,
    (1370, 1381): 5.88,
    (1381, 1392): 5.9,
    (1392, 1403): 5.93,
    (1403, 1414): 5.95,
    (1414, 1425): 5.98,
    (1425, 1436): 6.01,
    (1436, 1447): 6.03,
    (1447, 1458): 6.06,
    (1458, 1469): 6.08,
    (1469, 1480): 6.11,
    (1480, 1491): 6.13,
    (1491, 1502): 6.16,
    (1502, 1513): 6.18,
    (1513, 1524): 6.21,
    (1524, 1535): 6.23,
    (1535, 1546): 6.26,
    (1546, 1557): 6.28,
    (1557, 1568): 6.31,
    (1568, 1579): 6.33,
    (1579, 1590): 6.36,
    (1590, 1601): 6.38,
    (1601, 1612): 6.41,
    (1612, 1623): 6.43,
    (1623, 1634): 6.46,
    (1634, 1645): 6.48,
    (1645, 1656): 6.51,
    (1656, 1667): 6.53,
    (1667, 1678): 6.56,
    (1678, 1689): 6.58,
    (1689, 1700): 6.61,
    (1700, 1711): 6.63,
    (1711, 1722): 6.66,
    (1722, 1733): 6.68,
    (1733, 1744): 6.71,
    (1744, 1755): 6.73,
    (1755, 1766): 6.76,
    (1766, 1777): 6.78,
    (1777, 1788): 6.81,
    (1788, 1799): 6.83,
    (1799, 1810): 6.86,
    (1810, 1821): 6.88,
    (1821, 1832): 6.91,
    (1832, 1843): 6.93,
    (1843, 1854): 6.96,
    (1854, 1865): 6.98,
    (1865, 1876): 7.01,
    (1876, 1887): 7.03,
    (1887, 1898): 7.06,
    (1898, 1909): 7.08,
    (1909, 1920): 7.11,
    (1920, 1931): 7.13,
    (1931, 1942): 7.16,
    (1942, 1953): 7.18,
    (1953, 1964): 7.21,
    (1964, 1975): 7.23,
    (1975, 1986): 7.26,
    (1986, 1997): 7.28,
    (1997, 2008): 7.31,
    (2008, 2019): 7.33,
    (2019, 2030): 7.36,
    (2030, 2041): 7.38,
    (2041, 2052): 7.41,
    (2052, 2063): 7.43,
    (2063, 2074): 7.46,
    (2074, 2085): 7.48,
    (2085, 2096): 7.51,
}

def copy_pdf_column(from_xy: tuple, to_xy: tuple, delay: int = None) -> str:
    """Copy text from pdf. From xy tuple to xy tuple

    Args:
        from_xy (tuple): Top left corner
        to_xy (tuple): Bottom right corner

    Returns:
        str: String with the copied text
    """
    # print(f"to_xy[0]: {to_xy[0]}")
    # print(f"from_xy[0]: {from_xy[0]}")
    # print(f"to_xy[1]: {to_xy[1]}")
    # print(f"from_xy[1]: {from_xy[1]}")

    diag = sqrt((to_xy[1] - from_xy[1]) * (to_xy[1] - from_xy[1]) + (to_xy[0] - from_xy[0]) * (to_xy[0] - from_xy[0]))
    # print(f"diag: {diag}")
    for pair in DELAY_DICT.items():
        if diag >= pair[0][0] and diag < pair[0][1]:
            delay = pair[1] if delay is None else delay
            break

    # delay = (to_xy[1] - from_xy[1]) / 20 if delay is None else delay
    click(from_xy[0], from_xy[1])
    press("alt")
    sleep(0.5)
    dragTo(to_xy[0], to_xy[1], delay, button="left")
    release("alt")

    clipboardcopy("§ß&%$")
    clipboardcopy("§ß&%$")
    clipboardcopy("§ß&%$")
    paste_var = paste().strip()
    wait = 0.1
    while paste_var == "§ß&%$":
        keys("ctrl+c")
        paste_var = paste().strip()
        wait += 0.1
    return paste_var


def extract_rows_from_page(
        page_img: str,
        white: int = 255,
        output_folder: str = None,
        debug: bool = False) -> tuple:
    """This function will extract all rows from a page image
    It will split the images on white lines

    Args:
        page_img (str): filepath of the page image
        white (int, optional): white value. Defaults to 255.
            You can reduce this, but it will result in false positives.
        output_folder (str, optional): folder to save the images to.
            Defaults to None, in which case it will save to desktop in a content_rows folder.
        debug (bool, optional): If True, it will print some debug info. Defaults to False.

    Returns:
        tuple: (number of sub-images, list of row horizontal start coordinate)

    """
    page_img_name = page_img.split(sep)[-1]

    output_folder = join(expanduser("~/Desktop"), "content_rows") if output_folder is None else output_folder
    if not isdir(output_folder):
        mkdir(output_folder)

    page_img = Image.open(page_img)

    if debug:
        print(f"Image size is: {page_img.size}")

    width, height = page_img.size
    white_rows = []
    last_row_white = False
    for h_pix in tqdm(range(0, height)):
        row_img = page_img.crop((0, h_pix, width, h_pix + 1))
        if mean(row_img) >= white:
            last_row_white = True
        elif last_row_white:
            # Here we will make sure we keep only the last white row
            white_rows.append(h_pix)
            last_row_white = False
    if debug:
        print(f"white_rows: {white_rows}")
        print(f"len white_rows: {len(white_rows)}")
    for ind, white_row in enumerate(white_rows):
        try:
            # white_row - 2, adds a bit of space before the content row
            row_content_img = page_img.crop((0, white_row - 2, width, white_rows[ind + 1]))
            row_content_img.save(
                join(output_folder, f"{page_img_name}_row_{ind}.png")
            )
        except IndexError:
            row_content_img = page_img.crop((0, white_row - 2, width, height))
            row_content_img.save(
                join(output_folder, f"{page_img_name}_row_{ind}.png")
            )

    return len(white_rows), white_rows

def ste_dict_rows(
        page_img: str,
        output_folder: str = None,
        left: str = False,
        dictionary: list = None,
        debug: bool = False) -> tuple:
    """
    This function will try to extract all rows as separate images.
    
    Args:
        page_img (str): filepath of the page image
        output_folder (str, optional): folder to save the images to.
            Defaults to None, in which case it will save to desktop in a content_rows folder.
        left (str, optional): If True, different coordinates will be used to locate the lines.
            Defaults to False.
        dictionary (list, optional): dictionary with results to be updated.
            Defaults to None.
        debug (bool, optional): If True, it will print some debug info. Defaults to False.

        
    """
    page_img_name = page_img.split(sep)[-1]

    output_folder = join(expanduser("~/Desktop"), "content_rows") if output_folder is None else output_folder
    if not isdir(output_folder):
        mkdir(output_folder)

    page_img2 = Image.open(page_img)

    if debug:
        print(f"Image size is: {page_img2.size}")

    width, height = page_img2.size

    positions = []

    for pos in chain(
        locateAll(join(expanduser("~/Desktop"), "blue_line.png"), page_img2),  # , confidence=0.9
        locateAll(join(expanduser("~/Desktop"), "black_line_intermediate_1.png"), page_img2),  # , confidence=0.9
        locateAll(join(expanduser("~/Desktop"), "black_line_intermediate_2.png"), page_img2),  # , confidence=0.9
        # locateAll(join(expanduser("~/Desktop"), "black_line_last.png"), page_img2, confidence=0.9),
    ):
        positions.append(pos[1])
    positions.sort()
 
    # print(f"positions: {positions}")
 
    # pytesseract_conf = "--psm 6"  # v2
    # pytesseract_conf = "--psm 6 --oem 1"  # v5
    # pytesseract_conf = "--psm 6 --oem 1 -c preserve_interword_spaces=1"  # v6
    pytesseract_conf = "--psm 6 --oem 1 -c preserve_interword_spaces=1 tosp_min_sane_kn_sp=2.8 -l spa"  # v7
    # pytesseract_conf = "--psm 13 --oem 1 -c tessedit_char_whitelist=ABCDEFG0123456789"
    # pytesseract_conf = "--psm 10 --oem 3"  # v4
    # pytesseract_conf = "--psm 13 --oem 1"  # v3
 
    if debug:
        print(f"white_rows: {positions}")
        print(f"len white_rows: {len(positions)}")
    
    for ind, white_row in enumerate(positions):
        try:
            # white_row - 2, adds a bit of space before the content row
            row_content_img = page_img2.crop((0, white_row - 2, width, positions[ind + 1]))
            row_content_img.save(
                join(output_folder, f"{page_img_name}_row_{ind}.png")
            )
            
            row_height = positions[ind + 1] - white_row
            
            if left:
                word = row_content_img.crop((9, 1, 128, row_height))
            else:
                word = row_content_img.crop((25, 3, 150, row_height))
            word_text = image_to_string(word, lang='eng', config=pytesseract_conf)
            word.save(
                join(output_folder, f"{page_img_name}_row_{ind}_word.png")
            )
            if left:
                meaning = row_content_img.crop((129, 1, 268, row_height))
            else:
                meaning = row_content_img.crop((150, 3, 290, row_height))
            meaning_text = image_to_string(meaning, lang='eng', config=pytesseract_conf)
            meaning.save(
                join(output_folder, f"{page_img_name}_row_{ind}_meaning.png")
            )
            if left:
                ex1 = row_content_img.crop((269, 1, 410, row_height))
            else:
                ex1 = row_content_img.crop((290, 3, 435, row_height))
            ex1_text = image_to_string(ex1, lang='eng', config=pytesseract_conf)
            ex1.save(
                join(output_folder, f"{page_img_name}_row_{ind}_ex1.png")
            )
            if left:
                ex2 = row_content_img.crop((411, 1, width, row_height))
            else:
                ex2 = row_content_img.crop((436, 3, width, row_height))
            ex2_text = image_to_string(ex2, lang='eng', config=pytesseract_conf)
            ex2.save(
                join(output_folder, f"{page_img_name}_row_{ind}_ex2.png")
            )
            dictionary.append((word_text, meaning_text, ex1_text, ex2_text))
            
        except IndexError:
            row_content_img = page_img2.crop((0, white_row - 2, width, height))
            row_content_img.save(
                join(output_folder, f"{page_img_name}_row_{ind}.png")
            )
    return dictionary

def pdf_to_dictionary(first_page: int, last_page: int, output_folder: str, top_x: int, top_y: int, bottom_x: int, bottom_y: int):
    sleep(2)
    dictionary = []
    if not isfile(join(expanduser("~/Desktop"), "pdf_dictionary_part.txt")):
        with open(join(expanduser("~/Desktop"), "pdf_dictionary_part.txt"), "w") as f:
            f.write("[\n")
    
    page = first_page
    while page < last_page + 1:
        try:
            sleep(0.5)
            page_img = screenshot(region=(top_x + 40, top_y, 14, bottom_y - top_y))
            # page_img.save(join(expanduser("~/Desktop"), f"page_{page}.png"))

            width, height = page_img.size

            positions = []

            for pos in chain(
                locateAll(join(expanduser("~/Desktop"), "short_blue_line.png"), page_img),  # , confidence=0.9
                locateAll(join(expanduser("~/Desktop"), "short_black_line.png"), page_img),  # , confidence=0.9
            ):
                if (pos[1] + top_y) not in positions:
                    positions.append(pos[1] + top_y)
            positions.sort()
            print(positions)
            
            for ind, pos in enumerate(positions[:-1]):
                if page % 2 == 0:
                    if ind < len(positions) - 1:
                        word = copy_pdf_column((top_x + 12, pos + 3), (top_x + 126, positions[ind + 1] - 1))
                    else:
                        word = copy_pdf_column((top_x + 12, pos + 3), (top_x + 126, bottom_y))
                else:
                    if ind < len(positions) - 1:
                        word = copy_pdf_column((top_x + 34, pos + 3), (top_x + 148, positions[ind + 1] - 1))
                    else:
                        word = copy_pdf_column((top_x + 34, pos + 3), (top_x + 148, bottom_y))
                if page % 2 == 0:
                    if ind < len(positions) - 1:
                        meaning = copy_pdf_column((top_x + 130, pos + 3), (top_x + 266, positions[ind + 1] - 1))
                    else:
                        meaning = copy_pdf_column((top_x + 130, pos + 3), (top_x + 266, bottom_y))
                else:
                    if ind < len(positions) - 1:
                        meaning = copy_pdf_column((top_x + 154, pos + 3), (top_x + 290, positions[ind + 1] - 1))
                    else:
                        meaning = copy_pdf_column((top_x + 154, pos + 3), (top_x + 290, bottom_y))
                if page % 2 == 0:
                    if ind < len(positions) - 1:
                        ex1 = copy_pdf_column((top_x + 275, pos + 3), (top_x + 412, positions[ind + 1] - 1))
                    else:
                        ex1 = copy_pdf_column((top_x + 275, pos + 3), (top_x + 412, bottom_y))
                else:
                    if ind < len(positions) - 1:
                        ex1 = copy_pdf_column((top_x + 298, pos + 3), (top_x + 432, positions[ind + 1] - 1))
                    else:
                        ex1 = copy_pdf_column((top_x + 298, pos + 3), (top_x + 432, bottom_y))
                if page % 2 == 0:
                    if ind < len(positions) - 1:
                        ex2 = copy_pdf_column((top_x + 415, pos + 3), (bottom_x, positions[ind + 1] - 1))
                    else:
                        ex2 = copy_pdf_column((top_x + 415, pos + 3), (bottom_x, bottom_y))
                else:
                    if ind < len(positions) - 1:
                        ex2 = copy_pdf_column((top_x + 438, pos + 3), (bottom_x, positions[ind + 1] - 1))
                    else:
                        ex2 = copy_pdf_column((top_x + 438, pos + 3), (bottom_x, bottom_y - 1))
                
                word = word.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("  ", " ")
                meaning = meaning.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("  ", " ")
                ex1 = ex1.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("  ", " ")
                ex2 = ex2.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("  ", " ")
                
                dictionary.append((word, meaning, ex1, ex2))
                with open(join(expanduser("~/Desktop"), "pdf_dictionary_part.txt"), "a") as f:
                    f.write(str((word, meaning, ex1, ex2)) + ",\n")

            page += 1
            sleep(0.5)
            press("pagedown")
        except KeyboardInterrupt:
            break
        except Exception as e:
            print(f"There was an error: {e}")
            break
         
            
    print("Finished")
    with open(join(expanduser("~/Desktop"), "pdf_dictionary_part.txt"), "a") as f:
        f.write("\n]")
    with open(join(expanduser("~/Desktop"), "pdf_dictionary.txt"), "w") as f:
        f.write(str(dictionary))

def pdf_page_to_img(first_page: int, last_page: int, output_folder: str, top_x: int, top_y: int, bottom_x: int, bottom_y: int):
    sleep(2)
    page = first_page
    while page < last_page + 1:
        sleep(0.5)
        page_img = screenshot(region=(top_x, top_y, bottom_x - top_x, bottom_y - top_y))
        page_img.save(join(output_folder, f"page_{page}.png"))
        page += 1
        # page_down
        press("pagedown")
    print("Finished taking screenshots")

if __name__ == "__main__":
    
    # dictionary = []
    
    # for img in list_files(join(expanduser("~/Desktop"), "STE_dictionary"), True, include_tqdm=True):
    #     filename = img.split(sep)[-1]
    #     number_in_filename = int(filename.split("_")[-1].split(".")[0])
    #     if number_in_filename % 2 == 0:
    #         dictionary = ste_dict_rows(img, left=True, dictionary=dictionary)
    #     else:
    #         dictionary = ste_dict_rows(img, left=False, dictionary=dictionary)
        
    # with open(join(expanduser("~/Desktop"), "dictionary.txt"), "w") as f:
    #     f.write(str(dictionary))
    
    # ste_dict_rows(
    #     join(
    #         expanduser("~/Desktop"),
    #         "STE_dictionary",
    #         "page_423.png"  # odd = right, even = left
    #     ),
    #     left=False,
    # )
    
    # pdf_page_to_img(
    #     137,
    #     424,
    #     join(
    #         expanduser("~/Desktop"),
    #         "STE_dictionary"
    #     ),
    #     540,
    #     165,
    #     1126,
    #     896
    # )

    pdf_to_dictionary(
        167,
        # 424,
        join(
            expanduser("~/Desktop"),
            "STE_dictionary"
        ),
        540,  # was 540  # for row detection needed 580
        165,
        1126,  # was 1126  # for row detection needed 594
        896
    )
