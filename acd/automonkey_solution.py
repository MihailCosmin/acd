from os import listdir
from os import path
from os import remove
from os import rename
from os import walk
from os.path import join
from os.path import isdir

from re import sub

from lxml import etree
from lxml.etree import tostring
from lxml.etree import fromstring

from regex import search

from time import sleep

from traceback import format_exc

from shutil import rmtree
from shutil import copyfile

from automonkey import chain
from .docx_ import replace_media
from .docx_ import get_regex_string
from .docx_ import replace_copyright
from .docx_ import adjust_column_widths
from .docx_ import get_template_version
from .archive import seven_unzip
from .archive import unarchive_file
from .archive import zip_word_folder
from tqdm import tqdm

import skimage.metrics
import skimage.transform

from glob import glob

folder_path = r"C:\Users\bakalarz\Desktop\LED-Project\docs"
image1_path = r"C:\Users\bakalarz\Desktop\LED-Project\image1.png"

word_open_image = r"C:\Users\bakalarz\Desktop\LED-Project\word_open.png"
word_save_image = r"C:\Users\bakalarz\Desktop\LED-Project\word_save.png"

# zip_word_folder(r"C:\Users\bakalarz\Desktop\LED-Project\docs\OV_P15")


def process_file(file_path: str, old_version: str, new_version: str):
    chain(
        # Replacement of the version number
        dict(startfile=file_path, wait=1),
        dict(waituntil=word_open_image, wait=1),
        dict(click=word_open_image, wait=1),
        dict(msoffice_replace=(old_version, new_version, True, True, 5), wait=1),
        dict(keys2="ctrl+s", wait=1),
        dict(keys2="alt+f4", wait=5),

        # Replacement of LIEBHERR with Liebherr
        dict(startfile=file_path, wait=1),
        dict(waituntil=word_open_image, wait=1),
        dict(msoffice_replace=("LIEBHERR", "Liebherr", True, True, 5), wait=1),
        dict(keys2="ctrl+s", wait=1),
        dict(keys2="alt+f4", wait=5),

        # Replacement of Liebherr-Elektronik GmbH with Liebherr-Electronics and Drives GmbH
        dict(startfile=file_path, wait=1),
        dict(waituntil=word_open_image, wait=1),
        dict(msoffice_replace=("Liebherr-Elektronik GmbH",
             "Liebherr-Electronics and Drives GmbH", True, True, 5), wait=1),
        dict(keys2="ctrl+s", wait=1),
        dict(keys2="alt+f4", wait=5),

        # Replacement of LEG with LED
        dict(startfile=file_path, wait=1),
        dict(waituntil=word_open_image, wait=1),
        dict(msoffice_replace=("LEG", "LED", True, True, 5), wait=1),
        dict(keys2="ctrl+s", wait=1),
        dict(keys2="alt+f4", wait=10),
        debug=True
    )


def calculate_version(old_version: str) -> str:
    parts = old_version.split("_")
    last_part = parts[-1]

    if last_part.isdigit():
        number = int(last_part)
        incremented_number = number + 1
        new_last_part = str(incremented_number).zfill(len(last_part))
        parts[-1] = new_last_part
        return "_".join(parts)
    return old_version


def calculate_image_similarity2(image1_path, image2_path):
    # Load images
    image1 = skimage.io.imread(image1_path, as_gray=True)
    try:
        image2 = skimage.io.imread(image2_path, as_gray=True)
    except SyntaxError:  # Cosmin: added this because one emf file was not correctly read
        return 0

    # Resize images to match size of image1
    image2 = skimage.transform.resize(image2, image1.shape)

    # Calculate SSIM score
    ssim_score = skimage.metrics.structural_similarity(
        image1, image2, data_range=image1.max() - image1.min())
    return ssim_score * 100


folder_cont = [file for file in listdir(
    folder_path) if not file.startswith("~$")]
print(folder_cont)
for file_num, filename in enumerate(tqdm(folder_cont), start=1):
    # Cosmin: added file_num as a way to skip some processed files
    if (filename.endswith(".docx") or filename.endswith(".docm")) and file_num > 0:
        file_path = path.join(folder_path, filename)
        word_ext = '.' + filename.split(".")[-1]

        # Get new version string
        filename_temp = filename.replace(".docx", "").replace(".docm", "")
        # Cosmin: added ver with some fixes
        ver = filename_temp.replace("-draft", "").replace("draft", "")
        old_version = get_regex_string(file_path, fr"({ver}.*)", debug=True)
        if old_version == 0:
            # Cosmin: added get_regex_string2
            old_version = get_template_version(
                file_path, fr"({ver}.*)", debug=True)

        replace_copyright(file_path, debug=False)

        if ver not in old_version:  # Cosmin added this because everything sucks
            new_version = calculate_version(ver)
        else:
            new_version = calculate_version(old_version)

        try:  # Cosmin: added try/except because it was failing for some files with IndexError
            adjust_column_widths(file_path, where="footer",
                                 col=3, difference=1.4, back_up=False, debug=False)
        except IndexError:
            print(f"IndexError for file: {filename}")
            if isdir(file_path.replace(word_ext, "")):
                rmtree(file_path.replace(word_ext, ""))

        try:
            # Cosmin added this method to unarchive the file. seven_unzip is not working for some files
            unarchive_file(file_path)
        except Exception as err:
            print(
                f"Error while unarchiving file: {filename}\n{err}\n{format_exc()}")
            seven_unzip(file_path)

        ssim_score = None
        print(listdir(r"C:\Users\bakalarz\Desktop\LED-Project\docs" +
              "\\" + filename_temp + r"\word\media"))
        for image in listdir(r"C:\Users\bakalarz\Desktop\LED-Project\docs" + "\\" + filename_temp + r"\word\media"):
            image_path = r"C:\Users\bakalarz\Desktop\LED-Project\docs" + \
                "\\" + filename_temp + r"\word\media" + "\\" + image
            image_file_name = image_path.split("\\")[-1]
            if image_path:
                # Code to check if the ssim score is above 30%
                ssim_score = calculate_image_similarity2(
                    image1_path, image_path)
                if ssim_score < 30:
                    print(
                        f"SSIM score for image1: {image1_path} and image2: {image_path} is {ssim_score} for {filename}")
                    copyfile(image_path, join(
                        r"C:\Users\bakalarz\Desktop\LED-Project\ignored_media", filename_temp + '_' + image_file_name))
                    continue
                else:
                    print(
                        f"SSIM score for image1: {image1_path} and image2: {image_path} is {ssim_score} for {filename}")
                    # Create backup of the image that will be replaced
                    copyfile(image_path, join(
                        r"C:\Users\bakalarz\Desktop\LED-Project\replaced_media", filename_temp + '_' + image_file_name))
                    extension = image_file_name.split(".")[-1]
                    new_media = r"C:\Users\bakalarz\Desktop\LED-Project" + "\\" + "image1." + extension
                    replace_media(file_path, image_file_name, new_media)
        process_file(file_path, old_version, new_version)

        rmtree(file_path.replace(word_ext, ""))

        try:
            # Cosmin added this method to unaerchive the file. seven_unzip is not working for some files
            unarchive_file(file_path)
        except Exception as err:
            print(
                f"Error while unarchiving file: {filename}\n{err}\n{format_exc()}")
            seven_unzip(file_path)
        # Replace LEG-ID with LED-ID (Hardcoded Special Case)
        with open(join(r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp, "word", "document.xml"), "r", encoding="utf-8") as _:
            document_content = _.read()
        document_content = sub(
            r"(\<w:sdtContent\>.*?w:t\>)(.*?)(LEG)(.*?)(\</w:t\>)", r"\1\2LED\4\5", document_content)
        # document_content = document_content.replace("LEG-ID", "LED-ID")
        with open(join(r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp, "word", "document.xml"), "w", encoding="utf-8") as _:
            _.write(document_content)

        # Find out which case is used:
        case = "bookmark"
        directory = join(
            r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp, "word")
        files = glob(f"{directory}/footer*.xml")
        for file in files:
            with open(file, "r", encoding="utf-8") as _:
                content = _.read()
                if "<w:sdt>" in content:
                    case = "item"
                    break
        if case == "bookmark":
            # Bookmark case
            ref_found_file = None
            ref_not_found_file = None
            for f_path in glob(f"{directory}/footer*.xml"):
                with open(f_path, "r", encoding="utf-8") as _:
                    content = _.read()
                    if search(r"(REF )(.*?)( \\h)", content):
                        ref_found_file = f_path
                    else:
                        ref_not_found_file = f_path
            if ref_found_file:
                with open(ref_found_file, "r", encoding="utf-8") as _:
                    content = _.read()
                    if search(r"(REF )(.*?)( \\h)", content):
                        found_name = search(
                            r"(REF )(.*?)( \\h)", content).group(2)
            if ref_not_found_file:
                # Insert Bookmark start and ending at right position
                with open(ref_not_found_file, "r") as _:
                    content = _.read()
                content = etree.fromstring(content.encode('utf-8'))
                tc_nodes = content.xpath(
                    ".//w:tc", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                for ind, node in enumerate(tc_nodes, start=1):
                    if ind % 2 == 0:
                        node_content = ''.join(node.itertext())
                        if new_version in node_content:
                            for wr in node.xpath('.//w:r', namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
                                if new_version in ''.join(wr.itertext()):
                                    ns_lookup = etree.ElementNamespaceClassLookup()
                                    parent = wr.getparent()
                                    new_sibling = f'<w:bookmarkStart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="18" w:name="{found_name}"/>'
                                    new_sibling = fromstring(new_sibling)
                                    parent.insert(
                                        parent.index(wr), new_sibling)

                                    parent = wr.getparent()
                                    new_sibling = '<w:bookmarkEnd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="18"/>'
                                    new_sibling = fromstring(new_sibling)
                                    parent.insert(parent.index(
                                        wr) + 1, new_sibling)
                et = etree.ElementTree(content)
                et.write(ref_not_found_file, pretty_print=True)
        else:
            # Item case
            directory_customXML = join(
                r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp, "customXML")
            with open(f"{directory_customXML}\item1.xml", "r", encoding="utf-8") as _:
                item1_content = _.read()
            item1_content = item1_content.replace(' xmlns="KID Basis"', '')
            item1_content = etree.fromstring(item1_content.encode('utf-8'))
            vorlage_node = item1_content.xpath(".//Vorlage")[0]
            vorlage_node.text = '0' + str(int(vorlage_node.text) + 1)
            item1_content = tostring(item1_content).decode("utf-8")
            item1_content = item1_content.replace(
                "<KID_Basis", '<KID_Basis xmlns="KID Basis"')
            with open(f"{directory_customXML}\item1.xml", "w", encoding="utf-8") as _:
                _.write(item1_content)

        print(join(r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp))
        for root, dirs, files in walk(join(r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp)):
            for file in files:
                if file.endswith(".xml"):
                    xml_file_path = join(root, file)
                    with open(xml_file_path, 'r', encoding="utf-8") as _:
                        xml_content = _.read()
                    xml_content = xml_content.replace(
                        '<w:pStyle w:val="TitelGro"/>', '<w:pStyle w:val="TitelKlein"/>')
                    xml_content = xml_content.replace(" LEG-", " LED-").replace(" LEG ", " LED ").replace('"LEG"', "LED").replace(
                        '"LEG-', '"LED-').replace("(LEG)", "(LED)").replace("[LEG]", "[LED]").replace("{LEG}", "{LED}")
                    with open(xml_file_path, 'w', encoding="utf-8") as _:
                        _.write(xml_content)

        zip_word_folder(
            join(r"C:\Users\bakalarz\Desktop\LED-Project\docs", filename_temp))
        rmtree(file_path.replace(word_ext, ""))

        deleted = False
        time_slept = 0  # Cosmin: if time slept is greater than 300 seconds, stop trying to delete the file and continue
        while not deleted:  # Cosmin: This might fail to find the word file to delete
            try:
                remove(file_path)
                deleted = True
            except Exception as err:
                print(f"Error, could not delete file: {err}\n{format_exc()}")
                sleep(1)
                time_slept += 1
                if time_slept > 300:
                    break
        rename(file_path.replace(word_ext, ".zip"), file_path)

        if ssim_score is None:
            print(f"Could not find image for {filename}")
