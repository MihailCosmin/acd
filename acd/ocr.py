from os import environ
from os import pathsep
from os import replace
from os.path import join
from os.path import dirname
from os.path import abspath
from os.path import exists
from io import BytesIO
import tempfile

current_file_dir = dirname(abspath(__file__))
ccache_dir = join(current_file_dir, "3rd", "ccache-4.11.3-windows-x86_64")

environ["PATH"] = ccache_dir + pathsep + environ["PATH"]
environ["CCACHE"] = join(ccache_dir, "ccache.exe")
environ["PADDLE_PDX_CACHE_HOME"] = join(current_file_dir, "3rd")  # This needs to be at the top of the file

from typing import Any
import numpy as np
from PIL import Image

from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter

import pytesseract
from .paddle_runtime import require_paddle
require_paddle()
from paddleocr import PaddleOCR

pytesseract.pytesseract.tesseract_cmd = join(dirname(__file__), "3rd", "tesseract_5.4.0.20240606", "tesseract.exe")
POPPLER_PATH = join(dirname(abspath(__file__)), "3rd", "bin")
# Update here: https://github.com/oschwartz10612/poppler-windows/releases/

_ocr_instance = None  # Private module-level variable

def _initialize_paddle_ocr():
    """Initialize PaddleOCR with default parameters."""
    global _ocr_instance
    if _ocr_instance is None:
        _ocr_instance = PaddleOCR(
            text_detection_model_name="PP-OCRv5_mobile_det",
            text_recognition_model_name="PP-OCRv5_mobile_rec",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            lang="en+de"
        )
    return _ocr_instance

def get_ocr_pdf_content(pdf: str, engine: str = "paddle") -> str:
    """Extract text from a PDF file using OCR.

    Args:
        pdf (str): The path to the PDF file.
        engine (str, optional): The OCR engine to use. Defaults to "paddle".
            Other options include "tesseract", "surya", and "easyocr".

    Returns:
        str: The extracted text from the PDF.
    """
    if engine == "paddle":
        ocr = PaddleOCR(
            text_detection_model_name="PP-OCRv5_mobile_det",
            text_recognition_model_name="PP-OCRv5_mobile_rec",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            lang="en+de")

    images = convert_from_path(pdf, poppler_path=POPPLER_PATH)

    # Extract text using pytesseract
    pdf_content = ""
    for image in images:
        if engine == "tesseract":
            text = pytesseract.image_to_string(image)
        elif engine == "paddle":
            img_np = np.array(image)
            result = ocr.predict(img_np)
            text = ""
            for res in result:
                text += " ".join(res["rec_texts"]) + "\n"
        pdf_content += text

    return pdf_content

def ocr_pdf(pdf_path: str) -> None:
    """
    Perform OCR on each page and overwrite the original PDF with a searchable version.
    """
    with tempfile.TemporaryDirectory() as tempdir:
        images = convert_from_path(pdf_path, output_folder=tempdir, fmt='png', poppler_path=POPPLER_PATH)
        writer = PdfWriter()

        for img in images:
            ocr_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            ocr_reader = PdfReader(BytesIO(ocr_bytes))
            writer.add_page(ocr_reader.pages[0])

        output_path = pdf_path + ".ocr.pdf"
        with open(output_path, "wb") as f:
            writer.write(f)

        replace(output_path, pdf_path)

def ocr_image(image: Any, engine: str = "paddle") -> str:
    """
    Perform OCR on an image file and return the extracted text.

    Args:
        image (Any): The path to the image file or a PIL Image object.
        engine (str, optional): The OCR engine to use. Defaults to "paddle".
            Other options include "tesseract", "surya", and "easyocr".

    Returns:
        str: The extracted text from the image.
    """
    if engine == "tesseract":
        if exists(str(image)):
            image = Image.open(image)
        text = pytesseract.image_to_string(image)
        return text
    elif engine == "paddle":
        ocr = _initialize_paddle_ocr()
        if exists(str(image)):
            image = Image.open(image)
        img_np = np.array(image)
        result = ocr.predict(img_np)
        text = ""
        for res in result:
            text += " ".join(res["rec_texts"]) + "\n"
        return text
    return ""

if __name__ == "__main__":
    img_test = Image.open(r"C:\Users\munteanu\Downloads\CMM Automation-Drawings Summary TEST 1\CMM Automation-Drawings Summary TEST\4115-0056_Rev_04.pdf_page_1.png")
    print(
        ocr_image(
            img_test,
            "paddle"
        )
    )
