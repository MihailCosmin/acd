from os.path import join
from os.path import dirname
from os.path import abspath

from pdf2image import convert_from_path
import pytesseract
import tempfile
from pypdf import PdfReader, PdfWriter
import io
import os

pytesseract.pytesseract.tesseract_cmd = join(dirname(__file__), "3rd", "Tesseract-OCR", "tesseract.exe")
POPPLER_PATH = join(dirname(abspath(__file__)), "3rd", "bin")
# Update here: https://github.com/oschwartz10612/poppler-windows/releases/

def get_ocr_pdf_content(pdf: str) -> str:
    images = convert_from_path(pdf, poppler_path=POPPLER_PATH)

    # Extract text using pytesseract
    pdf_content = ""
    for image in images:
        text = pytesseract.image_to_string(image)
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
            ocr_reader = PdfReader(io.BytesIO(ocr_bytes))
            writer.add_page(ocr_reader.pages[0])

        output_path = pdf_path + ".ocr.pdf"
        with open(output_path, "wb") as f:
            writer.write(f)

        os.replace(output_path, pdf_path)

if __name__ == "__main__":
    ocr_pdf(r"D:\pdf\300944_LI537_Rev2.pdf")
