"""This module provides functions for PDF processing
"""
import sys

from os.path import join
from os.path import basename
from os.path import dirname

from fitz import open as pdf_open
from pikepdf import open as pike_open
if sys.version_info >= (3, 12):
    from pypdf import PdfWriter as PdfFileWriter
    from pypdf import PdfReader as PdfFileReader
else:
    from PyPDF4 import PdfFileWriter, PdfFileReader

from pdfreader import SimplePDFViewer
from .filelist import list_files

def pdf_page_count(file_path, engine: str = "pypdf") -> int:
    """
    Returns the number of pages in a PDF file.

    Args:
        file_path (str): The path to the PDF file.

    Returns:
        int: The number of pages in the PDF file.
    """
    try:
        pdf_pages = 0
        if engine == "pypdf":
            with open(file_path, 'rb') as _:
                pdf = PdfFileReader(_, strict=False)
                pdf_pages = pdf.getNumPages()
        elif engine in ("fitz", "pymupdf"):
            with pdf_open(file_path) as _:
                pdf_pages = _.page_count
        elif engine == "pdfreader":
            with open(file_path, 'rb') as _:
                viewer = SimplePDFViewer(_)
                viewer.render()
                pdf_pages = viewer.doc.page_count
        return pdf_pages
    except Exception as err:
        # See which exceptions can occur then add them here.
        print(err)
        return 0


def get_pdf_content(file_path: str, engine: str = "fitz") -> str:
    """
    Returns the content of a PDF file.

    Args:
        file_path (str): The path to the PDF file.
        engine (str): The engine to use. Choices: "fitz", "pymupdf", "pypdf", "pdfreader", "pikepdf"

    Returns:
        list: The content of the PDF file.
    """
    try:
        if engine == "pypdf":
            with open(file_path, 'rb') as _:
                pdf_content = str(PdfFileReader(_, strict=False).getPage(0).extractText())
            return pdf_content
        elif engine in ("fitz", "pymupdf"):
            with pdf_open(file_path) as _:
                pdf_content = ""
                for page in _:
                    pdf_content += page.get_text()
            return pdf_content
        elif engine == "pdfreader":
            with open(file_path, 'rb') as _:
                viewer = SimplePDFViewer(_)
                viewer.render()
            return viewer.canvas.strings
        elif engine == "pikepdf":
            with pike_open(file_path) as _:
                pdf_content = ""
                for page in _:
                    pdf_content += page.get_text()
            return pdf_content
    except Exception as err:
        # See which exceptions can occur then add them here.
        print(err)
        return None
    return None

def get_pdf_metadata(file_path: str, engine: str = "fitz", pike_meta: bool = False) -> dict:
    """
    Returns the metadata of a PDF file.

    Args:
        file_path (str): The path to the PDF file.
        engine (str): The engine to use. Choices: "fitz", "pymupdf", "pypdf", "pypdf"<, "pypdf", "pikepdf", "PikePDF", "PIKEPDF"
        pike_meta (bool): If True, returns the metadata of the PDF file as a dict.
                          If False, returns the metadata of the PDF file as the XMP metadata stream.

    Returns:
        dict: The metadata of the PDF file.
    """
    try:
        if engine in ("PyPDF", "PYPDF", "pypdf"):
            with open(file_path, 'rb') as _:
                return PdfFileReader(_, strict=False).getDocumentInfo()
        elif engine in ("fitz", "pymupdf"):
            with pdf_open(file_path) as _:
                return _.metadata
        elif engine in ("pikepdf", "PikePDF", "PIKEPDF"):
            with pike_open(file_path) as _:
                if not pike_meta:
                    # The Document Info block is an older, now deprecated object in which metadata may be stored.
                    # It is not recommended to use it anymore, but it is still supported.
                    pike_dict = {}
                    for key, value in dict(_.docinfo).items():
                        pike_dict[str(key)] = str(value)
                    return pike_dict
                # For newer versions of the PDF format, the metadata is stored in the root object.
                # Try with pike_meta set to true
                return _.open_metadata()
    except Exception as err:
        # See which exceptions can occur then add them here.
        print(err)
        return None
    return None

def merge_pdfs(folder: str, output_name: str, debug: bool = False) -> int:
    """
    Merges all PDF files in a folder into one PDF file.

    Args:
        folder (str): The folder containing the PDF files to merge.
        output_name (str): The name of the output file.
        debug (bool): If True, prints debug messages.

    Returns:
        int: 0 if successful, 1 if not.
    """
    try:
        pdf_writer = PdfFileWriter()
        for file in list_files(folder, True):
            if file.lower().endswith(".pdf"):
                pdf_reader = PdfFileReader(file, strict=False)
                for page in range(pdf_reader.getNumPages()):
                    pdf_writer.addPage(pdf_reader.getPage(page))
        with open(join(folder, output_name), 'wb') as _:
            pdf_writer.write(_)
    except Exception as err:
        # See which exceptions can occur then add them here.
        print(err)
        return 1
    return 0

if __name__ == "__main__":
    merge_pdfs(r"D:\Excel Tests", "merged.pdf")