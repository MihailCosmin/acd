from os import environ
from os import pathsep
from os import replace
from os.path import join
from os.path import dirname
from os.path import abspath
from os.path import basename
from os.path import exists
from os.path import splitext
from io import BytesIO
import tempfile

from tqdm import tqdm

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

if __name__ == "__main__":
    from filelist import list_files
    from filepath import clean_path
else:
    from .filelist import list_files
    from .filepath import clean_path

pytesseract.pytesseract.tesseract_cmd = join(dirname(__file__), "3rd", "tesseract_5.5.0.20241111", "tesseract.exe")
POPPLER_PATH = join(dirname(abspath(__file__)), "3rd", "bin")
# Update here: https://github.com/oschwartz10612/poppler-windows/releases/

_ocr_instance = None  # Private module-level variable


import re
from collections import Counter
from math import ceil
from typing import List, Dict, Tuple, Optional
from pathlib import Path

# ---------- helpers used by the detector ----------

def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().upper()

def _ngrams(tokens: List[str], nmin: int, nmax: int) -> List[str]:
    out = []
    L = len(tokens)
    for n in range(nmin, min(nmax, L) + 1):
        for i in range(L - n + 1):
            out.append(" ".join(tokens[i:i+n]))
    return out

def _is_region_match(row: str, phrase: str, region: str, max_offset: int = 60) -> bool:
    idx = row.find(phrase)
    if idx < 0:
        return False
    if region == "any":
        return True
    if region == "prefix":
        return idx <= max_offset
    if region == "suffix":
        return (len(row) - (idx + len(phrase))) <= max_offset
    return True

def detect_repeated_chunks(
    rows: List[str],
    min_words: int = 4,
    max_words: int = 50,
    min_chars: int = 25,
    support_ratio: float = 0.2,
    region: str = "any",
    region_offset: int = 60,
) -> Dict[str, List[str]]:
    """
    Find repeated n-gram phrases across rows. Returns normalized phrases.
    """
    if not rows:
        return {"phrases": [], "cleaned_rows": []}

    norm_rows = [_normalize(r) for r in rows]
    tokenized = [nr.split(" ") if nr else [] for nr in norm_rows]

    per_row_ngrams = []
    for toks, nr in zip(tokenized, norm_rows):
        grams = _ngrams(toks, min_words, max_words)
        grams = {
            g for g in grams
            if len(g) >= min_chars and _is_region_match(nr, g, region, region_offset)
        }
        per_row_ngrams.append(grams)

    c = Counter()
    for grams in per_row_ngrams:
        c.update(grams)

    needed = ceil(support_ratio * len(rows))
    candidates = [g for g, cnt in c.items() if cnt >= needed]

    # keep longest unique phrases
    candidates.sort(key=lambda s: (-len(s), s))
    kept = []
    for g in candidates:
        if not any(g in k for k in kept):
            kept.append(g)

    return {"phrases": kept, "cleaned_rows": norm_rows}

# ---------- bridge from detected, normalized phrases -> regex over original text ----------

def _phrase_to_regex(phrase_norm: str) -> re.Pattern:
    """
    Convert a normalized phrase (UPPER with single spaces) to a case-insensitive
    regex that tolerates arbitrary whitespace in the original line.
    """
    # escape specials and allow flexible whitespace between tokens
    parts = [re.escape(tok) for tok in phrase_norm.split(" ")]
    pattern = r"\s*".join(parts)  # tolerate extra spaces from OCR
    # word-ish boundaries are intentionally omitted—OCR can be messy
    return re.compile(pattern, flags=re.IGNORECASE)

def _match_is_in_region(line: str, m: re.Match, region: str, max_offset: int) -> bool:
    if region == "any":
        return True
    if region == "prefix":
        return m.start() <= max_offset
    if region == "suffix":
        return (len(line) - m.end()) <= max_offset
    return True

# ---------- main API ----------

def remove_repeated_header_footer(
    filepath: str,
    output_suffix: str = "_cleaned.txt",
    *,
    support_ratio: float = 0.2,
    min_words: int = 4,
    max_words: int = 50,
    min_chars: int = 25,
    header_offset: int = 80,   # how far from start a header may begin
    footer_offset: int = 80,   # how far from end a footer may end
) -> Tuple[str, Dict[str, List[str]]]:
    """
    Detects repeated header/footer phrases and removes them from the file.
    Returns (output_path, {'headers': [...], 'footers': [...]}).
    """
    p = Path(filepath)
    text = p.read_text(encoding="utf-8", errors="replace")
    rows = [line.rstrip("\n") for line in text.splitlines()]

    # Detect headers and footers separately (on normalized rows)
    hdr = detect_repeated_chunks(
        rows, min_words, max_words, min_chars, support_ratio, region="prefix", region_offset=header_offset
    )["phrases"]
    ftr = detect_repeated_chunks(
        rows, min_words, max_words, min_chars, support_ratio, region="suffix", region_offset=footer_offset
    )["phrases"]

    header_patterns = [_phrase_to_regex(ph) for ph in hdr]
    footer_patterns = [_phrase_to_regex(ph) for ph in ftr]

    cleaned_lines = []
    for line in rows:
        orig = line

        # Remove headers near the start
        for rx in header_patterns:
            m = rx.search(orig)
            if m and _match_is_in_region(orig, m, "prefix", header_offset):
                orig = (orig[:m.start()] + orig[m.end():]).strip()

        # Remove footers near the end
        for rx in footer_patterns:
            m = rx.search(orig)
            if m and _match_is_in_region(orig, m, "suffix", footer_offset):
                orig = (orig[:m.start()] + orig[m.end():]).strip()

        # normalize spaces a bit
        orig = re.sub(r"\s+", " ", orig).strip()
        cleaned_lines.append(orig)

    cleaned_text = "\n".join(cleaned_lines)
    output_path = str(p).replace(".txt", output_suffix).replace(".TXT", output_suffix)
    Path(output_path).write_text(cleaned_text + "\n", encoding="utf-8")

    return output_path, {"headers": hdr, "footers": ftr}

def batch_remove_repeated_header_footer(
    filepath: str,
    output_suffix: str = "_cleaned.txt",
    iterations: int = 2,
    overwrite: bool = False,
    *,
    support_ratio: float = 0.2,
    min_words: int = 4,
    max_words: int = 50,
    min_chars: int = 25,
    header_offset: int = 80,   # how far from start a header may begin
    footer_offset: int = 80   # how far from end a footer may end
):
    for txt in tqdm(list_files(filepath, True, [".txt", ".TXT"]), desc="Cleaning text files", colour="green"):
        for i in range(iterations):
            output, _ = remove_repeated_header_footer(
                txt,
                output_suffix=".txt" if overwrite else output_suffix.replace(".txt", f"_{i + 1}.txt"),
                support_ratio=support_ratio,
                min_words=min_words,
                max_words=max_words,
                min_chars=min_chars,
                header_offset=header_offset,
                footer_offset=footer_offset
            )


def _initialize_paddle_ocr():
    """Initialize PaddleOCR with default parameters."""
    if __name__ == "__main__":
        from paddle_runtime import require_paddle
    else:
        from .paddle_runtime import require_paddle
    require_paddle()
    from paddleocr import PaddleOCR

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

def get_ocr_pdf_content(
        pdf: str,
        engine: str = "paddle") -> str:
    """Extract text from a PDF file using OCR.

    Args:
        pdf (str): The path to the PDF file.
        engine (str, optional): The OCR engine to use. Defaults to "paddle".
            Other options include "tesseract", "surya", and "easyocr".
        remove_header_footer (bool, optional): Whether to remove repeated headers and footers. Defaults to False.

    Returns:
        str: The extracted text from the PDF.
    """
    if __name__ == "__main__":
        from paddle_runtime import require_paddle
    else:
        from .paddle_runtime import require_paddle
    require_paddle()
    from paddleocr import PaddleOCR

    if engine == "paddle":
        ocr = PaddleOCR(
            text_detection_model_name="PP-OCRv5_mobile_det",
            text_recognition_model_name="PP-OCRv5_mobile_rec",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            lang="en+de")

    try:
        images = convert_from_path(pdf, poppler_path=POPPLER_PATH)
    except Exception as e:
        # try to get pdf text simply with tesseract
        pdf_content = pytesseract.image_to_string(pdf)

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
        with open(clean_path(output_path), "wb") as f:
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

def pdfs_to_txts(
    folder: str,
    engine: str = "paddle",
    regex: str = None,
    skip_existing: bool = False,
    remove_header_footer: bool = False) -> None:
    """Convert all PDF files in a folder to text files."""
    for pdf in tqdm(list_files(folder, True, [".pdf", ".PDF"], regex), desc="Processing PDFs", colour="green"):
        if skip_existing and exists(splitext(pdf)[0] + ".txt"):
            continue
        text = get_ocr_pdf_content(pdf, engine=engine)
        text_filename = splitext(pdf)[0] + ".txt"
        with open(text_filename, "w", encoding="utf-8") as f:
            f.write(text)
        if remove_header_footer:
            remove_repeated_header_footer(
                text_filename,
                output_suffix="",
                overwrite=True
            )


if __name__ == "__main__":
    # pdfs_to_txts(
    #     r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\777 AIPC D633W111, Rev 11 Jul 25 - W0006 - Split",
    #     regex=r"\d\d\-\d\d\-\d\d\-\d\d",
    #     skip_existing=True
    # )
    
    # with open(r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\aipc txt\53-01-01-01 - Floor Pallets - Install.txt", "r", encoding="utf-8") as f:
    #     text = f.read()
    # rows = [line for line in text.splitlines() if line.strip()]
    # # Detect a header near the line start:
    # result_hdr = detect_repeated_chunks(rows, support_ratio=0.7, region="prefix")
    # # Detect a footer near the line end:
    # result_ftr = detect_repeated_chunks(rows, support_ratio=0.7, region="suffix")
    # # Print phrases found and show a sample cleaned line:
    # print("Header phrases:", result_hdr["phrases"])
    # print("Footer phrases:", result_ftr["phrases"])

    batch_remove_repeated_header_footer(
        r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\aipc txt",
        iterations=5,
        overwrite=True
    )