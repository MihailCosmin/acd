r"""Find stamps in PDFs and black them out.

A stamp says who accepted, approved or received a document, and when: an
approval stamp annotation, a stamp image pasted into a signature table
("Accepted by Customer"), a rubber stamp on a scan. Three kinds are found:

- **stamp annotations** (PDF /Stamp): all of them;
- **stamp images**: an image of at most `STAMP_MAX_SHARE` of the page in
  which OCR reads a stamp word (`STAMP_WORDS`: accepted, approved,
  freigegeben, geprüft, ...);
- **framed stamps on scans** and in large images: a stamp word in a frame of
  its own. A table cell is no stamp's frame -- "genehmigt / APPROVED" heads
  the columns of a drawing's approval table -- so a frame whose border runs
  on beyond its corners is left alone, and so is a stamp word in no frame.

    from acd.stamp_obfuscation import obfuscate_stamps
    output, areas = obfuscate_stamps("record.pdf", "record_nostamp.pdf")

The areas are redacted like signatures (`signature_obfuscation`): images,
vector strokes and text in them, under a black box.
"""

from os import fspath
from typing import Optional

import cv2
from numpy import array
from regex import IGNORECASE
from regex import compile as regex_compile

from .signature_obfuscation import SIGNATURE_DPI
from .signature_obfuscation import SIGNATURE_FAINT_LEVEL
from .signature_obfuscation import SIGNATURE_FILL
from .signature_obfuscation import _RULE_REACH
from .signature_obfuscation import _free_space
from .signature_obfuscation import _ink
from .signature_obfuscation import black_out_areas

# Words a stamp says, English and German.
STAMP_WORDS = regex_compile(
    r"(?<![\w-])(?:accepted|approved|released|rejected|verified|inspected|received|certified"
    r"|freigegeben|gepr(?:ü|ue|u)ft|genehmigt|abgelehnt|eingegangen|kontrolliert)(?![\w-])",
    IGNORECASE,
)

# A stamp is a short line or a few: a sentence that says "approved" is none.
STAMP_MAX_WORDS = 6

# Largest stamp (image or frame), as a share of the page.
STAMP_MAX_SHARE = 0.03

# A frame whose border runs on past a corner -- beyond the width of its own
# lines, this many points -- for this many points more is a table's.
STAMP_LINE_WIDTH = 4
STAMP_RUN_ON = 6

# Black boxes reach this far (points) beyond a stamp's frame: its ink.
STAMP_MARGIN = 2


def _stamp_lines(ocr_lines) -> list:
    """Rects (page coordinates) of the OCR lines that read like a stamp."""
    from pymupdf import Rect

    rects = []
    for line in ocr_lines or ():
        if len(line.text.split()) <= STAMP_MAX_WORDS and STAMP_WORDS.search(line.text):
            corners = line.to_image @ array([[0, line.width, line.width, 0], [0, 0, line.height, line.height],
                                             [1, 1, 1, 1]])
            rects.append(Rect(corners[0].min(), corners[1].min(), corners[0].max(), corners[1].max()))
    return rects


def _runs_on(rules, box: tuple, skip: int, reach: int) -> bool:
    """Whether a ruling line along the inside of a frame (x0, y0, x1, y1,
    pixels) runs on past one of its corners: `skip` pixels on, for `reach`."""
    height, width = rules.shape
    x0, y0, x1, y1 = box
    for y in (y0, y1 - 1):  # top and bottom border: on to the left or right?
        band = rules[max(y - skip, 0):min(y + skip + 1, height)]
        if band[:, max(x0 - skip - reach, 0):max(x0 - skip, 0)].any() or \
                band[:, min(x1 + skip, width):min(x1 + skip + reach, width)].any():
            return True
    for x in (x0, x1 - 1):  # left and right border: on up or down?
        band = rules[:, max(x - skip, 0):min(x + skip + 1, width)]
        if band[max(y0 - skip - reach, 0):max(y0 - skip, 0)].any() or \
                band[min(y1 + skip, height):min(y1 + skip + reach, height)].any():
            return True
    return False


def _own_frames(page, lines: list) -> list:
    """Rects of the smallest frame of its own around each of the lines, if there is one."""
    from pymupdf import Rect

    zoom = SIGNATURE_DPI / 72
    ink = _ink(page, zoom)
    # ruling lines, from the render alone: a scan has no others
    rules = 1 - _free_space(ink, zoom, faint=_ink(page, zoom, SIGNATURE_FAINT_LEVEL))
    height, width = ink.shape
    contours, _ = cv2.findContours(cv2.dilate(ink, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))),
                                   cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(contour) for contour in contours]
    boxes = [(x, y, x + w, y + h) for x, y, w, h in boxes if w * h <= STAMP_MAX_SHARE * width * height]
    frames = []
    skip, reach = int(STAMP_LINE_WIDTH * zoom) + 2 + _RULE_REACH, int(STAMP_RUN_ON * zoom)  # lines are drawn on
    for line in lines:
        lx0, ly0, lx1, ly1 = (value * zoom for value in tuple(line))
        around = [box for box in boxes
                  if box[0] < lx0 and box[1] < ly0 and box[2] > lx1 and box[3] > ly1]
        if not around:
            continue
        box = min(around, key=lambda box: (box[2] - box[0]) * (box[3] - box[1]))
        if not _runs_on(rules, box, skip, reach):
            frames.append(Rect(*(value / zoom for value in box)))
    return frames


def find_stamp_areas(page, ocr_lines=None) -> list:
    """Rects (page coordinates, as displayed) of the stamps on a page, but for stamp annotations.

    Args:
        page: a PyMuPDF page
        ocr_lines (list, optional): `line_ocr.OcrLine`s in page coordinates of
            what the page shows only as pixels: all of a scan, else its images
    """
    from pymupdf import Rect

    lines = _stamp_lines(ocr_lines)
    if not lines:
        return []
    page_area = abs(page.rect)
    images = [Rect(info["bbox"]) * page.rotation_matrix for info in page.get_image_info()]
    small = [image for image in images if 0 < abs(image) <= STAMP_MAX_SHARE * page_area]
    areas = []
    rest = []
    for line in lines:
        centre = (line.tl + line.br) / 2
        holding = [image for image in small if centre in image]
        if holding:
            areas.append(min(holding, key=abs))
        else:
            rest.append(line)
    areas += _own_frames(page, rest) if rest else []
    unique = []
    for area in areas:
        area = Rect(area.x0 - STAMP_MARGIN, area.y0 - STAMP_MARGIN, area.x1 + STAMP_MARGIN,
                    area.y1 + STAMP_MARGIN) & page.rect
        if not any(area in kept for kept in unique):
            unique = [kept for kept in unique if kept not in area] + [area]
    return unique


def obfuscate_stamps(filepath: str, output_path: str, find_in: Optional[str] = None, ocr: bool = True,
                     fill=SIGNATURE_FILL) -> tuple:
    """Write a copy of a PDF with its stamps blacked out and its stamp annotations deleted.

    Args:
        find_in (str, optional): the file to look for the stamps in, when
            filepath is a copy of it that was changed in place (by
            `obfuscate_document`, whose placeholders could hide a stamp word)
        ocr (bool, optional): read scans and images with OCR to find stamp
            words. Defaults to True; without it, only stamp annotations go.

    Returns:
        tuple: output_path, and [(page number, Rect)] of the areas blacked out
            (stamp annotations among them)
    """
    from pymupdf import PDF_ANNOT_STAMP
    from pymupdf import open as pdf_open

    filepath, output_path = fspath(filepath), fspath(output_path)
    find_in = fspath(find_in) if find_in is not None else filepath
    found = []
    with pdf_open(filepath) as doc, pdf_open(find_in) as original:
        if original.page_count != doc.page_count:
            raise ValueError(f"{find_in} has {original.page_count} pages, {filepath} {doc.page_count}")
        for page in doc:
            areas = []
            for annot in list(page.annots(types=[PDF_ANNOT_STAMP])):
                areas.append(annot.rect * page.rotation_matrix & page.rect)
                page.delete_annot(annot)
            if ocr:
                from .obfuscation import _pdf_ocr_lines

                source = original[page.number]
                lines = _pdf_ocr_lines(original, page.number, bool(source.get_text().strip()), False)
                areas += [area for area in find_stamp_areas(source, lines) if not any(area in kept for kept in areas)]
            if areas:
                black_out_areas(page, areas, fill)
                found += [(page.number, area) for area in areas]
        doc.save(output_path, garbage=3, deflate=True)
    return output_path, found
