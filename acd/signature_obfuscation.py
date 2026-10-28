r"""Find signatures and handwriting in table cells by their printed labels, and black them out.

A signature can be a scanned image in a PDF table, vector strokes from a pen
tablet, or ink on a scanned form; OCR cannot read it, so it cannot be found
by what it says. What can be found is where it is: in the table cell a
printed label such as "SIGNATURE", "Unterschrift" or "APPROVAL & DATE" names.

    from acd.signature_obfuscation import obfuscate_signatures
    output, areas = obfuscate_signatures("record.pdf", "record_nosig.pdf")

The cells come from the table's ruling lines, found on a render of the
page. A label that fills its cell is a column header: the cells below it,
of the same width, are the signature cells. A label with room below it in its
cell names that cell. Everything in those areas but the label itself is
redacted: images, vector strokes and text, under a black box. Pages without
text are read with OCR (`line_ocr`) to find the labels.
"""

from os import fspath
from typing import Optional

import cv2
from numpy import array
from numpy import frombuffer
from numpy import uint8
from regex import IGNORECASE
from regex import compile as regex_compile

# Printed labels that name a signature: also "Signatures & Date", "Signed by"
# -- not "Signature type", "Electronic signature" (values).
SIGNATURE_LABELS = regex_compile(
    r"(?<![\w-])(?<!electronic\s)(?:signatures?|unterschrift(?:en)?|signed(?:\s+by)?|visum)(?!\s*[-_]?\s*type)(?![\w-])",
    IGNORECASE,
)

# Words that name a signature only as part of a label ("APPROVAL & DATE",
# "Genehmigt durch:"), not alone: "Approval" is also a signature's type.
SIGNATURE_LABELS_QUALIFIED = regex_compile(r"(?<![\w-])(?:approval|approved\s+by|genehmig\w*)(?![\w-])", IGNORECASE)
SIGNATURE_QUALIFIERS = regex_compile(r"[:&]|\b(?:date|datum|by|durch)\b", IGNORECASE)

# A label is a short line: a sentence that mentions approval is no label.
SIGNATURE_LABEL_MAX_WORDS = 5

# Share of an OCR line's height above and below the text: the detector's margin.
_OCR_LABEL_INSET = 0.15

# Resolution the ruling lines are found at, and the shortest run of ink, in
# points, that counts as a ruling line.
SIGNATURE_DPI = 100
SIGNATURE_MIN_RULE = 25

# A cell is at most this share of the page: more means its lines do not close.
SIGNATURE_MAX_CELL = 0.2

# A label whose cell is at most this many times its own height is a header.
SIGNATURE_HEADER_HEIGHT = 2.2

# Cells below a header, at most.
SIGNATURE_MAX_ROWS = 30

SIGNATURE_FILL = (0, 0, 0)

# An image or curved path lying this much in signature areas is hidden whole.
SIGNATURE_PIECE_SHARE = 0.3

# Share of a cell's pixels that must be ink for it to hold something to hide.
SIGNATURE_MIN_INK = 0.004

# Gray levels below which a pixel is ink, and -- for ruling lines only, which
# copies of scanned forms often show faint -- part of a line.
SIGNATURE_INK_LEVEL = 160
SIGNATURE_FAINT_LEVEL = 215

# A run of faint ink at least this thick (pixels) is shading, not a line.
_FAINT_THICK = 6

# Ruling lines are drawn on this many pixels past their ends, to meet the
# lines a scan shows them stopping just short of.
_RULE_REACH = 4


def _ink(page, zoom: float, level: int = SIGNATURE_INK_LEVEL):
    from pymupdf import Matrix
    from pymupdf import csGRAY

    pixmap = page.get_pixmap(matrix=Matrix(zoom, zoom), colorspace=csGRAY, annots=False)
    gray = frombuffer(pixmap.samples, dtype=uint8).reshape(pixmap.height, pixmap.stride)[:, :pixmap.width]
    return (gray < level).astype(uint8)


def _free_space(ink, zoom: float, page=None, faint=None):
    """1 where no ruling line is: the cells, with whatever is written in them.

    The lines are runs of ink -- and of `faint` ink (`_ink` at
    SIGNATURE_FAINT_LEVEL) where they are thin --, and for a PDF page its
    vector lines and thin rectangles too: a signature image with a white
    background can hide the line it was pasted over.
    """
    run = max(8, int(SIGNATURE_MIN_RULE * zoom))

    def kernel(width, height):
        return cv2.getStructuringElement(cv2.MORPH_RECT, (width, height))
    rules = 0 * ink
    reach = _RULE_REACH * 2 + 1
    for along, thick, extend in ((kernel(run, 1), kernel(run, _FAINT_THICK), kernel(reach, 1)),
                                 (kernel(1, run), kernel(_FAINT_THICK, run), kernel(1, reach))):
        lines = cv2.morphologyEx(ink, cv2.MORPH_OPEN, along)
        if faint is not None:
            lines |= cv2.morphologyEx(faint, cv2.MORPH_OPEN, along) & (1 - cv2.morphologyEx(faint, cv2.MORPH_OPEN, thick))
        rules |= cv2.dilate(lines, extend)  # on along the line: a scan's line can stop short of the one it meets
    if page is not None:
        matrix = page.rotation_matrix * zoom
        for path in page.get_drawings():
            for item in path["items"]:
                if item[0] == "l":
                    segments = [(item[1], item[2])]
                elif item[0] == "re":
                    r = item[1]
                    segments = [(r.tl, r.tr), (r.tr, r.br), (r.br, r.bl), (r.bl, r.tl)]
                else:
                    continue
                for start, end in segments:
                    a, b = start * matrix, end * matrix
                    if abs(a.x - b.x) < 1.5 or abs(a.y - b.y) < 1.5:  # ruling lines are straight
                        cv2.line(rules, (int(a.x), int(a.y)), (int(b.x), int(b.y)), 1, 1)
    rules = cv2.dilate(rules, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)))
    return (1 - rules).astype(uint8)


def _cell(free, point) -> Optional[tuple]:
    """(x0, y0, x1, y1) in pixels of the cell around a point, or None if there is none."""
    x, y = int(point[0]), int(point[1])
    height, width = free.shape
    if not (0 <= x < width and 0 <= y < height) or not free[y, x]:
        return None
    _, _, _, (cx, cy, cw, ch) = cv2.floodFill(free.copy(), None, (x, y), 2, 0, 0, 4)
    if cw * ch > SIGNATURE_MAX_CELL * width * height or cx == 0 or cy == 0 or cx + cw >= width or cy + ch >= height:
        return None  # the "cell" runs off into the page: its lines do not close
    return cx, cy, cx + cw, cy + ch


def _neighbour(free, cell, side: str) -> Optional[tuple]:
    """The cell next to a cell (pixels) -- side "below", "above", "right" or
    "left" -- or None. Looked for a little further if the first try hits a
    line: a stamp's frame can lie on the border."""
    x0, y0, x1, y1 = cell
    for step in (4, 8, 12, 16):
        point = {"below": ((x0 + x1) / 2, y1 + step), "above": ((x0 + x1) / 2, y0 - step),
                 "right": (x1 + step, (y0 + y1) / 2), "left": (x0 - step, (y0 + y1) / 2)}[side]
        found = _cell(free, point)
        if found is not None:
            return found
    return None


def _lines_up(header, current, found, side: str) -> bool:
    """Whether a cell found next to `current` continues the header's column
    (side below or above: as wide) or row (right or left: as high)."""
    across = (0, 2) if side in ("below", "above") else (1, 3)
    overlap = min(found[across[1]], header[across[1]]) - max(found[across[0]], header[across[0]])
    gap = {"below": found[1] - current[3], "above": current[1] - found[3],
           "right": found[0] - current[2], "left": current[0] - found[2]}[side]
    return overlap >= 0.7 * max(found[across[1]] - found[across[0]], header[across[1]] - header[across[0]]) and gap <= 8


def _header_side(free, cell) -> str:
    """Where a header's cells are: "right" for a row header -- in the first
    column, under a cell of its width ("Unterschrift" between "Datum" and
    "Name") --, else "below"."""
    above, left = _neighbour(free, cell, "above"), _neighbour(free, cell, "left")
    if above is not None and _lines_up(cell, cell, above, "above") and \
            not (left is not None and _lines_up(cell, cell, left, "left")):
        return "right"
    return "below"


def _label_lines(page, ocr_lines=None) -> list:
    """(text, Rect) of the short lines that name a signature: from the text layer, or the OCR lines given."""
    from pymupdf import Rect

    def label_span(text):  # (0, end of the label's last word) in a line, or None
        if len(text.split()) > SIGNATURE_LABEL_MAX_WORDS:
            return None
        found = list(SIGNATURE_LABELS.finditer(text))
        if not found:
            qualified, qualifiers = SIGNATURE_LABELS_QUALIFIED.search(text), list(SIGNATURE_QUALIFIERS.finditer(text))
            found = [qualified] + qualifiers if qualified and qualifiers else []
        found += [qualifier for qualifier in SIGNATURE_QUALIFIERS.finditer(text)] if found else []
        return (0, max(match.end() for match in found)) if found else None

    labels = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", ()):
            text = "".join(span["text"] for span in line["spans"])
            if label_span(text):
                labels.append((text, Rect(line["bbox"])))
    for line in ocr_lines or ():
        span = label_span(line.text)
        if span:
            # The label's words alone, without the detector's margin: OCR's line can
            # take in a signature written beside or just under them
            corners = line.quad(*span, pad=0)
            rect = Rect(corners[:, 0].min(), corners[:, 1].min(), corners[:, 0].max(), corners[:, 1].max())
            inset = _OCR_LABEL_INSET * rect.height
            labels.append((line.text, Rect(rect.x0, rect.y0 + inset, rect.x1, rect.y1 - inset)))
    return labels


def find_signature_areas(page, ocr_lines=None) -> list:
    """Rects (page coordinates, as displayed) to black out: the cells signature labels name, without the labels.

    Args:
        page: a PyMuPDF page
        ocr_lines (list, optional): `line_ocr.OcrLine`s in page coordinates,
            for a page whose text exists only as pixels
    """
    from pymupdf import Rect

    labels = _label_lines(page, ocr_lines)
    if not labels:
        return []
    zoom = SIGNATURE_DPI / 72
    ink = _ink(page, zoom)
    free = _free_space(ink, zoom, page, _ink(page, zoom, SIGNATURE_FAINT_LEVEL))

    def written(rect) -> bool:  # is there anything in it to hide?
        x0, y0, x1, y1 = (int(value * zoom) for value in (rect.x0 + 2, rect.y0 + 2, rect.x1 - 2, rect.y1 - 2))
        return x1 > x0 and y1 > y0 and ink[y0:y1, x0:x1].mean() > SIGNATURE_MIN_INK
    # A label in two languages is two lines, one under the other, in one cell: "Unterschrift" over "SIGNATURE"
    found = []  # (cell, label)
    for _, label in sorted(labels, key=lambda item: item[1].y0):
        cell = _cell(free, ((label.x0 + label.x1) / 2 * zoom, (label.y0 + label.y1) / 2 * zoom))
        if cell is None:
            continue
        for index, (other_cell, other) in enumerate(found):
            if other_cell == cell and label.y0 - other.y1 <= 0.5 * max(label.height, other.height):
                found[index] = (cell, other | label)
                break
        else:
            found.append((cell, Rect(label)))
    areas = []
    for cell, label in found:
        x0, y0, x1, y1 = (value / zoom for value in cell)
        if y1 - y0 <= SIGNATURE_HEADER_HEIGHT * label.height:  # a header: the cells of its column, or row
            side = _header_side(free, cell)
            current = cell
            for _ in range(SIGNATURE_MAX_ROWS):
                next_cell = _neighbour(free, current, side)
                if next_cell is None or not _lines_up(cell, current, next_cell, side):
                    break
                cell_rect = Rect(*(value / zoom for value in next_cell))
                if written(cell_rect):
                    areas.append(cell_rect)
                current = next_cell
        else:  # the label names its own cell: all of it but the label
            box = Rect(x0, y0, x1, y1)
            keep = Rect(label) & box
            areas += [part for part in (Rect(x0, y0, x1, keep.y0), Rect(x0, keep.y1, x1, y1),
                                        Rect(x0, keep.y0, keep.x0, keep.y1), Rect(keep.x1, keep.y0, x1, keep.y1))
                      if part.width > 4 and part.height > 4 and written(part)]
    # a little inside the cell: its ruling lines stay
    inset = 1.5
    areas = [Rect(area.x0 + inset, area.y0 + inset, area.x1 - inset, area.y1 - inset) for area in areas
             if area.width > 2 * inset and area.height > 2 * inset]
    # an image or pen strokes mostly in a signature cell go whole: a signature runs over its cell's lines
    whole = []
    pieces = [info["bbox"] for info in page.get_image_info()]
    pieces += [path["rect"] for path in page.get_drawings() if any(item[0] == "c" for item in path["items"])]
    for piece in pieces:
        piece = Rect(piece) * page.rotation_matrix
        covered = sum((piece & area).get_area() for area in areas)
        if piece.get_area() and covered >= SIGNATURE_PIECE_SHARE * piece.get_area():
            whole.append(Rect(piece.x0 - 2, piece.y0 - 2, piece.x1 + 2, piece.y1 + 2) & page.rect)  # and its stroke width
    return areas + [piece for piece in whole if not any(piece in area for area in areas)]


def obfuscate_signatures(filepath: str, output_path: str, find_in: Optional[str] = None, ocr: bool = True,
                         fill=SIGNATURE_FILL) -> tuple:
    """Write a copy of a PDF with the signature cells blacked out.

    Args:
        find_in (str, optional): the file to look for the labels and cells in,
            when filepath is a copy of it that was changed in place (by
            `obfuscate_document`, whose placeholders could hide a label).
        ocr (bool, optional): read pages without text with OCR to find their
            labels. Defaults to True.

    Returns:
        tuple: output_path, and [(page number, Rect)] of the areas blacked out
    """
    from pymupdf import open as pdf_open

    filepath, output_path = fspath(filepath), fspath(output_path)
    find_in = fspath(find_in) if find_in is not None else filepath
    found = []
    with pdf_open(filepath) as doc, pdf_open(find_in) as original:
        if original.page_count != doc.page_count:
            raise ValueError(f"{find_in} has {original.page_count} pages, {filepath} {doc.page_count}")
        for page in doc:
            source = original[page.number]
            lines = None
            if ocr and not source.get_text().strip():
                from .obfuscation import _pdf_ocr_lines

                lines = _pdf_ocr_lines(original, page.number, False, False)
            areas = find_signature_areas(source, lines)
            if areas:
                black_out_areas(page, areas, fill)
                found += [(page.number, area) for area in areas]
        doc.save(output_path, garbage=3, deflate=True)
    return output_path, found


def black_out_areas(page, areas: list, fill=SIGNATURE_FILL) -> None:
    """Redact areas of a page (Rects as displayed) under black boxes: images, vector strokes and text in them."""
    from pymupdf import PDF_REDACT_IMAGE_NONE
    from pymupdf import PDF_REDACT_IMAGE_PIXELS
    from pymupdf import PDF_REDACT_LINE_ART_REMOVE_IF_COVERED
    from pymupdf import PDF_REDACT_TEXT_NONE
    from pymupdf import PDF_REDACT_TEXT_REMOVE
    from pymupdf import Rect

    # Pen strokes first, unfilled: MuPDF deletes a path only if a redaction covers its miter
    # joins too, which reach out ten line widths -- a black box that large would hide the
    # neighbours. This pass deletes nothing but line art it covers whole.
    strokes = []
    for path in page.get_drawings():
        rect = Rect(path["rect"]) * page.rotation_matrix
        if any(item[0] == "c" for item in path["items"]) and \
                sum((rect & area).get_area() for area in areas) >= SIGNATURE_PIECE_SHARE * rect.get_area():
            reach = 10 * (path.get("width") or 1) + 2
            strokes.append(Rect(rect.x0 - reach, rect.y0 - reach, rect.x1 + reach, rect.y1 + reach))
    for rect in strokes:
        page.add_redact_annot(rect * page.derotation_matrix, fill=False)
    if strokes:
        page.apply_redactions(images=PDF_REDACT_IMAGE_NONE, graphics=PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                              text=PDF_REDACT_TEXT_NONE)
    for area in areas:
        page.add_redact_annot(area * page.derotation_matrix, fill=fill)
    page.apply_redactions(images=PDF_REDACT_IMAGE_PIXELS, graphics=PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                          text=PDF_REDACT_TEXT_REMOVE)
