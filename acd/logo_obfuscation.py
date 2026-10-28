r"""Find logos in PDFs and images by their similarity to reference images, and remove them.

A logo comes in many forms: an embedded image in a header, vector outlines in
a drawing's title block, pixels on a scanned page. Reference images of the
logo -- crops of each variant, e.g. outlined, solid and with a slogan -- are
searched for on every page at many sizes with normalized cross-correlation,
on the ink of the page (dark or coloured against light), so colour, line
weight and scan noise matter little. Every candidate found on the coarse
render is checked again on a sharp one.

    from acd.logo_obfuscation import load_logo_references, obfuscate_logos
    references = load_logo_references(["logo_solid.png", "logo_outline.png"])
    output, found = obfuscate_logos("drawing.pdf", references, "drawing_nologo.pdf")

In a PDF, a found logo is redacted: image pixels under it are blanked, vector
paths inside it and text over it are deleted, and a box is drawn in its
place. In an image file, the box is painted over the pixels.
`python -m acd.obfuscation_ai apply --logos` runs it on the output of
`obfuscation.obfuscate_document`, which has already cleaned the metadata.
"""

from collections import namedtuple
from os import fspath
from os.path import basename
from os.path import splitext
from typing import Optional

import cv2
from numpy import array
from numpy import ascontiguousarray
from numpy import float32
from numpy import frombuffer
from numpy import nan_to_num
from numpy import percentile
from numpy import uint8
from numpy import where

# Coarse search: pages are rendered so that their longer side has at most this
# many pixels (A4 at 100 dpi, A0 at about 50 dpi).
LOGO_SEARCH_PIXELS = 2400
LOGO_SEARCH_DPI = 100

# Logo heights searched for: in points on a PDF page, in pixels of the
# coarse render in an image file; and the step between two sizes (each next
# size is this much larger). On a PDF page the letters of a found logo must
# be at least LOGO_MIN_POINTS high (2.8 mm): smaller is a line of text.
LOGO_MIN_POINTS = 8
LOGO_MAX_POINTS = 72
LOGO_MIN_HEIGHT = 12
LOGO_MAX_HEIGHT = 240
LOGO_SCALE_STEP = 1.1

# Correlation a candidate needs on the coarse render, and on the sharp one.
LOGO_CANDIDATE_SCORE = 0.5
LOGO_MAX_CANDIDATES = 200  # per reference and page
LOGO_SCORE = 0.6

# Sharp render: the candidate is rendered so that the logo is this many
# pixels high.
LOGO_VERIFY_HEIGHT = 60

# Longest side of the sharp page render the candidates are cropped from
# (an A0 sheet then has 2.7 pixels per point; A4 gets 7.5).
LOGO_SHARP_PIXELS = 9000

# Blur on the sharp render, in its pixels: little, so that only the shapes of
# the letters match, not any dark band of text, but enough for outlines drawn
# a little differently (2.5 found a drawing's own outlined wordmark at 0.61).
LOGO_VERIFY_BLUR = 2.5

# Rotations searched (degrees counter-clockwise). Each one costs as much as
# the upright search; add 90 and 270 for turned title blocks and scans.
LOGO_ROTATIONS = (0,)

# How much more or less of its box a found logo may ink than the reference:
# a line of text correlates with a wordmark, but inks less of its box.
LOGO_INK_RATIO = 1.6

# Widths searched, as a share of the reference's: drawings print the same
# wordmark with taller letters.
LOGO_ASPECTS = (1.0, 0.72)

# A found logo may have at most this share of mid-tones in its box -- ink
# between paper and the logo's own colour (photos have many; a blue logo is
# all one tone) -- and at most this much ink on average around it (on paper).
LOGO_MAX_MIDTONES = 0.3
LOGO_MAX_SURROUNDING_INK = 0.35

# A wordmark is letters: no one connected shape in it covers this share of its box.
LOGO_MAX_SHAPE_SHARE = 0.4

# Margin around a found logo, as a share of its height.
LOGO_MARGIN = 0.12

LOGO_FILL = (0, 0, 0)

# Variance of ink a page area needs to be compared at all (blank paper has 0).
LOGO_MIN_VARIANCE = 0.002

# ink: the logo's ink (0 paper .. 1 ink), cropped; letters: the share of its
# height its tallest connected shape (a letter) takes, 1.0 for a plain wordmark;
# words: what OCR reads in it ("AIRBUS", "HELICOPTERS"), empty if nothing
LogoReference = namedtuple("LogoReference", "name ink letters words", defaults=(frozenset(),))

# A match the page's text (or OCR) reads as other words than its reference
# is text in a similar typeface, not the logo: a word is like a reference
# word at this similarity (difflib's ratio, which allows for misreadings).
LOGO_WORD_SIMILARITY = 0.7

# A reference whose letters take less of its height than this is several lines
# of text (a name over a slogan): its matches must read like it.
LOGO_LINES_SHARE = 0.7
LogoMatch = namedtuple("LogoMatch", "page box score reference rotation")


def load_logo_references(paths, read: bool = True) -> list:
    """LogoReference for each image file: its ink, cropped to the logo, and
    -- with `read`, if OCR (`line_ocr`) is installed -- the words in it."""
    from PIL import Image

    references = []
    for path in paths:
        path = fspath(path)
        with Image.open(path) as image:
            image = image.convert("RGBA")
            white = Image.new("RGBA", image.size, (255, 255, 255, 255))
            gray = Image.alpha_composite(white, image).convert("L")
        ink = _crop_to_ink(_ink(gray))
        references.append(LogoReference(basename(path), ink, _letter_share(ink), _read_words(gray) if read else frozenset()))
    return references


def _words(text: str) -> set:
    """The words of a text worth comparing: three letters or digits and more, upper case."""
    return {word for word in "".join(char if char.isalnum() else " " for char in text.upper()).split() if len(word) >= 3}


def _read_words(gray) -> frozenset:
    """The words OCR reads in a reference image (on a white margin, at least 120 pixels high), or none."""
    try:
        from .line_ocr import ocr_lines
    except ImportError:
        return frozenset()
    from numpy import asarray
    from PIL import ImageOps

    scale = max(1.0, 120 / gray.height)
    image = ImageOps.expand(gray.resize((round(gray.width * scale), round(gray.height * scale))), border=40, fill=255)
    try:
        return frozenset(word for line in ocr_lines(asarray(image)) for word in _words(line.text))
    except Exception:  # OCR is a help here, not a must
        return frozenset()


def _like(words: set, reference) -> bool:
    """Is one of the words like one of the reference's (allowing for misreadings)?"""
    from difflib import SequenceMatcher

    return any(word in known or known in word or SequenceMatcher(None, word, known).ratio() >= LOGO_WORD_SIMILARITY
               for word in words for known in reference.words)


def _reads_otherwise(page, rect, reference, ocr_lines=None) -> bool:
    """Is the match text that reads otherwise than its reference?

    A word like the reference's ("LIEBHERR") keeps it: then it is the logo,
    read. Else yes, if text -- the page's words, or OCR lines -- covers at
    least half of it. For a logo of several lines of text ("AIRBUS" over
    "HELICOPTERS"), which hatching and two lines of bold type resemble, yes
    also if any other word is in it, or if OCR read the area (`ocr_lines`
    given, and a page without text or an image on one) and found nothing
    like the logo's words. A reference OCR could not read keeps every match.
    """
    from pymupdf import Rect

    if not reference.words or not rect.get_area():
        return False
    pieces = [(Rect(x0, y0, x1, y1), _words(word)) for x0, y0, x1, y1, word, *_ in page.get_text("words", clip=rect)]
    for line in ocr_lines or ():
        corners = line.to_image @ array([[0, line.width, line.width, 0], [0, 0, line.height, line.height], [1, 1, 1, 1]])
        pieces.append((Rect(corners[0].min(), corners[1].min(), corners[0].max(), corners[1].max()), _words(line.text)))
    covered = 0.0
    other = False
    for box, words in pieces:
        overlap = (box & rect).get_area()
        if not overlap or not words:
            continue
        if _like(words, reference):
            return False
        covered += overlap
        other = True
    if covered >= 0.5 * rect.get_area():
        return True
    if reference.letters < LOGO_LINES_SHARE:  # a logo of several lines of text
        if other:
            return True
        if ocr_lines is not None:  # OCR read the page, or the image the match is in: it would have read the logo
            if not page.get_text("words") or any(
                    (Rect(info["bbox"]) * page.rotation_matrix & rect).get_area() >= 0.8 * rect.get_area()
                    for info in page.get_image_info()):
                return True
    return False


def _letter_share(ink) -> float:
    """Height of the tallest connected shape over the height of the logo."""
    count, _, stats, _ = cv2.connectedComponentsWithStats((ink > 0.5).astype(uint8))
    tallest = max((stats[i, cv2.CC_STAT_HEIGHT] for i in range(1, count)), default=ink.shape[0])
    return float(min(1.0, tallest / ink.shape[0]))


def _ink(gray):
    """0 for paper, 1 for full ink; colour counts by its darkness."""
    return (255 - ascontiguousarray(gray, dtype=uint8)).astype(float32) / 255


def _crop_to_ink(ink):
    rows = where(ink.max(axis=1) > 0.25)[0]
    columns = where(ink.max(axis=0) > 0.25)[0]
    if not len(rows):
        raise ValueError("reference image has no logo in it: it is blank")
    return ink[rows[0]:rows[-1] + 1, columns[0]:columns[-1] + 1].copy()


def _rotated(ink, rotation: int):
    return {0: ink, 90: cv2.rotate(ink, cv2.ROTATE_90_COUNTERCLOCKWISE),
            180: cv2.rotate(ink, cv2.ROTATE_180), 270: cv2.rotate(ink, cv2.ROTATE_90_CLOCKWISE)}[rotation]


def _match(page, template, integrals=None):
    """Correlation of the template at every position; 0 where the page is (nearly) blank.

    `integrals`: cv2.integral2 of the page, when it is matched more than once.
    """
    if template.shape[0] > page.shape[0] or template.shape[1] > page.shape[1] or template.std() < 1e-3:
        return None
    result = nan_to_num(cv2.matchTemplate(page, template, cv2.TM_CCOEFF_NORMED), nan=0.0, posinf=0.0, neginf=0.0)
    # On blank paper the correlation divides by ~0 and comes out anything, 1.0 too
    total, squares = integrals if integrals is not None else cv2.integral2(page)
    height, width = template.shape
    count = height * width

    def window(table):
        return table[height:, width:] - table[:-height, width:] - table[height:, :-width] + table[:-height, :-width]

    mean = window(total) / count
    variance = window(squares) / count - mean * mean
    result[variance < LOGO_MIN_VARIANCE] = 0
    return result


def _peaks(result, threshold: float, size: tuple) -> list:
    """(score, x, y) of local maxima above threshold, at least half a template apart."""
    peaks = []
    result = result.copy()
    height, width = size
    while True:
        _, score, _, (x, y) = cv2.minMaxLoc(result)
        if score < threshold or len(peaks) >= 50:
            return peaks
        peaks.append((score, x, y))
        result[max(0, y - height // 2):y + height // 2 + 1, max(0, x - width // 2):x + width // 2 + 1] = 0


def _overlap(a, b) -> float:
    """Intersection over union: a box inside a much larger one (a letter of a logo) overlaps it little."""
    width = min(a[2], b[2]) - max(a[0], b[0])
    height = min(a[3], b[3]) - max(a[1], b[1])
    if width <= 0 or height <= 0:
        return 0.0
    intersection = width * height
    return intersection / ((a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - intersection)


def _best_first(candidates: list, contained: bool = False) -> list:
    """Candidates (score, box, ...) without those overlapping a better one;
    with `contained`, also without those lying mostly inside a better one (a
    verified logo's letters can match a smaller reference too)."""
    kept = []
    for candidate in sorted(candidates, key=lambda c: -c[0]):
        if all(_overlap(candidate[1], other[1]) < 0.3 and not (contained and _inside(candidate[1], other[1]) > 0.6)
               for other in kept):
            kept.append(candidate)
    return kept


def _inside(a, b) -> float:
    """Share of box a that lies in box b."""
    width = min(a[2], b[2]) - max(a[0], b[0])
    height = min(a[3], b[3]) - max(a[1], b[1])
    if width <= 0 or height <= 0:
        return 0.0
    return width * height / ((a[2] - a[0]) * (a[3] - a[1]))


def _template(reference, rotation: int, height: float, aspect: float):
    """The reference's ink at a logo height (unrotated pixels), its width stretched by aspect, then rotated."""
    scale = height / reference.ink.shape[0]
    return _rotated(cv2.resize(reference.ink, None, fx=scale * aspect, fy=scale, interpolation=cv2.INTER_AREA), rotation)


def find_logos(gray, references: list, rotations=LOGO_ROTATIONS, min_height: int = LOGO_MIN_HEIGHT,
               max_height: int = LOGO_MAX_HEIGHT, threshold: float = LOGO_CANDIDATE_SCORE) -> list:
    """Candidates [(score, (x0, y0, x1, y1), reference name, rotation, logo height, aspect)] in a gray image (uint8).

    Overlapping candidates are dropped only within a reference: a small
    reference can correlate better with a part of a logo than the right one
    with the whole of it, and verifying sorts that out.
    """
    ink = _ink(gray)
    candidates = {reference.name: [] for reference in references}
    height = min_height
    while height <= max_height:  # sizes outermost: one blurred page in memory at a time
        sigma = max(1.0, height / 12)
        page = cv2.GaussianBlur(ink, (0, 0), sigma)
        integrals = cv2.integral2(page)
        for reference in references:
            for rotation in rotations:
                for aspect in LOGO_ASPECTS:
                    template = _template(reference, rotation, height, aspect)
                    result = _match(page, cv2.GaussianBlur(template, (0, 0), sigma), integrals)
                    if result is None:
                        continue
                    th, tw = template.shape
                    for score, x, y in _peaks(result, threshold, (th, tw)):
                        candidates[reference.name].append((score, (x, y, x + tw, y + th), reference.name, rotation,
                                                           height, aspect))
        del page, integrals
        height *= LOGO_SCALE_STEP
    return [candidate for found in candidates.values() for candidate in _best_first(found)[:LOGO_MAX_CANDIDATES]]


def _verify(gray, reference, rotation: int, logo_height: float, aspect: float = 1.0) -> Optional[tuple]:
    """(score, box) of the best match in a sharp crop around a candidate, or None.

    Sizes 80-125% of the estimate and widths 88-113% of it are tried. A match
    in a photo or a dark area is none: a logo is clean ink on paper, with few
    mid-tones in it and mostly paper around it.
    """
    ink = _ink(gray)
    page = cv2.GaussianBlur(ink, (0, 0), LOGO_VERIFY_BLUR)
    integrals = cv2.integral2(page)
    best = None
    for stretch in (1.0, 0.88, 1.13):
        if stretch != 1.0 and (best is None or best[0] < LOGO_SCORE - 0.2):
            break  # nowhere near a logo at its own width: other widths will not make it one
        factor = 0.8
        while factor <= 1.25:
            template = _template(reference, rotation, logo_height * factor, aspect * stretch)
            factor *= 1.04
            result = _match(page, cv2.GaussianBlur(template, (0, 0), LOGO_VERIFY_BLUR), integrals)
            if result is None:
                continue
            _, score, _, (x, y) = cv2.minMaxLoc(result)
            if best is None or score > best[0]:
                best = (score, (x, y, x + template.shape[1], y + template.shape[0]), template)
    if best is None:
        return None
    score, (x0, y0, x1, y1), template = best
    inside = ink[y0:y1, x0:x1]
    found_ink = float((inside > 0.5).mean())
    reference_ink = float((template > 0.5).mean())
    if not 1 / LOGO_INK_RATIO <= (found_ink + 0.01) / (reference_ink + 0.01) <= LOGO_INK_RATIO:
        return None
    colour = float(percentile(inside, 85))  # the logo's ink: 1 for black, ~0.65 for a dark blue
    if float(((inside > 0.15) & (inside < colour - 0.15)).mean()) > LOGO_MAX_MIDTONES:
        return None
    count, _, stats, _ = cv2.connectedComponentsWithStats((inside > 0.5).astype(uint8))
    largest = max((stats[i, cv2.CC_STAT_AREA] for i in range(1, count)), default=0)
    if largest > LOGO_MAX_SHAPE_SHARE * inside.size:  # a filled box or bar, not letters
        return None
    pad = max(2, round(min(y1 - y0, x1 - x0) * 0.25))
    sides = (ink[max(0, y0 - pad):y0, x0:x1], ink[y1:y1 + pad, x0:x1],  # above, below, left, right
             ink[y0:y1, max(0, x0 - pad):x0], ink[y0:y1, x1:x1 + pad])
    if any(side.size and float(side.mean()) > LOGO_MAX_SURROUNDING_INK for side in sides):
        return None  # not on paper on every side: in a photo, or at its edge
    return score, (x0, y0, x1, y1)


def _gray_pixmap(pixmap):
    return frombuffer(pixmap.samples, dtype=uint8).reshape(pixmap.height, pixmap.stride)[:, :pixmap.width]


def find_pdf_logos(page, references: list, rotations=LOGO_ROTATIONS, score: float = LOGO_SCORE,
                   ocr_lines=None) -> list:
    """LogoMatch for every logo on a PDF page; boxes are Rects in page coordinates (as displayed).

    A match that the text layer -- or `ocr_lines`, `line_ocr.OcrLine`s in
    page coordinates -- reads as other words than its reference is skipped:
    two lines of bold text can look like a two-line logo.
    """
    from pymupdf import Matrix
    from pymupdf import Rect
    from pymupdf import csGRAY

    longest = max(page.rect.width, page.rect.height)
    zoom = min(LOGO_SEARCH_DPI / 72, LOGO_SEARCH_PIXELS / longest)
    coarse = _gray_pixmap(page.get_pixmap(matrix=Matrix(zoom, zoom), colorspace=csGRAY))
    by_name = {reference.name: reference for reference in references}
    matches = []
    candidates = find_logos(coarse, references, rotations, LOGO_MIN_POINTS * zoom, LOGO_MAX_POINTS * zoom)
    if not candidates:
        return []
    # One sharp render to crop the candidates from: rendering clips one by one
    # would decode a scanned page's image again for every candidate
    sharp_zoom = min(LOGO_VERIFY_HEIGHT / LOGO_MIN_POINTS, LOGO_SHARP_PIXELS / longest)
    sharp_page = _gray_pixmap(page.get_pixmap(matrix=Matrix(sharp_zoom, sharp_zoom), colorspace=csGRAY))
    for _, box, name, rotation, logo_height, aspect in candidates:
        # the box and half a logo height around it, scaled so that the logo is LOGO_VERIFY_HEIGHT high
        pad = logo_height / 2
        clip = Rect((box[0] - pad) / zoom, (box[1] - pad) / zoom, (box[2] + pad) / zoom, (box[3] + pad) / zoom)
        clip &= page.rect
        x0, y0 = int(clip.x0 * sharp_zoom), int(clip.y0 * sharp_zoom)
        x1, y1 = int(clip.x1 * sharp_zoom) + 1, int(clip.y1 * sharp_zoom) + 1
        z = LOGO_VERIFY_HEIGHT / (logo_height / zoom)  # crop pixels per point
        factor = z / sharp_zoom
        crop = cv2.resize(sharp_page[y0:y1, x0:x1], None, fx=factor, fy=factor,
                          interpolation=cv2.INTER_AREA if factor < 1 else cv2.INTER_CUBIC)
        verified = _verify(crop, by_name[name], rotation, LOGO_VERIFY_HEIGHT, aspect)
        if verified is None or verified[0] < score:
            continue
        found, (a, b, c, d) = verified
        rect = Rect(x0 / sharp_zoom + a / z, y0 / sharp_zoom + b / z, x0 / sharp_zoom + c / z, y0 / sharp_zoom + d / z)
        if min(rect.width, rect.height) * by_name[name].letters < LOGO_MIN_POINTS:  # a line of small text
            continue
        if _reads_otherwise(page, rect, by_name[name], ocr_lines):
            continue
        matches.append(LogoMatch(page.number, rect, round(float(found), 3), name, rotation))
    kept = _best_first([(m.score, tuple(m.box), m) for m in matches], contained=True)
    return [entry[2] for entry in kept]


def find_image_logos(image, references: list, rotations=LOGO_ROTATIONS, score: float = LOGO_SCORE) -> list:
    """LogoMatch for every logo in a PIL image; boxes are pixel boxes (x0, y0, x1, y1)."""
    from numpy import asarray

    gray = asarray(image.convert("L"))
    zoom = min(1.0, LOGO_SEARCH_PIXELS / max(gray.shape))
    coarse = cv2.resize(gray, None, fx=zoom, fy=zoom, interpolation=cv2.INTER_AREA) if zoom < 1 else gray
    by_name = {reference.name: reference for reference in references}
    matches = []
    for _, box, name, rotation, logo_height, aspect in find_logos(coarse, references, rotations):
        pad = logo_height / 2
        x0, y0 = max(0, int((box[0] - pad) / zoom)), max(0, int((box[1] - pad) / zoom))
        x1, y1 = min(gray.shape[1], int((box[2] + pad) / zoom) + 1), min(gray.shape[0], int((box[3] + pad) / zoom) + 1)
        sharp = LOGO_VERIFY_HEIGHT / (logo_height / zoom)
        crop = cv2.resize(gray[y0:y1, x0:x1], None, fx=sharp, fy=sharp,
                          interpolation=cv2.INTER_AREA if sharp < 1 else cv2.INTER_CUBIC)
        verified = _verify(crop, by_name[name], rotation, LOGO_VERIFY_HEIGHT, aspect)
        if verified is None or verified[0] < score:
            continue
        found, (a, b, c, d) = verified
        pixel_box = (x0 + a / sharp, y0 + b / sharp, x0 + c / sharp, y0 + d / sharp)
        matches.append(LogoMatch(0, pixel_box, round(float(found), 3), name, rotation))
    kept = _best_first([(m.score, tuple(m.box), m) for m in matches], contained=True)
    return [entry[2] for entry in kept]


def _grown(box, rotation: int) -> tuple:
    x0, y0, x1, y1 = box
    margin = LOGO_MARGIN * ((x1 - x0) if rotation in (90, 270) else (y1 - y0))
    return x0 - margin, y0 - margin, x1 + margin, y1 + margin


def obfuscate_logos(filepath: str, references: list, output_path: str, fill=LOGO_FILL,
                    rotations=LOGO_ROTATIONS, score: float = LOGO_SCORE, find_in: Optional[str] = None,
                    ocr: bool = False) -> tuple:
    """Write a copy of a PDF or image file with every logo like the references covered.

    Args:
        find_in (str, optional): the file to look for the logos in, when
            filepath is a copy of it that was changed in place -- e.g. by
            `obfuscate_document`, whose black boxes and placeholders could
            hide a logo or pass for one. Same pages, same sizes.
        ocr (bool, optional): read what a PDF shows only as pixels with OCR,
            to tell text that looks like a logo from the logo. Defaults to False.

    Returns:
        tuple: output_path, and the LogoMatch list (boxes before the margin).
    """
    filepath, output_path = fspath(filepath), fspath(output_path)
    find_in = fspath(find_in) if find_in is not None else filepath
    if splitext(filepath)[1].lower() == ".pdf":
        return output_path, _obfuscate_pdf_logos(filepath, output_path, references, fill, rotations, score, find_in,
                                                 ocr)
    from PIL import Image
    from PIL import ImageDraw

    with Image.open(find_in) as original:
        matches = find_image_logos(original, references, rotations, score)
    with Image.open(filepath) as image:
        image.load()
        if image.mode not in ("RGB", "L", "RGBA"):
            image = image.convert("RGB")
        draw = ImageDraw.Draw(image)
        color = fill if image.mode != "L" else round(sum(fill) / 3)
        for match in matches:
            draw.rectangle(_grown(match.box, match.rotation), fill=color)
        image.save(output_path)
    return output_path, matches


def _obfuscate_pdf_logos(filepath, output_path, references, fill, rotations, score, find_in, ocr=False) -> list:
    from pymupdf import PDF_REDACT_IMAGE_PIXELS
    from pymupdf import PDF_REDACT_LINE_ART_REMOVE_IF_COVERED
    from pymupdf import PDF_REDACT_TEXT_REMOVE
    from pymupdf import Rect
    from pymupdf import open as pdf_open

    found = []
    with pdf_open(filepath) as doc, pdf_open(find_in) as original:
        if original.page_count != doc.page_count:
            raise ValueError(f"{find_in} has {original.page_count} pages, {filepath} {doc.page_count}")
        for page in doc:
            lines = None
            if ocr:
                from .obfuscation import _pdf_ocr_lines

                source = original[page.number]
                lines = _pdf_ocr_lines(original, page.number, bool(source.get_text().strip()), False)
            matches = find_pdf_logos(original[page.number], references, rotations, score, lines)
            for match in matches:
                rect = Rect(_grown(tuple(match.box), match.rotation)) * page.derotation_matrix
                page.add_redact_annot(rect, fill=fill)
            if matches:
                page.apply_redactions(images=PDF_REDACT_IMAGE_PIXELS, graphics=PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                                      text=PDF_REDACT_TEXT_REMOVE)
            found += matches
        doc.save(output_path, garbage=3, deflate=True)
    return found
