"""OCR of an image into lines of text whose characters can be located.

`ocr_lines(image)` finds and reads the text lines of an image with RapidOCR
and its bundled PP-OCRv6 models (`pip install rapidocr onnxruntime`, no
download at run time), and returns each as an `OcrLine` that tells where any
run of its characters is -- what blacking out a word of a scan needs:

    for line in ocr_lines(gray):
        start = line.text.find("D9893")
        if start >= 0:
            corners = line.quad(start, start + 5)  # image pixels

RapidOCR is only used to find the lines and to read them: its own character
boxes come out mirrored on every line it turns over to read -- upside-down
text, and text running down the page, as dimensions on drawings often do. So
each line is straightened and turned the right way up here, and its
characters are placed through our own transform. Large images are scanned for
lines tile by tile, which keeps memory flat (a whole A0 sheet at 200 dpi in
one go takes 8 GB).
"""

from collections import namedtuple

import numpy as np

try:
    from cv2 import COLOR_RGB2GRAY
    from cv2 import INTER_LINEAR
    from cv2 import cvtColor
    from cv2 import getPerspectiveTransform
    from cv2 import perspectiveTransform
    from cv2 import warpPerspective
except ImportError as error:  # RapidOCR brings OpenCV along
    raise ImportError("OCR needs RapidOCR: pip install rapidocr onnxruntime") from error


# Lines are looked for in tiles this big, overlapping by this much; a line
# that tiles cut apart is joined up again.
OCR_TILE = 2048
OCR_TILE_OVERLAP = 384

# Lines read with less confidence are dropped.
OCR_MIN_SCORE = 0.3

# Changes when `ocr_lines` reads the same image differently: caches of its
# results key on it (2: boxes that read as nothing are detected again, runs
# of digits split where they stand apart).
OCR_VERSION = 2

# The recogniser reads numbers that stand apart -- item numbers in the
# neighbouring balloons of a drawing -- as one run of digits ("0220 1850" as
# "02201850"). The gaps tell them apart: one size of gap is this much wider
# than the next smaller, and splitting at it leaves numbers of this many
# digits at least (the gaps in one number vary too, around a narrow "1",
# but do not cut it into such pieces).
_NUMBER_GAP = 1.3
_MIN_NUMBER_DIGITS = 3
# Only runs standing on their own are split, as balloons do: between these
# ("" is the end of the line).
_NUMBER_EDGES = ("", " ", "(", ")", "[", "]")

# A horizontal line read with at least this confidence is taken to be upright;
# below it, the line is read upside down too and the better reading kept.
_UPRIGHT_SCORE = 0.85

# The recogniser places a character a fifth of a character late along its
# line, on average, measured against where rendered characters really were.
# Corrected, 90% of them -- in all four directions -- are placed within 0.35
# of a character: what `OcrLine.quad` widens a span by.
_LAG = 0.2
_PAD = 0.35

# Every call names every step: RapidOCR keeps the ones of the call before.
_DETECT = dict(use_det=True, use_cls=False, use_rec=False, return_word_box=False, return_single_char_box=False)
_READ = dict(use_det=False, use_cls=False, use_rec=True, return_word_box=True, return_single_char_box=False)

_engines = {}


class OcrLine(namedtuple("OcrLine", "text score positions width height turned to_image")):
    """A line of text read from an image.

    The line was read from a straightened strip `width` x `height` pixels big
    -- upside down if `turned` -- and `positions` holds, for each character of
    `text`, its (start, end) along that strip, None for whitespace. The 3x3
    perspective matrix `to_image` maps strip coordinates into the image.
    """

    def quad(self, start: int, end: int, pad: float = _PAD) -> np.ndarray:
        """The corners, in image coordinates, of `text[start:end]` widened on
        either side by `pad` of a character's width: top-left, top-right,
        bottom-right, bottom-left of the text the right way up."""
        spans = [span for span in self.positions[start:end] if span is not None]
        x0, x1 = min(span[0] for span in spans), max(span[1] for span in spans)
        widen = pad * (x1 - x0) / len(spans)
        x0, x1 = max(x0 - widen, 0), min(x1 + widen, self.width)
        if self.turned:
            corners = [[x1, self.height], [x0, self.height], [x0, 0], [x1, 0]]
        else:
            corners = [[x0, 0], [x1, 0], [x1, self.height], [x0, self.height]]
        return perspectiveTransform(np.array([corners], dtype=np.float64), self.to_image)[0]


def ocr_lines(image: np.ndarray) -> list:
    """Find and read the text lines of an image: a 2-D grayscale or an RGB
    array. Returns a list of `OcrLine`."""
    gray = image if image.ndim == 2 else cvtColor(image, COLOR_RGB2GRAY)
    engine = _rapidocr()
    lines = []
    for corners, vertical in _detect(engine, gray):
        line = _read(engine, gray, corners, vertical)
        if line is not None and line.score >= OCR_MIN_SCORE:
            lines.append(line)
        else:
            lines += _read_again(engine, gray, corners)
    return lines


def _read_again(engine, gray: np.ndarray, corners: np.ndarray) -> list:
    """The lines in a box that read as nothing, found again in that part of
    the image alone. On a whole page the detector can take a printed label,
    the signature written across it and the text beside it for one box
    ("SUPPLIERS SIGNATURES & DATE" on a scanned form); on its own, the part
    splits into lines that read."""
    height, width = gray.shape
    x0, y0 = corners.min(axis=0)
    x1, y1 = corners.max(axis=0)
    pad = max(8.0, min(x1 - x0, y1 - y0) / 2)
    left, top = int(max(x0 - pad, 0)), int(max(y0 - pad, 0))
    right, bottom = int(min(x1 + pad, width)), int(min(y1 + pad, height))
    if right - left > OCR_TILE or bottom - top > OCR_TILE:
        return []  # a tile's worth: the first detection already saw it this way
    part = gray[top:bottom, left:right]
    area = (x1 - x0) * (y1 - y0)
    lines = []
    for sub_corners, vertical in _detect(engine, part):
        sub_corners = sub_corners + (left, top)
        (a0, b0), (a1, b1) = sub_corners.min(axis=0), sub_corners.max(axis=0)
        overlap = max(min(a1, x1) - max(a0, x0), 0) * max(min(b1, y1) - max(b0, y0), 0)
        if overlap >= 0.8 * max(area, (a1 - a0) * (b1 - b0)):
            continue  # the same box again: it would read as nothing again
        line = _read(engine, gray, sub_corners, vertical)
        if line is not None and line.score >= OCR_MIN_SCORE:
            lines.append(line)
    return lines


def _rapidocr():
    if "rapidocr" not in _engines:
        try:
            from rapidocr import RapidOCR
        except ImportError as error:
            raise ImportError("OCR needs RapidOCR: pip install rapidocr onnxruntime") from error
        _engines["rapidocr"] = RapidOCR(params={"Global.log_level": "error", "Global.use_preprocess_img": False})
    return _engines["rapidocr"]


def _as_rgb(gray: np.ndarray) -> np.ndarray:
    return np.ascontiguousarray(np.dstack([gray] * 3))


def _detect(engine, gray: np.ndarray) -> list:
    """[(corners, vertical)] of the text lines of an image, found tile by tile:
    corners as a 4x2 array (top-left, top-right, bottom-right, bottom-left)."""
    height, width = gray.shape
    boxes = []  # x0, y0, x1, y1, tile
    for tile, (y, x) in enumerate((y, x) for y in _tile_starts(height) for x in _tile_starts(width)):
        found = engine(_as_rgb(gray[y:y + OCR_TILE, x:x + OCR_TILE]), **_DETECT)
        for quad in found.boxes if found.boxes is not None else ():
            quad = np.asarray(quad, dtype=np.float64) + (x, y)
            boxes.append((*quad.min(axis=0), *quad.max(axis=0), tile))
    return [
        (np.array([[x0, y0], [x1, y0], [x1, y1], [x0, y1]]), vertical)
        for x0, y0, x1, y1, vertical in _widen(_join_cut_lines(boxes), width, height)
    ]


def _widen(boxes: list, width: int, height: int) -> list:
    """The detector's box can end inside a line's first or last character:
    widen each along its line by a quarter of its thickness -- into empty
    space only, not into the next box on the line, whose characters it would
    read too. Returns (x0, y0, x1, y1, vertical)."""
    array = np.array(boxes, dtype=np.float64).reshape(-1, 4)
    widened = []
    for index, (x0, y0, x1, y1) in enumerate(boxes):
        vertical = y1 - y0 >= 1.5 * (x1 - x0)
        # along the line: x for a horizontal one, y for a vertical one
        start, end, low, high = (y0, y1, x0, x1) if vertical else (x0, x1, y0, y1)
        starts, ends, lows, highs = array.T[[1, 3, 0, 2] if vertical else [0, 2, 1, 3]]
        same_line = np.minimum(highs, high) - np.maximum(lows, low) >= 0.5 * (high - low)
        same_line[index] = False
        along, limit = 0.25 * (high - low), height if vertical else width
        before = float(max(start - along, ends[same_line & (ends <= start)].max(initial=-np.inf), 0))
        after = float(min(end + along, starts[same_line & (starts >= end)].min(initial=np.inf), limit))
        widened.append((x0, before, x1, after, True) if vertical else (before, y0, after, y1, False))
    return widened


def _tile_starts(size: int) -> list:
    """Where the tiles along one side start: OCR_TILE apart less the overlap,
    the last one moved back to end at the edge. Every tile is full size, as
    the detector splits a line of text running down the page into pieces in
    some smaller ones."""
    last = max(size - OCR_TILE, 0)
    return sorted(set(range(0, last, OCR_TILE - OCR_TILE_OVERLAP)) | {last})


def _join_cut_lines(boxes: list) -> list:
    """Join the pieces of a line: of overlapping boxes, one mostly inside the
    other (a copy two tiles both saw, or a piece the edge of a tile cut off),
    or two running on along the same line -- side by side, or one above the
    other in a line running down the page. The detector splits a line where
    the size of its text changes, and lines running down the page in some
    tile sizes; within a tile, pieces must overlap along the line by a
    seventh of the shorter, which neighbouring cells and lines do not.
    Returns the joined (x0, y0, x1, y1)."""
    parent = list(range(len(boxes)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    array = np.array([box[:4] for box in boxes], dtype=np.float64).reshape(-1, 4)
    for i, a in enumerate(boxes):
        rest = array[i + 1:]
        overlapping = (np.minimum(rest[:, 2], a[2]) > np.maximum(rest[:, 0], a[0])) & (
            np.minimum(rest[:, 3], a[3]) > np.maximum(rest[:, 1], a[1])
        )
        for j in (np.flatnonzero(overlapping) + i + 1).tolist():
            b = boxes[j]
            across_x = min(a[2], b[2]) - max(a[0], b[0])
            across_y = min(a[3], b[3]) - max(a[1], b[1])
            same_tile = a[4] == b[4]
            width = max(a[2], b[2]) - min(a[0], b[0])
            height = max(a[3], b[3]) - min(a[1], b[1])
            smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
            side_by_side = across_y >= 0.6 * min(a[3] - a[1], b[3] - b[1]) and width >= height
            stacked = across_x >= 0.6 * min(a[2] - a[0], b[2] - b[0]) and height >= 1.5 * width
            if same_tile:  # the pieces must overlap along the line
                side_by_side = side_by_side and across_x >= min(a[2] - a[0], b[2] - b[0]) / 7
                stacked = stacked and across_y >= min(a[3] - a[1], b[3] - b[1]) / 7
            if (not same_tile and across_x * across_y >= 0.7 * smaller) or side_by_side or stacked:
                parent[root(i)] = root(j)
    joined = {}
    for i, box in enumerate(boxes):
        group = joined.setdefault(root(i), list(box[:4]))
        group[:] = min(group[0], box[0]), min(group[1], box[1]), max(group[2], box[2]), max(group[3], box[3])
    return list(joined.values())


def _read(engine, gray: np.ndarray, corners: np.ndarray, vertical: bool):
    """Read one line: straighten it into a strip, read it the right way up --
    a vertical line both ways, as text runs up and down the page alike -- and
    place its characters along the strip."""
    top_left, top_right, bottom_right, bottom_left = corners
    if vertical:  # turned to read bottom to top; the other way is the strip upside down
        source = (bottom_left, top_left, top_right, bottom_right)
        width, height = top_left[1] - bottom_left[1], top_right[0] - top_left[0]
    else:
        source = (top_left, top_right, bottom_right, bottom_left)
        width, height = top_right[0] - top_left[0], bottom_left[1] - top_left[1]
    width, height = max(int(round(abs(width))), 2), max(int(round(abs(height))), 2)
    target = np.array([[0, 0], [width, 0], [width, height], [0, height]], dtype=np.float32)
    to_strip = getPerspectiveTransform(np.array(source, dtype=np.float32), target)
    strip = warpPerspective(gray, to_strip, (width, height), flags=INTER_LINEAR, borderValue=255)
    best = None
    for turned in (False, True):
        read = engine(_as_rgb(strip[::-1, ::-1] if turned else strip), **_READ)
        if read.txts and (best is None or read.scores[0] > best[1]):
            best = (read.txts[0], float(read.scores[0]), read.word_results[0] if read.word_results else None, turned)
        if best and not vertical and best[1] >= _UPRIGHT_SCORE:
            break
    if best is None or not best[0].strip():
        return None
    text, score, info, turned = best
    text, positions = _split_numbers(text, _positions(text, info, width))
    if turned:  # along the turned strip -> along the strip
        positions = [None if span is None else (width - span[1], width - span[0]) for span in positions]
    return OcrLine(text, score, positions, width, height, turned, np.linalg.inv(to_strip))


def _split_numbers(text: str, positions: list) -> tuple:
    """`text` and `positions` with a space wherever a run of digits holds
    several numbers, told by the gaps between their characters' centres."""
    runs = []
    start = None
    for index, char in enumerate(text + " "):
        if "0" <= char <= "9":
            start = index if start is None else start
        elif start is not None:
            runs.append((start, index))
            start = None
    cuts = set()
    runs = [(start, end) for start, end in runs  # not part of a code: E2-2019390, S4-3000520W01
            if text[start - 1:start] in _NUMBER_EDGES and text[end:end + 1] in _NUMBER_EDGES]
    while runs:
        start, end = runs.pop()
        if end - start < 2 * _MIN_NUMBER_DIGITS:
            continue
        centres = [(positions[index][0] + positions[index][1]) / 2 for index in range(start, end)]
        gaps = [right - left for left, right in zip(centres, centres[1:])]
        sizes = sorted(set(gaps), reverse=True)
        for wider, narrower in zip(sizes, sizes[1:]):
            if narrower <= 0 or wider < _NUMBER_GAP * narrower:
                continue
            bounds = [0] + [index + 1 for index, gap in enumerate(gaps) if gap >= wider] + [end - start]
            if all(right - left >= _MIN_NUMBER_DIGITS for left, right in zip(bounds, bounds[1:])):
                cuts.update(start + bound for bound in bounds[1:-1])
                runs += [(start + left, start + right) for left, right in zip(bounds, bounds[1:])]
                break
    if not cuts:
        return text, positions
    new_text, new_positions = [], []
    for index, (char, span) in enumerate(zip(text, positions)):
        if index in cuts:
            new_text.append(" ")
            new_positions.append(None)
        new_text.append(char)
        new_positions.append(span)
    return "".join(new_text), new_positions


def _positions(text: str, info, width: int) -> list:
    """(start, end) along the strip of each character of `text`, None for
    whitespace: from the recogniser's column for each character where it
    gives them, else spread evenly."""
    letters = [index for index, char in enumerate(text) if not char.isspace()]
    columns = [column for word in getattr(info, "word_cols", None) or () for column in word]
    measured = len(columns) == len(letters) and getattr(info, "line_txt_len", 0)
    if measured:
        centres = [(column + 0.5) * width / info.line_txt_len for column in columns]
    else:
        centres = [(rank + 0.5) * width / max(len(letters), 1) for rank in range(len(letters))]
    pitch = (centres[-1] - centres[0]) / (len(centres) - 1) if len(centres) > 1 else width / max(len(letters), 1)
    if measured:
        centres = [centre - _LAG * pitch for centre in centres]
    positions = [None] * len(text)
    for index, centre in zip(letters, centres):
        positions[index] = (centre - pitch / 2, centre + pitch / 2)
    return positions
