"""OCR lines whose characters can be located (`acd.line_ocr`).

Needs RapidOCR (`pip install rapidocr onnxruntime`); skipped without it. The
images are text rendered by PyMuPDF, so where every character really is, is
known.
"""

import numpy as np
import pymupdf
import pytest

pytest.importorskip("rapidocr")

from acd.line_ocr import ocr_lines  # noqa: E402

ZOOM = 200 / 72


def render(page):
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM), colorspace=pymupdf.csGRAY)
    return np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width)


def glyph_centres(page):
    centres = {}
    for block in page.get_text("rawdict")["blocks"]:
        for line in block["lines"]:
            for span in line["spans"]:
                for char in span["chars"]:
                    box = pymupdf.Rect(char["bbox"]) * ZOOM
                    centres.setdefault(char["c"], []).append(((box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2))
    return centres


def test_characters_are_located_in_every_orientation():
    # RapidOCR's own character boxes are mirrored on lines it turns over to
    # read: upside down, and running down the page
    doc = pymupdf.open()
    page = doc.new_page(width=400, height=400)
    for angle, point in {0: (60, 80), 90: (60, 360), 180: (340, 330), 270: (340, 60)}.items():
        page.insert_text(point, "AB-1234", fontsize=20, rotate=angle)
    centres = glyph_centres(page)

    lines = ocr_lines(render(page))

    assert sorted(line.text for line in lines) == ["AB-1234"] * 4
    for line in lines:
        for index, char in enumerate(line.text):
            corners = line.quad(index, index + 1, pad=0)
            x, y = corners[:, 0].mean(), corners[:, 1].mean()
            # within half a character (some 30 pixels wide here); mirrored
            # positions were off by up to six characters
            assert min(np.hypot(x - cx, y - cy) for cx, cy in centres[char]) < 15, (line, char)


@pytest.mark.parametrize("angle, along, up", [
    (0, (1, 0), (0, -1)), (90, (0, -1), (-1, 0)), (180, (-1, 0), (0, 1)), (270, (0, 1), (1, 0)),
])
def test_corners_follow_the_upright_text(angle, along, up):
    doc = pymupdf.open()
    page = doc.new_page(width=300, height=300)
    page.insert_text((150, 150), "D9893", fontsize=20, rotate=angle)

    (line,) = ocr_lines(render(page))

    top_left, top_right, _, bottom_left = line.quad(0, len(line.text), pad=0)
    assert np.sign(np.round(top_right - top_left, -1)).tolist() == list(along)
    assert np.sign(np.round(top_left - bottom_left, -1)).tolist() == list(up)


def test_lines_cut_by_detection_tiles_come_back_whole():
    # tiles are 2048 pixels, 1664 apart: this line starts before the second
    # one and ends after the first, so neither sees all of it
    doc = pymupdf.open()
    page = doc.new_page(width=1400, height=120)
    page.insert_text((500, 60), "LIEBHERR-AEROSPACE LINDENBERG GmbH", fontsize=24)
    image = render(page)
    assert 500 * ZOOM < 1664 and 2048 < 950 * ZOOM < image.shape[1]

    texts = ["".join(line.text.split()) for line in ocr_lines(image)]

    assert texts == ["LIEBHERR-AEROSPACELINDENBERGGmbH"]


def spans(*gaps, start=10.0):
    """Positions of characters whose centres lie the given gaps apart."""
    centres = [start]
    for gap in gaps:
        centres.append(centres[-1] + gap)
    return [(centre - 5, centre + 5) for centre in centres]


def test_numbers_read_as_one_run_are_split_where_they_stand_apart():
    from acd.line_ocr import _split_numbers

    # two balloons, (0220) and (1850), read as "(02201850)": gaps measured on a scanned drawing
    positions = [(0, 10)] + spans(41, 41, 41, 55, 41, 27, 41, start=20) + [(330, 340)]
    text, split = _split_numbers("(02201850)", positions)
    assert text == "(0220 1850)"
    assert split[5] is None and split[6] == positions[5] and len(split) == len(positions) + 1
    # a row of balloons
    row = spans(34, 34, 51, 67, 34, 34, 51, 67, 34, 34, 51, 67, 34, 34, 51)
    assert _split_numbers("1640169017301610", row)[0] == "1640 1690 1730 1610"
    # one number: the narrow 1 makes its gaps vary, but not into numbers of 3 digits and more
    assert _split_numbers("11711791", spans(15, 15, 23, 15, 22, 15, 22))[0] == "11711791"
    assert _split_numbers("085170749", spans(29, 22, 22, 29, 22, 29, 29, 22))[0] == "085170749"
    # digits in a code stay together, gaps or not
    code = [(0, 10), (10, 20), (20, 30)] + spans(20, 20, 40, 20, 20, 20, start=40)
    assert _split_numbers("E2-2019390", code)[0] == "E2-2019390"


def test_a_box_that_reads_as_nothing_is_detected_again_on_its_own(monkeypatch):
    # on the whole page, the detector takes a label and the signature across it for one box
    from acd import line_ocr

    whole = np.array([[100, 200], [500, 200], [500, 260], [100, 260]], dtype=np.float64)
    seen = []

    def detect(engine, gray):
        seen.append(gray.shape)
        if len(seen) == 1:
            return [(whole, False)]
        # in the part: the label, and the same box again, which must not be read twice
        return [(np.array([[38, 32], [238, 32], [238, 52], [38, 52]], dtype=np.float64), False),
                (whole - (70, 170), False)]

    def read(engine, gray, corners, vertical):
        width = corners[:, 0].max() - corners[:, 0].min()
        if width > 300:
            return None
        return line_ocr.OcrLine("SIGNATURES & DATE", 0.99, [(0, 1)] * 17, int(width), 20, False, np.eye(3))

    monkeypatch.setattr(line_ocr, "_rapidocr", lambda: None)
    monkeypatch.setattr(line_ocr, "_detect", detect)
    monkeypatch.setattr(line_ocr, "_read", read)
    lines = line_ocr.ocr_lines(np.full((1000, 1000), 255, dtype=np.uint8))
    assert [line.text for line in lines] == ["SIGNATURES & DATE"]
    assert seen[1] == (120, 460)  # the box, padded by half its height
