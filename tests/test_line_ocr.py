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
