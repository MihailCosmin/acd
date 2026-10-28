"""Tests for acd.stamp_obfuscation: stamp annotations, stamp images and framed stamps on scans."""

import numpy as np
import pytest

pytest.importorskip("cv2")
pymupdf = pytest.importorskip("pymupdf")
from PIL import Image
from PIL import ImageDraw

from acd.line_ocr import OcrLine
from acd.stamp_obfuscation import find_stamp_areas
from acd.stamp_obfuscation import obfuscate_stamps


def read(text, rect):
    """An OCR line as `_pdf_ocr_lines` gives it: in page coordinates."""
    to_page = np.array([[1, 0, rect[0]], [0, 1, rect[1]], [0, 0, 1]], dtype=np.float64)
    return OcrLine(text, 0.99, [(0, 1)] * len(text), rect[2] - rect[0], rect[3] - rect[1], False, to_page)


def test_stamp_annotations_go(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((50, 50), "Released for production")
    page.add_stamp_annot(pymupdf.Rect(300, 100, 500, 160), stamp=pymupdf.STAMP_Approved)
    doc.save(tmp_path / "stamped.pdf")
    _, found = obfuscate_stamps(tmp_path / "stamped.pdf", tmp_path / "out.pdf", ocr=False)
    [(number, area)] = found
    assert number == 0 and area in pymupdf.Rect(300, 100, 500, 160) and area.width > 190  # its aspect kept
    with pymupdf.open(tmp_path / "out.pdf") as out:
        page = out[0]
        assert not list(page.annots())
        assert "Released for production" in page.get_text()  # the page's text is no stamp
        assert page.get_pixmap(clip=pymupdf.Rect(380, 120, 420, 140), dpi=72).pixel(10, 10) == (0, 0, 0)


def test_an_image_that_reads_like_a_stamp_goes_whole(tmp_path):
    image = Image.new("RGB", (400, 80), "white")
    ImageDraw.Draw(image).rectangle([2, 2, 397, 77], outline=(40, 70, 140), width=4)
    image.save(tmp_path / "stamp.png")
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(pymupdf.Rect(400, 500, 500, 520), filename=str(tmp_path / "stamp.png"))
    lines = [read("Accepted by Customer", (405, 502, 495, 512)),
             read("Approved", (60, 300, 120, 310)),  # outside any image and frame: a word of the page
             read("This change was approved by the board in its last meeting", (405, 512, 495, 518))]
    areas = find_stamp_areas(page, lines)
    assert len(areas) == 1 and pymupdf.Rect(400, 500, 500, 520) in areas[0]


def scan(tmp_path, draw_on):
    """A page that is one image, drawn on at 2 pixels per point."""
    image = Image.new("L", (595 * 2, 842 * 2), 255)
    draw_on(ImageDraw.Draw(image))
    image.save(tmp_path / "scan.png")
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, filename=str(tmp_path / "scan.png"))
    return doc, page


def test_a_framed_stamp_on_a_scan_goes_and_a_table_header_stays(tmp_path):
    def draw_on(draw):
        # an approval table: "genehmigt / APPROVED" heads a column, its cells' borders run on
        for x in (60, 160, 260, 360):
            draw.line([(x * 2, 100 * 2), (x * 2, 200 * 2)], fill=0, width=3)
        for y in (100, 130, 160, 200):
            draw.line([(60 * 2, y * 2), (360 * 2, y * 2)], fill=0, width=3)
        # a rubber stamp in a frame of its own, below the table
        draw.rectangle([(300 * 2, 400 * 2), (420 * 2, 440 * 2)], outline=0, width=3)

    doc, page = scan(tmp_path, draw_on)
    lines = [read("APPROVED", (170, 110, 250, 120)), read("RECEIVED", (320, 410, 400, 420))]
    areas = find_stamp_areas(page, lines)
    assert len(areas) == 1
    assert pymupdf.Rect(301, 401, 419, 439) in areas[0] and areas[0] in pymupdf.Rect(296, 396, 424, 444)
