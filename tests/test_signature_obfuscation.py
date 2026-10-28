"""Tests for acd.signature_obfuscation: signature cells found by their labels, and blacked out."""

import pytest

pytest.importorskip("cv2")
pymupdf = pytest.importorskip("pymupdf")
from PIL import Image
from PIL import ImageDraw

from acd.signature_obfuscation import find_signature_areas
from acd.signature_obfuscation import obfuscate_signatures


def scribble(path, colour):
    image = Image.new("RGB", (300, 80), "white")  # white background: it hides the line it is pasted over
    draw = ImageDraw.Draw(image)
    draw.line([(10, 60), (60, 10), (90, 70), (140, 20), (200, 65), (280, 30)], fill=colour, width=5)
    image.save(path)
    return str(path)


def make_record(tmp_path):
    """A validation record: a table NAME | TYPE | SIGNATURE, two signed rows, and a form cell 'Signed by:'."""
    doc = pymupdf.open()
    page = doc.new_page()
    shape = page.new_shape()
    columns, rows = (60, 200, 320, 520), (100, 120, 160, 200)
    for x in columns:
        shape.draw_line((x, rows[0]), (x, rows[-1]))
    for y in rows:
        shape.draw_line((columns[0], y), (columns[-1], y))
    shape.draw_rect(pymupdf.Rect(60, 260, 520, 330))  # the form cell
    shape.finish(color=(0, 0, 0), width=0.8)
    # pen strokes in the form cell
    shape.draw_bezier((200, 310), (230, 270), (260, 340), (300, 290))
    shape.finish(color=(0, 0, 0.6), width=1.5)
    shape.commit()
    page.insert_text((65, 114), "NAME", fontsize=9)
    page.insert_text((205, 114), "TYPE", fontsize=9)
    page.insert_text((325, 114), "SIGNATURE", fontsize=9)
    for row, (name, kind) in enumerate([("Heike Wochner", "Approval"), ("J. Miller", "Check")]):
        y = rows[row + 1]
        page.insert_text((65, y + 24), name, fontsize=9)
        page.insert_text((205, y + 24), kind, fontsize=9)
        # the scribble runs over the cell's lines, as signatures do
        page.insert_image(pymupdf.Rect(330, y + 2, 515, y + 46),
                          filename=scribble(tmp_path / f"sig{row}.png", (20, 40, 160)))
    page.insert_text((65, 275), "Signed by:", fontsize=9)
    path = tmp_path / "record.pdf"
    doc.save(path)
    return path


def test_signature_column_and_form_cell_are_found(tmp_path):
    with pymupdf.open(make_record(tmp_path)) as doc:
        areas = find_signature_areas(doc[0])
    column = [area for area in areas if area.x0 >= 320 and area.y1 <= 205]
    assert len(column) >= 2 and all(area.y0 >= 118 for area in column)  # the cells below the header, not it
    assert any(area.intersects(pymupdf.Rect(200, 280, 300, 320)) for area in areas)  # the pen strokes
    assert not any(area.intersects(pymupdf.Rect(205, 140, 240, 150)) for area in areas)  # "Approval" is a value


def test_signatures_are_redacted_and_the_rest_stays(tmp_path):
    source = make_record(tmp_path)
    output, found = obfuscate_signatures(source, tmp_path / "out.pdf")
    assert found
    with pymupdf.open(output) as doc:
        page = doc[0]
        text = page.get_text()
        assert "SIGNATURE" in text and "Signed by:" in text and "Heike Wochner" in text and "Approval" in text
        # the scribbles' pixels are gone: what is left of the image is blank or black
        pixmap = page.get_pixmap(clip=pymupdf.Rect(340, 125, 500, 155), dpi=72)
        assert not any(b > 120 and r < 100 for r, g, b in
                       (pixmap.pixel(x, y) for x in range(pixmap.width) for y in range(pixmap.height)))
        assert not [path for path in page.get_drawings() if any(item[0] == "c" for item in path["items"])]


def test_a_copy_is_cleaned_with_what_the_original_shows(tmp_path):
    source = make_record(tmp_path)
    copy = tmp_path / "copy.pdf"
    with pymupdf.open(source) as doc:  # obfuscate_document wrote a placeholder over the label
        page = doc[0]
        page.add_redact_annot(pymupdf.Rect(323, 104, 380, 117), text="PH-0001", fontsize=8)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
        doc.save(copy)
    _, found = obfuscate_signatures(copy, tmp_path / "out.pdf", find_in=source)
    assert any(area.x0 >= 320 and area.y1 <= 205 for _, area in found)


def test_no_label_no_redaction(tmp_path):
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 50), "Approval of the change is pending, see the signature list in annex B.")
    doc.save(tmp_path / "plain.pdf")
    _, found = obfuscate_signatures(tmp_path / "plain.pdf", tmp_path / "out.pdf")
    assert found == []
