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


def test_row_header_names_the_cells_to_its_right(tmp_path):
    # an approval table: row labels on the left, "Unterschrift / SIGNATURE" in two lines of one cell
    doc = pymupdf.open()
    page = doc.new_page()
    shape = page.new_shape()
    columns, rows = (60, 180, 300, 420), (100, 130, 160, 200, 230)
    for x in columns:
        shape.draw_line((x, rows[0]), (x, rows[-1]))
    for y in rows:
        shape.draw_line((columns[0], y), (columns[-1], y))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    page.insert_text((65, 120), "erstellt", fontsize=9)
    page.insert_text((65, 150), "Datum / DATE", fontsize=9)
    page.insert_text((65, 175), "Unterschrift", fontsize=9)
    page.insert_text((65, 190), "SIGNATURE", fontsize=7)
    page.insert_text((65, 220), "Name / NAME", fontsize=9)
    for column, x in enumerate(columns[1:3]):
        page.insert_text((x + 5, 150), "26.5.2000", fontsize=9)
        page.insert_image(pymupdf.Rect(x + 4, 163, x + 116, 197), filename=scribble(tmp_path / f"s{column}.png", (0, 0, 0)))
        page.insert_text((x + 5, 220), "J. Klett", fontsize=9)
    areas = find_signature_areas(page)
    assert any(area.intersects(pymupdf.Rect(190, 170, 290, 190)) for area in areas)
    assert any(area.intersects(pymupdf.Rect(310, 170, 410, 190)) for area in areas)
    assert not any(area.intersects(pymupdf.Rect(62, 162, 178, 198)) for area in areas)  # the label stays
    assert not any(area.intersects(pymupdf.Rect(62, 202, 410, 228)) for area in areas)  # and the names


def test_faint_lines_of_a_scanned_form_close_its_cells(tmp_path):
    # a copy of a form shows its cell borders light gray; the signature sits above its label
    zoom = 2
    image = Image.new("L", (595 * zoom, 842 * zoom), 255)  # a scan: the page is one image
    draw = ImageDraw.Draw(image)
    for x in (10, 290):
        draw.line([(x * zoom, 10 * zoom), (x * zoom, 190 * zoom)], fill=180, width=2)
    for y in (10, 100, 190):
        draw.line([(10 * zoom, y * zoom), (290 * zoom, y * zoom)], fill=180, width=2)
    # the cell beside it: its border stops 2 points short of the line above, as scanned lines do
    draw.line([(200 * zoom, 12 * zoom), (200 * zoom, 100 * zoom)], fill=180, width=2)
    draw.line([(60 * zoom, 80 * zoom), (100 * zoom, 30 * zoom), (140 * zoom, 85 * zoom), (180 * zoom, 40 * zoom)],
              fill=0, width=3)
    path = tmp_path / "form.png"
    image.save(path)
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, filename=str(path))
    page.insert_text((20, 96), "Signature", fontsize=8)
    areas = find_signature_areas(page)
    assert any(area.intersects(pymupdf.Rect(100, 40, 140, 70)) for area in areas)  # the signature
    assert all(area.y1 <= 101 and area.x1 <= 201 for area in areas)  # not the cells below and beside


def ocr_line(text, rect):
    """An OCR line as `_pdf_ocr_lines` gives it, in page coordinates, its characters evenly apart."""
    import numpy as np

    from acd.line_ocr import OcrLine

    width = rect[2] - rect[0]
    positions = [(index * width / len(text), (index + 1) * width / len(text)) for index in range(len(text))]
    to_page = np.array([[1, 0, rect[0]], [0, 1, rect[1]], [0, 0, 1]], dtype=np.float64)
    return OcrLine(text, 0.99, positions, width, rect[3] - rect[1], False, to_page)


def scan(tmp_path, draw_on):
    """A page that is one image, drawn on at 2 pixels per point."""
    image = Image.new("L", (595 * 2, 842 * 2), 255)
    draw_on(ImageDraw.Draw(image))
    image.save(tmp_path / "scan.png")
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, filename=str(tmp_path / "scan.png"))
    return page


def test_handwriting_on_the_dotted_line_of_a_field_is_found(tmp_path):
    # "Design: . . . . . . . ." with a signature on the dots, and an empty field under it
    def draw_on(draw):
        for y in (200, 260):
            for x in range(130, 330, 8):
                draw.rectangle([(x * 2, y * 2 - 2), (x * 2 + 2, y * 2)], fill=0)
        draw.line([(150 * 2, 198 * 2), (180 * 2, 175 * 2), (210 * 2, 202 * 2), (260 * 2, 180 * 2)], fill=0, width=4)

    page = scan(tmp_path, draw_on)
    lines = [ocr_line("Design:", (60, 190, 120, 201)), ocr_line("Commerce:", (60, 250, 125, 261))]
    areas = find_signature_areas(page, lines)
    assert len(areas) == 1
    assert pymupdf.Rect(150, 178, 260, 200) in areas[0] and areas[0].x0 >= 119  # the label stays


def test_labels_over_a_table_head_the_columns_under_them(tmp_path):
    # "Datum:  Kurzz.:" at the foot of a wide cell over the table, initials in the column under "Kurzz."
    doc = pymupdf.open()
    page = doc.new_page()
    shape = page.new_shape()
    shape.draw_rect(pymupdf.Rect(60, 100, 460, 160))
    for y in (200, 240, 280):
        shape.draw_line((60, y), (460, y))
    for x in (60, 300, 380, 460):
        shape.draw_line((x, 160), (x, 280))
    shape.draw_line((60, 160), (460, 160))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    page.insert_text((310, 155), "Datum:", fontsize=9)
    page.insert_text((390, 155), "Kurzz.:", fontsize=9)
    for row, y in enumerate((160, 200)):
        page.insert_text((70, y + 25), "Konstruktion", fontsize=9)
        page.insert_image(pymupdf.Rect(384, y + 4, 456, y + 36), filename=scribble(tmp_path / f"k{row}.png", (0, 0, 0)))
    areas = find_signature_areas(page)
    assert any(area.intersects(pymupdf.Rect(390, 170, 450, 190)) for area in areas)
    assert any(area.intersects(pymupdf.Rect(390, 210, 450, 230)) for area in areas)
    assert not any(area.intersects(pymupdf.Rect(65, 165, 295, 275)) for area in areas)  # the roles stay
