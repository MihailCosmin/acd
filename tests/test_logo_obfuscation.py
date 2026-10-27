"""Tests for acd.logo_obfuscation: logos found by similarity in PDFs and images, and removed."""

import pytest

pytest.importorskip("cv2")
pymupdf = pytest.importorskip("pymupdf")
from PIL import Image
from PIL import ImageDraw

from acd.logo_obfuscation import find_image_logos
from acd.logo_obfuscation import find_pdf_logos
from acd.logo_obfuscation import load_logo_references
from acd.logo_obfuscation import obfuscate_logos

# A wordmark-like logo of bars, rings and a slant, in logo units (width 200, height 40)
BARS = [(0, 0, 12, 40), (20, 0, 32, 40), (20, 0, 52, 10), (20, 30, 52, 40), (60, 0, 72, 40), (60, 15, 90, 25),
        (98, 0, 110, 40), (118, 0, 150, 12), (118, 28, 150, 40), (138, 0, 150, 40), (160, 0, 200, 14)]
RINGS = [(170, 18, 192, 40)]


def draw_logo(draw, x, y, scale, colour):
    for x0, y0, x1, y1 in BARS:
        draw.rectangle([x + x0 * scale, y + y0 * scale, x + x1 * scale, y + y1 * scale], fill=colour)
    for x0, y0, x1, y1 in RINGS:
        draw.ellipse([x + x0 * scale, y + y0 * scale, x + x1 * scale, y + y1 * scale], outline=colour,
                     width=max(1, round(4 * scale)))


@pytest.fixture
def reference(tmp_path):
    image = Image.new("RGB", (440, 120), "white")
    draw_logo(ImageDraw.Draw(image), 20, 20, 2, (0, 40, 140))  # blue, like many logos
    path = tmp_path / "logo.png"
    image.save(path)
    return load_logo_references([path])


def logo_png(tmp_path, scale):
    image = Image.new("RGB", (round(200 * scale) + 20, round(40 * scale) + 20), "white")
    draw_logo(ImageDraw.Draw(image), 10, 10, scale, (0, 0, 0))
    path = tmp_path / f"embedded_{scale}.png"
    image.save(path)
    return path


def make_pdf(tmp_path):
    """A4 page: the logo as an image in the header, as vector paths in a title block, and text."""
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(pymupdf.Rect(380, 30, 560, 66), filename=str(logo_png(tmp_path, 3)))
    shape = page.new_shape()
    x, y, scale = 300, 700, 0.9  # vector logo, 36 pt high
    for x0, y0, x1, y1 in BARS:
        shape.draw_rect(pymupdf.Rect(x + x0 * scale, y + y0 * scale, x + x1 * scale, y + y1 * scale))
    shape.finish(fill=(0, 0, 0), color=None)
    for x0, y0, x1, y1 in RINGS:
        shape.draw_oval(pymupdf.Rect(x + x0 * scale, y + y0 * scale, x + x1 * scale, y + y1 * scale))
    shape.finish(color=(0, 0, 0), width=3.6 * scale, fill=None)
    shape.commit()
    for line in range(12):
        page.insert_text((60, 150 + 22 * line), "INSPECTION SHEET PART NUMBER 1641B0100-01 QTY 4 " * 2, fontsize=9)
    page.insert_text((60, 450), "Heading in large bold type", fontsize=30, fontname="hebo")
    path = tmp_path / "doc.pdf"
    doc.save(path)
    return path


def test_finds_image_and_vector_logos_but_not_text(tmp_path, reference):
    with pymupdf.open(make_pdf(tmp_path)) as doc:
        found = sorted(find_pdf_logos(doc[0], reference), key=lambda match: match.box.y0)
    assert len(found) == 2
    header, title_block = found
    assert header.box.intersects(pymupdf.Rect(380, 30, 560, 66)) and header.box.height > 20
    assert title_block.box.intersects(pymupdf.Rect(300, 700, 480, 736))
    assert all(match.score >= 0.65 and match.reference == "logo.png" for match in found)


def test_pdf_logos_are_redacted(tmp_path, reference):
    output, found = obfuscate_logos(make_pdf(tmp_path), reference, tmp_path / "out.pdf")
    assert len(found) == 2
    with pymupdf.open(output) as doc:
        page = doc[0]
        assert find_pdf_logos(page, reference) == []
        assert "INSPECTION SHEET" in page.get_text()  # the text around stays
        # the vector logo's paths are gone, not only painted over
        assert not [path for path in page.get_drawings() if path["rect"].intersects(pymupdf.Rect(310, 705, 470, 730))
                    and path.get("fill") == (0.0, 0.0, 0.0) and path["rect"].width < 40]


def test_image_logo_is_painted_over(tmp_path, reference):
    image = Image.new("RGB", (1200, 900), "white")
    draw = ImageDraw.Draw(image)
    draw_logo(draw, 700, 60, 2.2, (0, 40, 140))
    for line in range(20):
        draw.text((60, 300 + 25 * line), "Part 1641B0100-01 inspected by J. Miller " * 3, fill="black")
    source = tmp_path / "scan.png"
    image.save(source)
    assert len(find_image_logos(Image.open(source), reference)) == 1
    output, found = obfuscate_logos(source, reference, tmp_path / "scan_out.png")
    assert len(found) == 1
    x0, y0, x1, y1 = found[0].box
    assert x0 < 710 and x1 > 1130 and y0 < 70 and y1 > 140
    with Image.open(output) as result:
        assert result.getpixel((800, 100)) == (0, 0, 0)
        assert find_image_logos(result, reference) == []


def test_logos_found_in_the_original_are_removed_from_the_copy(tmp_path, reference):
    source = make_pdf(tmp_path)
    copy = tmp_path / "copy.pdf"
    with pymupdf.open(source) as doc:  # as obfuscate_document leaves it: a black box over the header logo
        doc[0].draw_rect(pymupdf.Rect(380, 30, 560, 66), color=None, fill=(0, 0, 0))
        doc.save(copy)
    with pymupdf.open(copy) as doc:
        assert len(find_pdf_logos(doc[0], reference)) == 1  # the copy no longer shows the header logo
    _, found = obfuscate_logos(copy, reference, tmp_path / "out.pdf", find_in=source)
    assert len(found) == 2


def test_filled_box_is_no_logo(tmp_path, reference):
    image = Image.new("RGB", (900, 400), "white")
    ImageDraw.Draw(image).rectangle([200, 150, 640, 238], fill="black")
    assert find_image_logos(image, reference) == []


def test_noise_and_photos_are_no_logo(tmp_path, reference):
    import numpy

    random = numpy.random.default_rng(1)
    noise = Image.fromarray((random.random((700, 900)) * 255).astype(numpy.uint8)).convert("RGB")
    assert find_image_logos(noise, reference) == []


def test_blank_reference_is_refused(tmp_path):
    path = tmp_path / "blank.png"
    Image.new("RGB", (100, 30), "white").save(path)
    with pytest.raises(ValueError):
        load_logo_references([path])
