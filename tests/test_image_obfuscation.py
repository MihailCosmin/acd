"""Obfuscation of image files (`acd.image_obfuscation`).

Images are read with OCR: these tests need RapidOCR (`pip install rapidocr
onnxruntime`) and are skipped without it -- all but those that paint, or feed
in a broken file, which need only the OpenCV RapidOCR brings along. The images
are drawn here with Pillow, so their text is known.
"""

import io
import re
import struct

import numpy as np
import pytest
from PIL import ExifTags
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
from PIL import PngImagePlugin

pytest.importorskip("cv2")

from acd import image_obfuscation  # noqa: E402
from acd.line_ocr import OcrLine  # noqa: E402
from acd.obfuscation import obfuscate_document  # noqa: E402

needs_ocr = pytest.mark.skipif(
    __import__("importlib").util.find_spec("rapidocr") is None,
    reason="OCR needs rapidocr: pip install rapidocr onnxruntime",
)

TARGETS = {"Heike Wochner": "blackout", "D9893": "placeholder", re.compile(r"\b1644-\d{4}\b"): "placeholder"}
FONT_SIZE = 36


def drawing(lines, size=(1000, 300), fill="black"):
    """A white RGB image with `lines` of text."""
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    for number, text in enumerate(lines):
        draw.text((40, 40 + 90 * number), text, fill=fill, font=ImageFont.load_default(size=FONT_SIZE))
    return image


def bilevel(image):
    """An image at one bit per pixel, as a fax or a drawing scan is."""
    return image.convert("L").point(lambda value: 255 if value > 128 else 0).convert("1")


def read(image):
    """What OCR reads in an image, one line per line."""
    from acd.line_ocr import ocr_lines
    if image.mode == "I;16":
        image = Image.fromarray((np.asarray(image) // 257).astype(np.uint8))
    return [line.text for line in ocr_lines(np.asarray(image.convert("L")))]


TEXT = ["Created by Heike Wochner today", "Part 1644-0035 cage D9893 stays", "keep this line as it is"]


@needs_ocr
@pytest.mark.parametrize("name, convert, params", [
    ("png", None, {}),
    ("jpg", None, {"quality": 92}),
    ("tif", bilevel, {"compression": "group4"}),
    ("webp", None, {"quality": 90}),
    ("avif", None, {"quality": 90}),
    ("gif", lambda image: image.convert("P", palette=Image.Palette.ADAPTIVE, colors=16), {}),
    ("png", lambda image: image.convert("RGBA"), {}),
    ("png", lambda image: Image.fromarray(np.asarray(image.convert("L"), dtype=np.uint16) * 257), {}),
])
def test_formats(tmp_path, name, convert, params):
    image = drawing(TEXT)
    image = convert(image) if convert else image
    source = tmp_path / f"sample.{name}"
    image.save(source, **params)

    output, mapping = obfuscate_document(source, TARGETS)

    assert output == str(tmp_path / f"sample_obfuscated.{name}")
    assert mapping == {"Heike Wochner": "█████ ███████", "1644-0035": "PH-0001", "D9893": "PH-0002"}
    with Image.open(output) as result:
        assert (result.format, result.mode) == (Image.open(source).format, image.mode)
        if name == "tif":
            assert result.info["compression"] == "group4"
        text = " | ".join(read(result))
    assert "Heike" not in text and "Wochner" not in text and "D9893" not in text and "1644-0035" not in text
    for kept in ("Created by", "PH-0001", "PH-0002", "stays", "keep this line as it is"):
        assert kept in text, text


def exif_with_thumbnail():
    """EXIF as a camera writes it: text in IFD0 and the Exif directory, a
    maker note, and a thumbnail of the picture in IFD1."""
    thumbnail = io.BytesIO()
    Image.new("RGB", (32, 16), "red").save(thumbnail, "JPEG")
    thumbnail = thumbnail.getvalue()

    def entry(tag, kind, count, value):
        return struct.pack("<HHI", tag, kind, count) + value

    description, artist = b"Part 1644-0035 by D9893\0", b"Heike Wochner\0"
    author = "Heike Wochner".encode("utf-16-le") + b"\0\0"
    comment, maker = b"ASCII\0\0\0made for D9893", b"SERIAL Heike Wochner"
    data = 8 + 2 + 4 * 12 + 4  # header, IFD0 with 4 entries
    exif_ifd = data + len(description) + len(artist) + len(author)
    exif_data = exif_ifd + 2 + 2 * 12 + 4
    ifd1 = exif_data + len(comment) + len(maker)
    return thumbnail, (
        b"Exif\0\0II*\0" + struct.pack("<I", 8)
        + struct.pack("<H", 4)
        + entry(0x010E, 2, len(description), struct.pack("<I", data))
        + entry(0x013B, 2, len(artist), struct.pack("<I", data + len(description)))
        + entry(0x9C9D, 1, len(author), struct.pack("<I", data + len(description) + len(artist)))
        + entry(0x8769, 4, 1, struct.pack("<I", exif_ifd))
        + struct.pack("<I", ifd1)
        + description + artist + author
        + struct.pack("<H", 2)
        + entry(0x9286, 7, len(comment), struct.pack("<I", exif_data))
        + entry(0x927C, 7, len(maker), struct.pack("<I", exif_data + len(comment)))
        + struct.pack("<I", 0)
        + comment + maker
        + struct.pack("<H", 2)
        + entry(0x0201, 4, 1, struct.pack("<I", ifd1 + 30))
        + entry(0x0202, 4, 1, struct.pack("<I", len(thumbnail)))
        + struct.pack("<I", 0)
        + thumbnail
    )


@needs_ocr
def test_jpeg_metadata(tmp_path):
    thumbnail, exif = exif_with_thumbnail()
    xmp = b'<x:xmpmeta xmlns:x="adobe:ns:meta/"><dc:creator>Heike Wochner</dc:creator></x:xmpmeta>'
    source = tmp_path / "photo.jpg"
    Image.new("RGB", (300, 120), "white").save(
        source, exif=exif, xmp=xmp, comment=b"Scanned by Heike Wochner, cage D9893"
    )
    assert thumbnail[2:40] in source.read_bytes()

    output, mapping = obfuscate_document(source, TARGETS)

    assert mapping == {"1644-0035": "PH-0001", "D9893": "PH-0002", "Heike Wochner": "█████ ███████"}
    with Image.open(output) as result:
        exif = result.getexif()
        assert exif[0x010E] == "Part PH-0001 by PH-0002"
        assert exif[0x013B] == "XXXXX XXXXXXX"  # an ASCII tag cannot hold the block character
        assert exif[0x9C9D].decode("utf-16-le").rstrip("\0") == "█████ ███████"
        details = exif.get_ifd(0x8769)
        assert details[0x9286] == b"UNICODE\0" + "made for PH-0002".encode("utf-16-be")
        assert 0x927C not in details  # the maker note goes
        assert dict(exif.get_ifd(ExifTags.IFD.IFD1)) == {}  # and the thumbnail
        assert "xmp" not in result.info  # XMP with a match goes
        assert result.info["comment"].decode() == "Scanned by █████ ███████, cage PH-0002"
    data = open(output, "rb").read()
    for gone in (thumbnail[2:40], b"Heike", "Heike".encode("utf-16-le"), "Heike".encode("utf-16-be"), b"D9893", b"1644-0035"):
        assert gone not in data


@needs_ocr
def test_png_text_chunks(tmp_path):
    chunks = PngImagePlugin.PngInfo()
    chunks.add_text("Author", "Heike Wochner")
    chunks.add_itxt("Description", "Grüße von Heike Wochner, 1644-0035")
    chunks.add_text("Software", "ACME Scan")
    source = tmp_path / "scan.png"
    Image.new("RGB", (300, 120), "white").save(source, pnginfo=chunks)

    output, mapping = obfuscate_document(source, TARGETS)

    with Image.open(output) as result:
        assert result.text == {
            "Author": "█████ ███████", "Description": "Grüße von █████ ███████, PH-0001", "Software": "ACME Scan",
        }


@needs_ocr
def test_every_page_of_a_tiff_keeps_its_mode_and_compression(tmp_path):
    first = drawing(["page one by Heike Wochner"], size=(900, 150))
    second = bilevel(drawing(["page two, cage D9893"], size=(900, 150)))
    second.encoderinfo = {"compression": "group4"}
    source = tmp_path / "pages.tif"
    first.save(source, save_all=True, append_images=[second], compression="tiff_lzw")

    output, mapping = obfuscate_document(source, TARGETS)

    assert mapping == {"Heike Wochner": "█████ ███████", "D9893": "PH-0001"}
    with Image.open(output) as result:
        assert result.n_frames == 2
        pages = []
        for index in range(2):
            result.seek(index)
            pages.append((result.mode, result.info["compression"], " ".join(read(result))))
    assert pages == [("RGB", "tiff_lzw", "page one by"), ("1", "group4", "page two, cage PH-0001")]


@needs_ocr
@pytest.mark.parametrize("name, params", [("gif", {}), ("webp", {"lossless": True}), ("png", {})])
def test_animations_keep_frames_and_timing(tmp_path, name, params):
    frames = [drawing([text], size=(900, 150)) for text in ("frame one", "frame two Heike Wochner", "frame three D9893")]
    source = tmp_path / f"animation.{name}"
    frames[0].save(source, save_all=True, append_images=frames[1:], duration=[200, 300, 400], loop=0, **params)

    output, mapping = obfuscate_document(source, TARGETS)

    assert mapping == {"Heike Wochner": "█████ ███████", "D9893": "PH-0001"}
    with Image.open(output) as result:
        timing, texts = [], []
        for index in range(result.n_frames):
            result.seek(index)
            result.load()
            timing.append(int(result.info["duration"]))
            texts.append(" ".join(read(result)))
    assert timing == [200, 300, 400]
    assert texts == ["frame one", "frame two", "frame three PH-0001"]


@needs_ocr
def test_a_format_pillow_cannot_write_comes_back_as_png(tmp_path):
    source = tmp_path / "tiny.xpm"
    source.write_text('/* XPM */\nstatic char *x[] = {\n"4 2 2 1",\n"a c #000000",\n"b c #FFFFFF",\n"abab",\n"baba"\n};\n')

    output, mapping = obfuscate_document(source, TARGETS)

    assert output == str(tmp_path / "tiny_obfuscated.png")
    with Image.open(output) as result:
        assert (result.format, result.size) == ("PNG", (4, 2))


def test_files_that_are_no_images(tmp_path):
    broken = tmp_path / "broken.png"
    broken.write_bytes(b"not an image at all")
    with pytest.raises(ValueError, match="Cannot read .* as an image"):
        obfuscate_document(broken, TARGETS)
    assert ".jpg" in image_obfuscation.image_extensions()
    assert ".avif" in image_obfuscation.image_extensions()


def a_line(text, left, top, width, height):
    """An OCR line as `line_ocr` returns it: `text` evenly spread over a
    straight box."""
    pitch = width / len(text)
    to_image = np.array([[1.0, 0, left], [0, 1.0, top], [0, 0, 1.0]])
    positions = [(index * pitch, (index + 1) * pitch) for index in range(len(text))]
    return OcrLine(text, 1.0, positions, width, height, False, to_image)


@pytest.mark.parametrize("mode", ["1", "L", "P", "RGB", "RGBA", "CMYK", "I;16", "I", "F", "LA"])
def test_painting_in_every_pixel_mode(mode):
    frame = drawing(["D9893 x"], size=(400, 120)).convert("L" if mode in ("I;16", "I", "F") else "RGB")
    if mode in ("I;16", "I"):
        frame = Image.fromarray(np.asarray(frame, dtype=np.uint16) * 257).convert(mode)
    elif mode == "F":
        frame = Image.fromarray(np.asarray(frame, dtype=np.float32) / 255)
    else:
        frame = frame.convert(mode)
    line = a_line("D9893", 40, 40, 110, 40)

    black = frame.copy()
    image_obfuscation._paint(black, line, 0, 5, "█████")
    placeholder = frame.copy()
    image_obfuscation._paint(placeholder, line, 0, 5, "PH-0001")

    def luminance(image):  # 0 black ... 1 white
        values = np.asarray(image.convert("L") if image.mode not in ("I;16", "I", "F") else image, dtype=np.float64)
        top = {"I;16": 65535, "I": 65535, "F": 1.0}.get(image.mode, 255)
        return values[48:72, 60:130] / top

    assert luminance(black).max() < 0.1  # all of the match black
    painted = luminance(placeholder)
    assert painted.mean() > 0.6 and painted.min() < 0.3  # white, with the placeholder's dark strokes


def test_a_blackout_is_opaque_where_the_transparent_colour_is_black():
    # GIFs often make black their transparent colour: painted in it, a
    # blackout would come out clear
    image = Image.new("P", (400, 120), 1)
    image.putpalette([0, 0, 0, 255, 255, 255, 40, 40, 40, 0, 0, 0])  # transparent, white, text, black
    image.info["transparency"] = 0
    ImageDraw.Draw(image).text((40, 40), "D9893 x", fill=2, font=ImageFont.load_default(size=FONT_SIZE))

    image_obfuscation._paint(image, a_line("D9893", 40, 40, 110, 40), 0, 5, "█████")

    assert set(np.asarray(image)[48:72, 60:130].ravel().tolist()) == {3}
