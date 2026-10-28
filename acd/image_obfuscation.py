"""Obfuscation of image files, for `obfuscation.obfuscate_document`.

Every format Pillow reads: JPEG, PNG, TIFF, WebP, AVIF, GIF, BMP, JPEG 2000,
ICO, TGA, PPM, PSD and the rest -- HEIC/HEIF too with `pillow-heif`
installed, JPEG XL with `pillow-jxl-plugin`. An image is written back in its
own format, or as PNG where Pillow cannot write that (PSD, XPM, ...).

An image's text is its pixels, so it is read with OCR (`line_ocr`: `pip install
rapidocr onnxruntime`), every frame of it -- the pages of a TIFF, the frames of
an animation. A match is painted over in the pixels: black, or in the colour
around it with its placeholder written across it in the text's colour and
direction. The image is encoded anew, so nothing of the original pixels
under a match survives; what OCR cannot read stays -- check the result.

Metadata is obfuscated as text is: EXIF descriptions, authors and comments
(the Windows XP tags and UserComment too), PNG text chunks, JPEG and GIF
comments, TIFF text tags. What cannot be obfuscated goes: the EXIF thumbnail
(a small copy of the original image), maker notes, IPTC and Photoshop blocks,
and XMP with a match in it. The image itself is kept as it was where the
format allows: mode, palette and transparency, resolution, colour profile,
JPEG quantisation, TIFF compression per page, animation timing.
"""

from contextlib import contextmanager

from html import unescape

from math import atan2
from math import degrees
from math import sqrt

from os import remove
from os.path import splitext

from typing import Optional

import numpy as np

from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
from PIL import JpegImagePlugin
from PIL import PngImagePlugin

from .obfuscation import BLACKOUT_CHAR
from .obfuscation import ObfuscationError
from .obfuscation import _OCR_PITCH
from .obfuscation import _Replacements
from .obfuscation import _every_occurrence
from .obfuscation import _identity
from .obfuscation import _is_blackout
from .obfuscation import _kinds_by_identity
from .obfuscation import _known_strings
from .obfuscation import _leftover_matches
from .obfuscation import _match_regions
from .obfuscation import _misreadings
from .obfuscation import _ocr_regions
from .obfuscation import _replace_regions


# Frames are read for OCR at no more than this many pixels: an A0 sheet
# scanned at 400 dpi is read at half size, where its lettering is still some
# 20 pixels high.
_OCR_MAX_PIXELS = 64_000_000

# A placeholder too long for what it replaces shrinks to fit, but not below
# this height in pixels (it overflows instead).
_MIN_PLACEHOLDER_PIXELS = 8

# EXIF tags of the image's own directory that are carried over -- metadata,
# not structure: DocumentName, ImageDescription, Make, Model, Orientation,
# PageName, Software, DateTime, Artist, HostComputer, Copyright and the
# Windows XP title, comment, author, keywords and subject.
_EXIF_KEPT = (
    0x010D, 0x010E, 0x010F, 0x0110, 0x0112, 0x011D, 0x0131, 0x0132, 0x013B, 0x013C, 0x8298,
    0x9C9B, 0x9C9C, 0x9C9D, 0x9C9E, 0x9C9F,
)
_EXIF_IFD = 0x8769
_GPS_IFD = 0x8825
_EXIF_XMP = 0x02BC  # where TIFF keeps XMP

# Sub-directory tags left out: the maker note (vendor data, opaque) and the
# interoperability directory.
_EXIF_DROPPED = (0x927C, 0xA005)

# Timestamps: a digit pattern would only mangle them.
_EXIF_TIMESTAMPS = (0x0132, 0x9003, 0x9004, 0x001D)

# Windows XP tags hold UTF-16; UserComment, GPSProcessingMethod and
# GPSAreaInformation an 8-byte character code, then the text.
_EXIF_XP = (0x9C9B, 0x9C9C, 0x9C9D, 0x9C9E, 0x9C9F)
_EXIF_CODED = (0x9286, 0x001B, 0x001C)

# What each format can carry of the metadata.
_SAVES_EXIF = ("JPEG", "MPO", "PNG", "TIFF", "WEBP", "AVIF", "HEIF")
_SAVES_XMP = ("JPEG", "MPO", "WEBP", "AVIF", "HEIF")  # PNG and TIFF keep it their own way
_SAVES_COMMENT = ("JPEG", "MPO", "GIF", "JPEG2000")
_SAVES_ICC = ("JPEG", "MPO", "PNG", "TIFF", "WEBP", "AVIF", "HEIF", "JPEG2000")
_SAVES_DPI = ("JPEG", "MPO", "PNG", "TIFF", "BMP", "PCX", "TGA")
_SAVES_FRAMES = ("TIFF", "GIF", "PNG", "WEBP", "AVIF", "MPO", "HEIF")

# TIFF compressions Pillow writes as read (old-style JPEG and the rest become
# JPEG or Deflate).
_TIFF_COMPRESSIONS = ("raw", "tiff_lzw", "tiff_deflate", "tiff_adobe_deflate", "packbits", "group3", "group4")

_plugins = []


def image_extensions() -> set:
    """The file extensions read as images: those of every format Pillow, with
    the HEIF and JPEG XL plugins where installed, can open."""
    _load_plugins()
    return {
        extension for extension, format_ in Image.registered_extensions().items()
        if (format_ in Image.OPEN or format_ == "MPO") and extension != ".pdf"
    }


def default_image_output(filepath: str) -> str:
    """`<name>_obfuscated<ext>` beside the input -- `.png` if Pillow cannot
    write the input's format."""
    stem, suffix = splitext(filepath)
    return stem + "_obfuscated" + (suffix if _writable(suffix) else ".png")


def _load_plugins() -> None:
    if _plugins:
        return
    try:
        from pillow_heif import register_heif_opener
        register_heif_opener()
        _plugins.append("pillow_heif")
    except ImportError:
        pass
    try:
        import pillow_jxl  # noqa: F401 -- registers itself
        _plugins.append("pillow_jxl")
    except ImportError:
        pass
    _plugins.append(None)


def _writable(suffix: str) -> bool:
    return Image.registered_extensions().get(suffix.lower()) in Image.SAVE


@contextmanager
def _any_size():
    """Lift Pillow's decompression-bomb limit (179 megapixels), which an A0
    drawing scanned at 400 dpi is over: these are the caller's own files."""
    limit = Image.MAX_IMAGE_PIXELS
    Image.MAX_IMAGE_PIXELS = None
    try:
        yield
    finally:
        Image.MAX_IMAGE_PIXELS = limit


def _obfuscate_image(
    filepath: str, output_path: str, rules: list, placeholders: Optional[dict], placeholder_format: str
) -> dict:
    """Obfuscate an image file: see the module docstring."""
    from .line_ocr import ocr_lines  # needs RapidOCR, which images always do

    _load_plugins()
    with _any_size():
        try:
            with Image.open(filepath) as source:
                frames = _frames(source)
                original = _Original(source, filepath)
        except OSError as error:
            raise ValueError(f"Cannot read {filepath} as an image: {error}") from error
        target = Image.registered_extensions().get(splitext(output_path)[1].lower())
        if target not in Image.SAVE:
            raise ValueError(f"Cannot write images as {splitext(output_path)[1] or 'extensionless'} files")
        if len(frames) > 1 and target not in _SAVES_FRAMES:
            raise ValueError(f"{len(frames)} frames do not fit into a {target} file: write a TIFF")

        # As for a PDF: find everything first, then obfuscate it.
        read = {}  # OCR of each distinct frame -- an animation repeats many
        lines = []
        for frame in frames:
            key = _fingerprint(frame)
            if key not in read:
                read[key] = _frame_lines(frame, ocr_lines)
            lines.append(read[key])
        fields = original.texts()
        texts = [line.text for frame_lines in lines for line in frame_lines] + fields
        rules = _every_occurrence(rules, texts)
        found = [(text, _match_regions(text, rules)) for text in texts]
        kinds = _kinds_by_identity(found)
        misreadings = _misreadings(_known_strings(rules, found))
        for line in (line for frame_lines in lines for line in frame_lines):
            for _, _, kind, original_text in _ocr_regions(line.text, rules, misreadings):
                if kinds.get(_identity(original_text)) != "blackout":
                    kinds[_identity(original_text)] = kind
        replacements = _Replacements(rules, kinds, placeholders, placeholder_format, "\n".join(texts))

        for frame, frame_lines in zip(frames, lines):
            for line in frame_lines:
                for start, end, _, original_text in _ocr_regions(line.text, rules, misreadings):
                    _paint(frame, line, start, end, replacements[original_text])
        params = original.save_params(target, frames, rules, replacements)
        _save(frames, output_path, target, params)

    leftovers = _image_leftovers(output_path, rules, replacements.placeholders())
    if leftovers:
        remove(output_path)
        raise ObfuscationError(f"Obfuscating {filepath} left matches readable in: {', '.join(leftovers)}")
    return replacements.mapping


def _frames(image) -> list:
    """Copies of every frame of an image, each with its own `info`."""
    if image.format == "PSD":
        # Pillow's PSD "frames" are its layers, numbered from 1 (seek(0)
        # raises EOFError): the image is the composite, as opened
        frame = image.copy()
        frame.info = dict(image.info)
        return [frame]
    frames = []
    for index in range(getattr(image, "n_frames", 1)):
        image.seek(index)
        frame = image.copy()
        frame.info = dict(image.info)
        frames.append(frame)
    image.seek(0)
    return frames


def _fingerprint(frame) -> tuple:
    """What tells frames apart."""
    return frame.mode, frame.size, hash(frame.tobytes())


def _frame_lines(frame, ocr_lines) -> list:
    """OCR lines of a frame, mapping into its pixels."""
    gray = _gray(frame)
    scale = min(1.0, sqrt(_OCR_MAX_PIXELS / max(gray.width * gray.height, 1)))
    if scale < 1:
        size = (max(round(gray.width * scale), 1), max(round(gray.height * scale), 1))
        gray = gray.resize(size, Image.Resampling.BOX)
    back = np.diag([frame.width / gray.width, frame.height / gray.height, 1.0])  # OCR pixels -> frame pixels
    return [line._replace(to_image=back @ line.to_image) for line in ocr_lines(np.asarray(gray))]


def _gray(frame):
    """A frame as OCR reads it: 8-bit grayscale, transparency on white, 16-bit
    and float data stretched to their range."""
    if frame.mode in ("I;16", "I;16L", "I;16B", "I;16N", "I", "F"):
        values = np.asarray(frame, dtype=np.float64)
        low, high = float(values.min()), float(values.max())
        values = (values - low) * (255 / (high - low)) if high > low else np.zeros_like(values)
        return Image.fromarray(values.astype(np.uint8))
    if "A" in frame.mode or "transparency" in frame.info:
        rgba = frame.convert("RGBA")
        return Image.alpha_composite(Image.new("RGBA", rgba.size, "white"), rgba).convert("L")
    return frame.convert("L")


def _paint(frame, line, start: int, end: int, replacement: str) -> None:
    """Paint over line.text[start:end] in a frame: black for a blackout, else
    the colour around it with the placeholder across it."""
    corners = line.quad(start, end)
    polygon = [tuple(corner) for corner in corners]
    if _is_blackout(replacement):
        ImageDraw.Draw(frame).polygon(polygon, fill=_ink(frame, (0, 0, 0, 255)))
        return
    background, text = _colours(frame, corners)
    ImageDraw.Draw(frame).polygon(polygon, fill=background)
    _write(frame, line.quad(start, end, pad=0), end - start, replacement, text)


def _numeric(frame) -> bool:
    return frame.mode in ("I;16", "I;16L", "I;16B", "I;16N", "I", "F")


def _colours(frame, corners) -> tuple:
    """Inks, in the frame's own terms, for the background around a match and
    for its text: the commonest value in its box, and the commonest of those
    far from it -- black or white if the text left nothing to tell by."""
    x0, y0 = np.floor(corners.min(axis=0)).astype(int)
    x1, y1 = np.ceil(corners.max(axis=0)).astype(int)
    box = (max(x0, 0), max(y0, 0), min(max(x1, x0 + 1), frame.width), min(max(y1, y0 + 1), frame.height))
    crop = frame.crop(box)
    if _numeric(frame):
        values = np.asarray(crop, dtype=np.float64).ravel()
        low, high = frame.getextrema()
        background = float(np.median(values))
        text = low if background - low > high - background else high
        cast = float if frame.mode == "F" else int
        return cast(background), cast(text)
    pixels = np.asarray(crop.convert("RGBA"), dtype=np.int32).reshape(-1, 4)
    background = np.median(pixels, axis=0)
    far = pixels[np.abs(pixels - background).sum(axis=1) > 96]
    if len(far) > 0.02 * len(pixels):
        text = np.median(far, axis=0)
    else:
        luminance = 0.299 * background[0] + 0.587 * background[1] + 0.114 * background[2]
        text = np.array([0, 0, 0, 255] if luminance > 127 else [255, 255, 255, 255])
    return _ink(frame, tuple(int(v) for v in background)), _ink(frame, tuple(int(v) for v in text))


def _ink(frame, rgba: tuple):
    """An RGBA colour as a value of the frame's mode: a palette index (the
    transparent one for a transparent colour) in a palette image."""
    if _numeric(frame):
        low, high = frame.getextrema()
        value = low + (high - low) * (0.299 * rgba[0] + 0.587 * rgba[1] + 0.114 * rgba[2]) / 255
        return float(value) if frame.mode == "F" else int(round(value))
    if frame.mode in ("P", "PA"):
        transparency = frame.info.get("transparency")
        if rgba[3] < 128 and isinstance(transparency, int):
            return transparency
        palette = np.array(frame.getpalette(), dtype=np.int64).reshape(-1, 3)
        distances = ((palette - np.array(rgba[:3])) ** 2).sum(axis=1)
        if isinstance(transparency, int) and transparency < len(distances):
            distances[transparency] = distances.max() + 1  # often black too: a blackout would come out clear
        return int(np.argmin(distances))
    return Image.new("RGBA", (1, 1), rgba).convert(frame.mode).getpixel((0, 0))


def _write(frame, corners, length: int, placeholder: str, ink) -> None:
    """Write a placeholder over the `length` characters OCR read in `corners`
    (top-left, top-right, bottom-right, bottom-left of them the right way
    up): at their size, shrunk if longer, along their line."""
    top_left, top_right, _, bottom_left = corners
    along, down = top_right - top_left, bottom_left - top_left
    size = min(np.hypot(*along) / max(length, 1) / _OCR_PITCH, np.hypot(*down))
    font = ImageFont.load_default(size=max(size, 4))
    left, top, right, bottom = font.getbbox(placeholder)
    if right - left > np.hypot(*along) > 0:
        size = max(size * np.hypot(*along) / (right - left), min(size, _MIN_PLACEHOLDER_PIXELS))
        font = ImageFont.load_default(size=max(size, 4))
        left, top, right, bottom = font.getbbox(placeholder)
    mask = Image.new("L", (right - left + 4, bottom - top + 4))
    ImageDraw.Draw(mask).text((2 - left, 2 - top), placeholder, fill=255, font=font)
    angle = degrees(atan2(-along[1], along[0])) % 360
    square = round(angle / 90) * 90 % 360
    if min(abs(angle - square), 360 - abs(angle - square)) < 3:  # a scan's skew of a degree or two
        angle = square
    mask = mask.rotate(angle, resample=Image.Resampling.BICUBIC, expand=True)
    if frame.mode in ("P", "PA", "1"):  # palette indices and bits do not blend: no anti-aliasing
        mask = mask.point(lambda value: 255 if value > 127 else 0)
    centre_x, centre_y = corners.mean(axis=0)
    frame.paste(ink, (round(centre_x - mask.width / 2), round(centre_y - mask.height / 2)), mask)


class _Original:
    """What an image carries besides its pixels -- metadata, resolution,
    colour profile, encoding settings -- read before the frames are
    painted, and carried over, obfuscated, when they are saved."""

    def __init__(self, image, path: str):
        self.format = image.format
        self.info = dict(image.info)
        exif = image.getexif()
        self.exif = {tag: exif[tag] for tag in _EXIF_KEPT if tag in exif}
        self.sub_ifds = {
            ifd: {tag: value for tag, value in exif.get_ifd(ifd).items() if tag not in _EXIF_DROPPED}
            for ifd in (_EXIF_IFD, _GPS_IFD)
        }
        xmp = self.info.get("xmp") or self.info.get("XML:com.adobe.xmp") or exif.get(_EXIF_XMP)
        self.xmp = xmp.decode("utf-8", "replace") if isinstance(xmp, bytes) else xmp
        comment = self.info.get("comment")
        self.comment = comment.decode("utf-8", "replace") if isinstance(comment, bytes) else comment
        self.text = {
            key: value for key, value in getattr(image, "text", {}).items()
            if key != "XML:com.adobe.xmp" and isinstance(value, str)
            # ImageMagick's hex-encoded EXIF/IPTC/XMP copies: a match in them
            # can't be seen, let alone obfuscated (the EXIF is kept, cleaned)
            and not key.startswith("Raw profile type ")
        }
        self.quantization = getattr(image, "quantization", None)
        self.subsampling = JpegImagePlugin.get_sampling(image) if image.format in ("JPEG", "MPO") else -1
        self.lossless = image.format == "WEBP" and _webp_lossless(path)
        self.loop = self.info.get("loop")

    def texts(self) -> list:
        """The metadata text the targets may occur in."""
        texts = [self.comment] if self.comment else []
        texts += list(self.text.values())
        for tag, value in self._exif_items():
            text = _exif_text(tag, value)
            if text and tag not in _EXIF_TIMESTAMPS:
                texts.append(text)
        return texts

    def _exif_items(self):
        yield from self.exif.items()
        for ifd in self.sub_ifds.values():
            yield from ifd.items()

    def save_params(self, target: str, frames: list, rules: list, replacements: _Replacements) -> dict:
        """Keyword arguments for saving the frames as `target`, the metadata
        obfuscated."""
        def edit(text, blackout_char=BLACKOUT_CHAR):
            return _replace_regions(text, _match_regions(text, rules), replacements, blackout_char)

        params = {}
        exif = Image.Exif()
        for tag, value in self.exif.items():
            value = _edit_exif(tag, value, edit, rules)
            if value is not None:
                exif[tag] = value
        for ifd, tags in self.sub_ifds.items():
            kept = {tag: _edit_exif(tag, value, edit, rules) for tag, value in tags.items()}
            kept = {tag: value for tag, value in kept.items() if value is not None}
            if kept:
                exif[ifd] = kept
        xmp = self.xmp if self.xmp and not _match_regions(unescape(self.xmp), rules) else None
        if target == "TIFF" and xmp:
            exif[_EXIF_XMP] = xmp.encode("utf-8")
        if target in _SAVES_EXIF and len(exif):
            params["exif"] = exif.tobytes()
        if target in _SAVES_XMP and xmp:
            params["xmp"] = xmp.encode("utf-8")
        if target in _SAVES_COMMENT and self.comment:
            params["comment"] = edit(self.comment).encode("utf-8")
        if target == "PNG" and (self.text or xmp):
            chunks = PngImagePlugin.PngInfo()
            for key, value in self.text.items():
                chunks.add_itxt(key, edit(value))
            if xmp:
                chunks.add_itxt("XML:com.adobe.xmp", xmp)
            params["pnginfo"] = chunks
        if target in _SAVES_ICC and self.info.get("icc_profile"):
            params["icc_profile"] = self.info["icc_profile"]
        if target in _SAVES_DPI and self.info.get("dpi"):
            params["dpi"] = tuple(round(float(value)) for value in self.info["dpi"])
        if target in ("PNG", "GIF") and "transparency" in self.info and frames[0].mode in ("P", "L", "RGB", "1"):
            params["transparency"] = self.info["transparency"]
        if target in ("JPEG", "MPO") and self.format in ("JPEG", "MPO") and self.quantization:
            params["qtables"] = self.quantization  # as compressed as before, no more
            if self.subsampling >= 0:
                params["subsampling"] = self.subsampling
            params["progressive"] = bool(self.info.get("progressive") or self.info.get("progression"))
        if target in ("WEBP", "AVIF"):
            params["quality"] = 90
        if target == "WEBP" and self.lossless:
            params["lossless"] = True
        if target == "ICO" and self.info.get("sizes"):
            params["sizes"] = list(self.info["sizes"])
        if len(frames) > 1:
            params["save_all"] = True
            params["append_images"] = frames[1:]
            if target in ("GIF", "PNG", "WEBP", "AVIF"):
                params["duration"] = [frame.info.get("duration", 100) for frame in frames]
                params["loop"] = self.loop if self.loop is not None else 0
        if target == "TIFF":
            for frame in frames:  # each page keeps its compression, where its mode allows it
                frame.encoderinfo = {"compression": _tiff_compression(frame)}
                if frame.info.get("dpi"):
                    frame.encoderinfo["dpi"] = tuple(round(float(v)) for v in frame.info["dpi"])
            params.update(frames[0].encoderinfo)
        return params


def _webp_lossless(path: str) -> bool:
    """Whether a WebP file is encoded lossless: VP8L bitstreams, no VP8 one."""
    with open(path, "rb") as file_:
        data = file_.read()
    return b"VP8L" in data and b"VP8 " not in data


def _tiff_compression(frame) -> str:
    """The TIFF compression a page had, or the nearest one its mode allows."""
    compression = frame.info.get("compression", "raw")
    if compression in ("group3", "group4") and frame.mode != "1":
        return "tiff_deflate"
    if compression in ("jpeg", "tiff_jpeg"):
        return "jpeg" if frame.mode in ("L", "RGB", "CMYK", "YCbCr") else "tiff_deflate"
    if compression not in _TIFF_COMPRESSIONS:
        return "tiff_deflate"
    return compression


def _exif_text(tag: int, value) -> Optional[str]:
    """The text of an EXIF value that holds text, else None."""
    if isinstance(value, str):
        try:  # Pillow reads ASCII tags as latin-1; most writers put UTF-8 in them
            return value.encode("latin-1").decode("utf-8")
        except UnicodeError:
            return value
    if isinstance(value, bytes) and tag in _EXIF_XP:
        return value.decode("utf-16-le", "ignore").rstrip("\0")
    if isinstance(value, bytes) and tag in _EXIF_CODED and len(value) >= 8:
        body = value[8:]
        if value.startswith(b"UNICODE"):  # UTF-16 in the byte order of the file it came from
            return body.decode("utf-16-be" if body[:1] == b"\0" else "utf-16-le", "ignore").rstrip("\0")
        return body.decode("latin-1").rstrip("\0")
    return None


def _edit_exif(tag: int, value, edit, rules: list):
    """An EXIF value obfuscated in its own encoding -- an ASCII one with X for
    blacked-out characters, which it cannot hold -- or None for a binary value
    with a match in it, which cannot be obfuscated and goes."""
    text = _exif_text(tag, value)
    if tag in _EXIF_TIMESTAMPS:
        return value
    if text is None:
        if isinstance(value, bytes) and _match_regions(value.decode("latin-1"), rules):
            return None
        return value
    if isinstance(value, str):
        edited = edit(text, "X")  # matched in the decoded text
        return value if edited == text else edited
    edited = edit(text)
    if edited == text:
        return value
    if tag in _EXIF_XP:
        return edited.encode("utf-16-le") + b"\0\0"
    return b"UNICODE\0" + edited.encode("utf-16-be")  # Pillow writes EXIF big-endian


def _save(frames: list, output_path: str, target: str, params: dict) -> None:
    """Save the frames; into a format that cannot hold their mode (written
    as another format than read), as RGB -- or RGBA where that can be kept."""
    try:
        frames[0].save(output_path, target, **params)
    except (OSError, ValueError) as error:
        if "mode" not in str(error):
            raise
        converted = [
            frame.convert("RGBA" if "A" in frame.mode or "transparency" in frame.info else "RGB") for frame in frames
        ]
        try:
            converted[0].save(output_path, target, **{**params, "append_images": converted[1:]})
        except (OSError, ValueError):
            converted = [frame.convert("RGB") for frame in converted]
            converted[0].save(output_path, target, **{**params, "append_images": converted[1:]})


def _image_leftovers(path: str, rules: list, placeholders: list) -> list:
    """Where a saved image's metadata still holds a match."""
    with _any_size(), Image.open(path) as image:
        original = _Original(image, path)
        leftovers = []
        if any(_leftover_matches(text, rules, placeholders) for text in original.texts()):
            leftovers.append("metadata")
        if original.xmp and _match_regions(unescape(original.xmp), rules):
            leftovers.append("XMP metadata")
        return leftovers
