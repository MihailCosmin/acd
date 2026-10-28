r"""Obfuscate sensitive strings in documents.

`obfuscate_document` writes a copy of a document in which every match of the
given targets is blacked out or replaced by a placeholder, and returns where
the copy is and what was replaced:

    from re import compile
    output, mapping = obfuscate_document("BOM.pdf", {
        "Heike Wochner": "blackout",
        "D9893": "placeholder",
        compile(r"\b\d{4}B\d{4}-\d{2}\b"): "placeholder",
    })
    # output  == "BOM_obfuscated.pdf"
    # mapping == {"D9893": "PH-0001", "1641B0100-01": "PH-0002",
    #             "Heike Wochner": "█████ ███████"}

A placeholder stands for one string throughout the document, so the result
stays readable and `mapping` can turn placeholders back into the originals.

PDF pages are redacted, not painted over: MuPDF deletes the matched characters
from the page content, so they cannot be copied, searched or extracted from
the result. Annotations, form fields, links, bookmarks and the document
metadata are obfuscated too, and the saved file is read back to make sure no
match survived. What cannot be obfuscated is removed: attached files, and
digital signatures (which carry the signer's certificate, and no longer fit
the changed document anyway). Text formats (`TEXT_EXTENSIONS`, e.g. S1000D
XML) are rewritten with every match replaced in place, and `obfuscate_text`
does the same to a str.

Text that only exists as pixels -- scanned pages, logos -- or as vector
outlines is read with OCR when `ocr=True` (see `line_ocr`), and obfuscated
where it was read: its pixels are blanked under a black box or the
placeholder. OCR can only obfuscate what it can read, so its matching allows
for the usual misreadings of strings known to be sensitive. Image files --
JPEG, PNG, TIFF, WebP, AVIF and the rest -- are obfuscated the same way, with
their metadata: see `image_obfuscation`.
"""

from collections import namedtuple

from html import unescape

from math import atan2
from math import degrees

from os import fspath
from os import remove
from os.path import abspath
from os.path import normcase
from os.path import splitext

from typing import Optional
from typing import Tuple

from numpy import array
from numpy import frombuffer
from numpy import uint8

from regex import BESTMATCH
from regex import compile as regex_compile
from regex import escape as regex_escape

from pymupdf import INFINITE_RECT
from pymupdf import PDF_ANNOT_FILE_ATTACHMENT
from pymupdf import PDF_ANNOT_FREE_TEXT
from pymupdf import PDF_ANNOT_IS_HIDDEN
from pymupdf import PDF_REDACT_IMAGE_NONE
from pymupdf import PDF_REDACT_IMAGE_PIXELS
from pymupdf import PDF_REDACT_IMAGE_REMOVE
from pymupdf import PDF_REDACT_LINE_ART_NONE
from pymupdf import PDF_REDACT_LINE_ART_REMOVE_IF_COVERED
from pymupdf import PDF_REDACT_TEXT_NONE
from pymupdf import PDF_REDACT_TEXT_REMOVE
from pymupdf import Matrix
from pymupdf import Point
from pymupdf import Rect
from pymupdf import TEXT_FONT_BOLD
from pymupdf import TEXT_FONT_ITALIC
from pymupdf import TEXT_FONT_MONOSPACED
from pymupdf import TEXT_FONT_SERIFED
from pymupdf import TEXT_PRESERVE_WHITESPACE
from pymupdf import csGRAY
from pymupdf import get_pdf_str
from pymupdf import get_text_length
from pymupdf import open as pdf_open
from pymupdf import sRGB_to_pdf

from .filepath import clean_path


OBFUSCATION_TYPES = ("blackout", "placeholder")

# A blacked-out string maps to one block per character (whitespace kept), and
# that is also what replaces it in a text file.
BLACKOUT_CHAR = "\u2588"

# New placeholders are PH-0001, PH-0002, ... -- zero-padded, so that no
# placeholder is the start of another one (PH-1 / PH-10) when mapping back.
PLACEHOLDER_FORMAT = "PH-{:04d}"

TEXT_EXTENSIONS = (
    ".txt", ".xml", ".sgm", ".sgml", ".htm", ".html", ".xhtml", ".svg",
    ".csv", ".tsv", ".json", ".md", ".log", ".ini", ".cfg", ".yaml", ".yml",
)

# Text extraction for matching: whitespace as written, ligatures split into
# their letters (so "fi" matches "ﬁ"), and -- by passing no clip flag -- text
# placed outside the visible page as well, since that is extractable too.
_PDF_TEXT_FLAGS = TEXT_PRESERVE_WHITESPACE

# Document information entries that are timestamps, not text: a digit pattern
# would only mangle them.
_PDF_INFO_SKIPPED = ("CreationDate", "ModDate", "Trapped")

# Base-14 font for a placeholder, by the style of the text it replaces:
# (regular, italic, bold, bold italic)
_BASE14_SANS = ("helv", "heit", "hebo", "hebi")
_BASE14_SERIF = ("tiro", "tiit", "tibo", "tibi")
_BASE14_MONO = ("cour", "coit", "cobo", "cobi")

# Font names that tell the style better than the font's flags do (Crystal
# Reports, for one, flags Arial as serifed). Checked in this order.
_MONO_NAMES = ("courier", "mono", "consol")
_SANS_NAMES = (
    "sans", "arial", "helvetica", "calibri", "verdana", "tahoma", "segoe", "frutiger", "univers", "futura",
    "gill", "myriad", "lucida", "trebuchet",
)
_SERIF_NAMES = ("serif", "times", "roman", "garamond", "georgia", "cambria", "palatino", "bookman", "antiqua")

# A placeholder too long for the space of the original shrinks to fit, but
# not below this font size in points (it overflows instead): project-wide
# placeholders ("PH-E3B1Q-000000001") are long, and running over the
# neighbouring text makes both unreadable.
_MIN_PLACEHOLDER_SIZE = 3.0

# A target, ready to match: a compiled pattern, or a literal (without its
# whitespace) that matches whatever the whitespace in the text. `key` is how
# a str target appears in the mapping: as the caller wrote it.
_Rule = namedtuple("_Rule", "kind pattern literal key")

# OCR reads pixels at this resolution: drawing lettering (2.5 mm and up) comes
# out about 20 pixels high, which PP-OCRv6 reads well, and an A0 sheet stays
# at 62 megapixels.
_OCR_DPI = 200

# An image smaller than this, in points, either way holds no readable text;
# images covering more than this share of a page are read as the whole page.
_OCR_MIN_IMAGE = 12
_OCR_FULL_PAGE = 0.5

# The font size of OCR'd text is told by its characters' pitch, which is
# about this share of it; how high OCR boxes a line varies too much (1.3 to
# 2.2 times the font size).
_OCR_PITCH = 0.55

# Characters OCR reads for Latin letters and ASCII punctuation: Greek and
# Cyrillic look-alikes, full-width forms and dashes.
_LOOKALIKES = str.maketrans({
    **dict(zip("ΑΒΕΖΗΙΚΜΝΟΡΤΥΧαβγεικνορτυχ", "ABEZHIKMNOPTYXaßyeikvoptux")),
    **dict(zip("АВЕКМНОРСТУХаеорсухіјѕІЈЅ", "ABEKMHOPCTYXaeopcyxijsIJS")),
    **{chr(code): chr(code - 0xFEE0) for code in range(0xFF01, 0xFF5F)},
    **dict.fromkeys("‐‑‒–—―−", "-"),
})

# Characters OCR mistakes for one another, each set brought to one of them
# (after lower-casing) when strings known to be sensitive are looked for.
_CONFUSABLE = str.maketrans({"o": "0", "i": "1", "l": "1", "|": "1", "!": "1", "s": "5", "b": "8", "z": "2"})


class ObfuscationError(Exception):
    """Raised when a matched string can still be read in an obfuscated PDF.
    The PDF is deleted before this is raised."""


def obfuscate_document(
    filepath: str,
    targets: dict,
    output_path: Optional[str] = None,
    placeholders: Optional[dict] = None,
    blackout_images: bool = False,
    ocr: bool = False,
    placeholder_format: str = PLACEHOLDER_FORMAT,
) -> Tuple[str, dict]:
    r"""Write an obfuscated copy of a document.

    Args:
        filepath (str): a PDF; an image -- JPEG, PNG, TIFF, WebP, AVIF, GIF,
            BMP and every other format Pillow reads (`image_obfuscation`),
            always read with OCR; or a text file (`TEXT_EXTENSIONS`)
        targets (dict): what to obfuscate -> "blackout" or "placeholder".
            A str key is matched literally and case-sensitively, but whatever
            the whitespace in and around its characters: PDF text breaks
            names over lines, spaces headings out ("I n s p e c t i o n") and
            runs words together. A compiled `re` (or `regex`) pattern is
            matched as is -- use re.IGNORECASE for case-insensitive matching,
            and lookarounds to obfuscate only part of a match:
            `re.compile(r"(?<=Part Number : )\S+")`. PDF text is matched page
            by page, one line per line, as `Page.get_text()` returns it. A
            string a target matches anywhere is obfuscated everywhere, also
            where the target itself would not match it. Where matches
            overlap, their union is obfuscated, and blackout wins.
        output_path (str, optional): where to write the copy, in the format
            its extension names. Defaults to `<name>_obfuscated<ext>` beside
            the input -- `.png` for an image format Pillow cannot write; the
            input itself is never overwritten.
        placeholders (dict, optional): placeholder to use for a given string
            (original -> placeholder). Pass the mapping returned for another
            document to give its strings the same placeholders here.
        blackout_images (bool, optional): PDF only -- also black out every
            image (logos, photos, scanned pages). Defaults to False.
        ocr (bool, optional): PDF only -- also read the text of scanned
            pages, images and vector-outlined lettering with OCR, and
            obfuscate the targets found there (image files are read with OCR
            regardless). Needs `pip install rapidocr onnxruntime`; takes a
            few seconds per scanned page (A0 sheets about a minute). Text OCR
            cannot read stays: check the result. Defaults to False.
        placeholder_format (str, optional): format of new placeholders,
            filled with a running number. Defaults to `PLACEHOLDER_FORMAT`.

    Returns:
        tuple: the path of the obfuscated copy, and {original: replacement}
            for every string obfuscated in this document. The replacement is
            the placeholder, or for a blacked-out string one BLACKOUT_CHAR per
            character. An original matched by a str target is that target;
            any other has its whitespace collapsed to single spaces. The file
            name is not obfuscated: pass output_path if it gives too much
            away.

    Raises:
        ValueError: unsupported file type, unknown obfuscation type, empty
            target, output_path equal to filepath, or an encrypted PDF
        TypeError: a target that is neither a str nor a compiled pattern
        ObfuscationError: a matched string survived in the PDF or the
            image's metadata (the file is deleted)
    """
    filepath = fspath(filepath)  # a pathlib.Path will do too
    extension = splitext(filepath)[1].lower()
    image = extension != ".pdf" and extension not in TEXT_EXTENSIONS
    if image:
        from . import image_obfuscation  # imports this module: imported here, not at the top

        if extension not in image_obfuscation.image_extensions():
            raise ValueError(
                f"Cannot obfuscate {extension or 'extensionless'} files: supported are .pdf, "
                f"images ({', '.join(sorted(image_obfuscation.image_extensions()))}) and {', '.join(TEXT_EXTENSIONS)}"
            )
    rules = _compile_targets(targets)
    _check_placeholder_format(placeholder_format)

    if output_path is None:
        stem, suffix = splitext(filepath)
        output_path = image_obfuscation.default_image_output(filepath) if image else stem + "_obfuscated" + suffix
    output_path = fspath(output_path)
    if normcase(abspath(output_path)) == normcase(abspath(filepath)):
        raise ValueError("output_path must differ from filepath: the original is never overwritten")

    if extension == ".pdf":
        mapping = _obfuscate_pdf(filepath, output_path, rules, placeholders, placeholder_format, blackout_images, ocr)
    elif image:
        mapping = image_obfuscation._obfuscate_image(filepath, output_path, rules, placeholders, placeholder_format)
    else:
        mapping = _obfuscate_text_file(filepath, output_path, rules, placeholders, placeholder_format)
    return output_path, mapping


def obfuscate_text(
    text: str,
    targets: dict,
    placeholders: Optional[dict] = None,
    blackout_char: str = BLACKOUT_CHAR,
    placeholder_format: str = PLACEHOLDER_FORMAT,
) -> Tuple[str, dict]:
    r"""Obfuscate a string: every match of the targets replaced in place.

        text, mapping = obfuscate_text("Approved by Heike Wochner for D9893", {
            "Heike Wochner": "blackout",
            "D9893": "placeholder",
        })
        # text    == "Approved by █████ ███████ for PH-0001"
        # mapping == {"D9893": "PH-0001", "Heike Wochner": "█████ ███████"}

    Args:
        text (str): the text to obfuscate
        targets (dict): what to obfuscate -> "blackout" or "placeholder",
            matched as `obfuscate_document` matches them
        placeholders (dict, optional): placeholder to use for a given string
            (original -> placeholder). Pass the mapping returned for another
            text or document to give its strings the same placeholders here.
        blackout_char (str, optional): the character that replaces each
            blacked-out character (whitespace is kept). Defaults to
            BLACKOUT_CHAR; the mapping uses BLACKOUT_CHAR regardless.
        placeholder_format (str, optional): format of new placeholders,
            filled with a running number. Defaults to `PLACEHOLDER_FORMAT`.

    Returns:
        tuple: the obfuscated text, and {original: replacement} for every
            string obfuscated in it, as `obfuscate_document` returns it

    Raises:
        ValueError: unknown obfuscation type, empty target, or a
            placeholder_format without a field for the running number
        TypeError: a target that is neither a str nor a compiled pattern
    """
    rules = _compile_targets(targets)
    _check_placeholder_format(placeholder_format)
    return _obfuscate_str(text, rules, placeholders, placeholder_format, blackout_char)


def _check_placeholder_format(placeholder_format: str) -> None:
    if placeholder_format.format(1) == placeholder_format.format(2):
        raise ValueError(f"placeholder_format {placeholder_format!r} needs a field for the running number")


def _compile_targets(targets: dict) -> list:
    """Turn the targets dict into a list of `_Rule`."""
    rules = []
    for target, kind in targets.items():
        kind = str(kind).strip().lower()
        if kind not in OBFUSCATION_TYPES:
            raise ValueError(f"Unknown obfuscation type {kind!r} for {target!r}: use one of {OBFUSCATION_TYPES}")
        if isinstance(target, str):
            if not target.strip():
                raise ValueError("Empty obfuscation target")
            rules.append(_Rule(kind, None, _identity(target), _mapping_key(target)))
        elif hasattr(target, "finditer"):  # re.Pattern or regex.Pattern
            rules.append(_Rule(kind, target, None, None))
        else:
            raise TypeError(f"Obfuscation target must be a str or a compiled pattern, not {type(target).__name__}")
    return rules


def _match_regions(text: str, rules: list) -> list:
    """Every match of every rule in `text` as [(start, end, type)], in text
    order, without surrounding whitespace. Overlapping matches merge into one
    region, which is blacked out if any of them is."""
    hits = []
    compact = None
    for rule in rules:
        if rule.pattern is not None:
            spans = [match.span() for match in rule.pattern.finditer(text)]
        else:  # a literal: looked for in the text without its whitespace
            if compact is None:
                positions = [index for index, char in enumerate(text) if not char.isspace()]
                compact = "".join(text[index] for index in positions)
            spans = []
            found = compact.find(rule.literal)
            while found >= 0:
                spans.append((positions[found], positions[found + len(rule.literal) - 1] + 1))
                found = compact.find(rule.literal, found + 1)
        for start, end in spans:
            while start < end and text[start].isspace():
                start += 1
            while end > start and text[end - 1].isspace():
                end -= 1
            if start < end:
                hits.append((start, end, rule.kind))
    hits.sort()
    regions = []
    for start, end, kind in hits:
        if regions and start < regions[-1][1]:
            last_start, last_end, last_kind = regions[-1]
            merged_kind = "blackout" if "blackout" in (kind, last_kind) else "placeholder"
            regions[-1] = (last_start, max(last_end, end), merged_kind)
        else:
            regions.append((start, end, kind))
    return regions


def _every_occurrence(rules: list, texts: list) -> list:
    """`rules`, plus rules that match each string they match in `texts`
    wherever else it occurs: a pattern with a lookbehind or a word boundary
    finds a value in one place, but where it does not match -- as
    r"\\bLLI_DCN-\\d+" does not in "Appendix_LLI_DCN-114380" -- the value is
    just as sensitive."""
    kinds = _kinds_by_identity([(text, _match_regions(text, rules)) for text in texts])
    literals = {rule.literal for rule in rules}
    return rules + [
        _Rule(kind, None, identity, None) for identity, kind in kinds.items() if identity not in literals
    ]


def _mapping_key(original: str) -> str:
    """How a matched string appears in the mapping: whitespace collapsed."""
    return " ".join(original.split())


def _identity(original: str) -> str:
    """What makes two matches the same string: their characters, whatever the
    whitespace -- PDF text extraction finds a line break between two words in
    one place, a space in another and nothing at all in a third."""
    return "".join(original.split())


def _blackout_of(original: str) -> str:
    return "".join(char if char.isspace() else BLACKOUT_CHAR for char in original)


def _is_blackout(replacement: Optional[str]) -> bool:
    """Whether a replacement blacks out: block characters, maybe spaced."""
    if not replacement or BLACKOUT_CHAR not in replacement:
        return False
    return not replacement.replace(BLACKOUT_CHAR, "").strip()


def _kinds_by_identity(sources: list) -> dict:
    """Obfuscation type per matched string over all (text, regions) sources
    of a document: blackout if any match of it is blacked out."""
    kinds = {}
    for text, regions in sources:
        for start, end, kind in regions:
            identity = _identity(text[start:end])
            if kinds.get(identity) != "blackout":
                kinds[identity] = kind
    return kinds


class _Replacements:
    """The replacement of every string matched in one document.

    A string is blacked out everywhere if any of its matches is (or if
    `known` says it was blacked out before), else it gets its `known`
    placeholder or the next free one -- free meaning not handed out yet and
    not occurring anywhere in the document, so mapping back is unambiguous.
    """

    def __init__(
        self, rules: list, kinds: dict, known: Optional[dict], placeholder_format: str, document_text: str
    ):
        self.kinds = kinds
        self.known = {_identity(original): value for original, value in (known or {}).items()}
        self.placeholder_format = placeholder_format
        self.document_text = document_text
        self.taken = set(self.known.values())
        self.number = 0
        # identity -> key in the mapping: a str target as given, else the
        # string as first found
        self.keys = {rule.literal: rule.key for rule in rules if rule.key}
        self.mapping = {}

    def __getitem__(self, original: str) -> str:
        identity = _identity(original)
        key = self.keys.setdefault(identity, _mapping_key(original))
        if key not in self.mapping:
            known = self.known.get(identity)
            if self.kinds[identity] == "blackout" or _is_blackout(known):
                self.mapping[key] = _blackout_of(key)
            elif known is not None:
                self.mapping[key] = known
            else:
                self.mapping[key] = self._next_placeholder()
        return self.mapping[key]

    def _next_placeholder(self) -> str:
        while True:
            self.number += 1
            placeholder = self.placeholder_format.format(self.number)
            if placeholder not in self.taken and placeholder not in self.document_text:
                self.taken.add(placeholder)
                return placeholder

    def placeholders(self) -> list:
        """The placeholders handed out, longest first."""
        return sorted((value for value in self.mapping.values() if not _is_blackout(value)), key=len, reverse=True)


def _replace_regions(
    text: str, regions: list, replacements: _Replacements, blackout_char: str = BLACKOUT_CHAR
) -> str:
    """`text` with every region replaced: by its placeholder, or each of its
    characters except whitespace by `blackout_char`."""
    parts = []
    position = 0
    for start, end, _ in regions:
        original = text[start:end]
        replacement = replacements[original]
        if _is_blackout(replacement):
            replacement = "".join(char if char.isspace() else blackout_char for char in original)
        parts.append(text[position:start])
        parts.append(replacement)
        position = end
    parts.append(text[position:])
    return "".join(parts)


# -- Text files -------------------------------------------------------------

def _decode_text(data: bytes) -> Tuple[str, str]:
    """Decode a text file as UTF-8 (with or without BOM) or UTF-16 if it is,
    else as cp1252, else as latin-1 (which decodes anything). Returns the text
    and the encoding to write it back with."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16"), "utf-16"
    for encoding in ("utf-8", "cp1252"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1"), "latin-1"


def _obfuscate_text_file(
    filepath: str, output_path: str, rules: list, placeholders: Optional[dict], placeholder_format: str
) -> dict:
    """Obfuscate a text file: every match replaced in place, the encoding and
    line endings kept. Blacked-out characters become BLACKOUT_CHAR, or X where
    the file's encoding has no such character."""
    with open(clean_path(filepath), "rb") as file_:
        text, encoding = _decode_text(file_.read())

    try:
        BLACKOUT_CHAR.encode(encoding)
        blackout_char = BLACKOUT_CHAR
    except UnicodeEncodeError:
        blackout_char = "X"
    text, mapping = _obfuscate_str(text, rules, placeholders, placeholder_format, blackout_char)

    with open(clean_path(output_path), "wb") as file_:
        file_.write(text.encode(encoding))
    return mapping


def _obfuscate_str(
    text: str, rules: list, placeholders: Optional[dict], placeholder_format: str, blackout_char: str
) -> Tuple[str, dict]:
    """`text` with every match replaced in place, and the mapping."""
    rules = _every_occurrence(rules, [text])
    regions = _match_regions(text, rules)
    kinds = _kinds_by_identity([(text, regions)])
    replacements = _Replacements(rules, kinds, placeholders, placeholder_format, text)
    return _replace_regions(text, regions, replacements, blackout_char), replacements.mapping


# -- PDF ----------------------------------------------------------------------

def _pdf_page_text(page) -> Tuple[str, list]:
    """The text of a page as one string, each line followed by a newline (the
    shape `Page.get_text()` returns), and for every character of it the
    (line, span, char) dicts it came from -- None for the newlines."""
    text = []
    chars = []
    for block in page.get_text("rawdict", flags=_PDF_TEXT_FLAGS, clip=INFINITE_RECT())["blocks"]:
        for line in block.get("lines", ()):
            for span in line["spans"]:
                for char in span["chars"]:
                    text.append(char["c"])
                    chars.extend([(line, span, char)] * len(char["c"]))
            text.append("\n")
            chars.append(None)
    return "".join(text), chars


def _pdf_content_text(doc, page_number: int) -> Tuple[str, list]:
    """`_pdf_page_text` of the page content alone. Text extraction includes
    what annotations and form fields show, which redacting the page does not
    remove (they are obfuscated through their own fields instead), so they are
    hidden while the text is extracted."""
    page = doc[page_number]
    flags = {xref: doc.xref_get_key(xref, "F") for xref, *_ in page.annot_xrefs()}
    for xref, (kind, value) in flags.items():
        doc.xref_set_key(xref, "F", str((int(value) if kind == "int" else 0) | PDF_ANNOT_IS_HIDDEN))
    try:
        return _pdf_page_text(page)
    finally:
        for xref, (_, value) in flags.items():
            doc.xref_set_key(xref, "F", value)  # "null" removes the key again


def _char_box(entry: tuple) -> tuple:
    """The (x0, y0, x1, y1) box a character visibly takes up: one font size
    high across its line. (Text extraction makes it as high as the font's
    ascender plus descender, which in tightly set text reaches well into the
    neighbouring lines.)"""
    line, span, char = entry
    x0, y0, x1, y1 = char["bbox"]
    ascender, descender = span["ascender"], span["descender"]
    if ascender <= descender:
        return x0, y0, x1, y1
    above = span["size"] * ascender / (ascender - descender)  # baseline to top
    below = above - span["size"]  # baseline to bottom, negative
    x, y = char["origin"]
    cos, sin = line["dir"]
    if abs(sin) < 1e-3:  # horizontal, glyph tops pointing up (or down)
        y0, y1 = (y - above, y - below) if cos > 0 else (y + below, y + above)
    elif abs(cos) < 1e-3:  # vertical, glyph tops pointing left (or right)
        x0, x1 = (x - above, x - below) if sin < 0 else (x + below, x + above)
    return x0, y0, x1, y1


def _union(boxes) -> tuple:
    boxes = list(boxes)
    return (
        min(box[0] for box in boxes), min(box[1] for box in boxes),
        max(box[2] for box in boxes), max(box[3] for box in boxes),
    )


def _bounds(points) -> tuple:
    return _union((point.x, point.y, point.x, point.y) for point in points)


def _transpose(box: tuple) -> tuple:
    return box[1], box[0], box[3], box[2]


def _line_segments(chars: list) -> list:
    """The characters of a region grouped by text line, in reading order,
    leaving out any without extent (a zero-width space, say)."""
    segments = []
    for entry in chars:
        if entry is None:
            continue
        x0, y0, x1, y1 = entry[2]["bbox"]
        if x0 >= x1 or y0 >= y1:
            continue
        if segments and segments[-1][0][0] is entry[0]:
            segments[-1].append(entry)
        else:
            segments.append([entry])
    return segments


def _kept_boxes(chars: list, regions: list) -> list:
    """Extraction boxes of the characters a page keeps, by line: [(line box,
    [character boxes])]."""
    matched = set()
    for start, end, _ in regions:
        matched.update(range(start, end))
    lines = {}
    for index, entry in enumerate(chars):
        if entry is not None and index not in matched:
            lines.setdefault(id(entry[0]), []).append(tuple(entry[2]["bbox"]))
    return [(_union(boxes), boxes) for boxes in lines.values()]


def _removal_rect(segment: list, kept: list) -> Rect:
    """The redaction rect that deletes the characters of one line of a match,
    and no others.

    MuPDF deletes every character whose extraction box, shrunk by a tenth on
    each side, a redaction overlaps -- and the extraction boxes of tightly set
    lines overlap. So the rect spans the shrunk boxes of the segment along the
    line, and across it keeps to the widest band that is clear of the (whole)
    boxes of every character that stays.
    """
    cos, sin = segment[0][0]["dir"]
    frame = _transpose if abs(cos) < abs(sin) else tuple  # along the line is x
    targets = []
    for _, _, char in segment:
        x0, y0, x1, y1 = char["bbox"]
        dx, dy = (x1 - x0) / 10, (y1 - y0) / 10
        targets.append(frame((x0 + dx, y0 + dy, x1 - dx, y1 - dy)))
    start, end = min(box[0] for box in targets), max(box[2] for box in targets)
    low, high = max(box[1] for box in targets), min(box[3] for box in targets)
    if low >= high:  # different sizes or baselines: no band crosses them all
        low, high = min(box[1] for box in targets), max(box[3] for box in targets)
    free = [(low, high)]
    for line_box, boxes in kept:
        line_box = frame(line_box)
        if line_box[0] >= end or line_box[2] <= start or line_box[1] >= high or line_box[3] <= low:
            continue
        for box in map(frame, boxes):
            if box[0] < end and box[2] > start:  # cut the box (plus a hair) out of the free bands
                free = [
                    part
                    for band_low, band_high in free
                    for part in (
                        (band_low, min(band_high, box[1] - 0.01)),
                        (max(band_low, box[3] + 0.01), band_high),
                    )
                    if part[0] < part[1]
                ]
    if free:  # else the text overlaps text that stays: it cannot go alone
        low, high = max(free, key=lambda band: band[1] - band[0])
    return Rect(frame((start, low, end, high)))


def _obstacles(page, chars: list) -> list:
    """What a placeholder must not run into, as boxes: the characters of a
    page (the deleted ones too, as their replacements go there), images, and
    every edge of its line art -- a table cell's border, say."""
    obstacles = [_char_box(entry) for entry in chars if entry is not None and not entry[2]["c"].isspace()]
    obstacles += [tuple(info["bbox"]) for info in page.get_image_info()]
    for path in page.get_drawings():
        for item in path["items"]:
            if item[0] == "re":
                corners = (item[1].tl, item[1].tr, item[1].br, item[1].bl)
            elif item[0] == "qu":
                corners = (item[1].ul, item[1].ur, item[1].lr, item[1].ll)
            else:  # "l" line or "c" curve: no edges to tell apart
                obstacles.append(_bounds(item[1:]))
                continue
            obstacles += [_bounds((corner, corners[i - 1])) for i, corner in enumerate(corners)]
    return obstacles


def _base14_font(span: dict) -> str:
    """The Base-14 font closest in style to a span's: going by the name of
    its font where that tells, else by its flags."""
    name = span["font"].lower()
    flags = span["flags"]
    if any(word in name for word in _MONO_NAMES):
        family = _BASE14_MONO
    elif any(word in name for word in _SANS_NAMES):
        family = _BASE14_SANS
    elif any(word in name for word in _SERIF_NAMES):
        family = _BASE14_SERIF
    elif flags & TEXT_FONT_MONOSPACED:
        family = _BASE14_MONO
    elif flags & TEXT_FONT_SERIFED:
        family = _BASE14_SERIF
    else:
        family = _BASE14_SANS
    bold = flags & TEXT_FONT_BOLD or any(word in name for word in ("bold", "black", "heavy"))
    italic = flags & TEXT_FONT_ITALIC or any(word in name for word in ("italic", "oblique"))
    return family[(2 if bold else 0) + (1 if italic else 0)]


def _insert_placeholder(shape, rect: Rect, entry: tuple, placeholder: str, obstacles) -> None:
    """Write a placeholder where the text it replaces began: same baseline,
    direction, size and colour, in the closest Base-14 font. If it is longer
    than the original, it runs on into free space (horizontal text only) and
    shrinks as far as it then still has to, to `_MIN_PLACEHOLDER_SIZE`. `obstacles` returns the
    boxes it must not run into."""
    line, span, char = entry
    fontname = _base14_font(span)
    fontsize = span["size"]
    cos, sin = line["dir"]
    room = rect.width if abs(cos) >= abs(sin) else rect.height
    width = get_text_length(placeholder, fontname=fontname, fontsize=fontsize)
    if width > room and cos > 0 and abs(sin) < 1e-3:
        stop = shape.page.cropbox.x1
        for x0, y0, _, y1 in obstacles():
            if rect.x1 - 0.01 <= x0 < stop and y0 < rect.y1 and y1 > rect.y0:
                stop = x0
        room = max(room, stop - rect.x0 - 0.25 * fontsize)  # a space's width clear of it
    if width > room > 0:
        fontsize = max(fontsize * room / width, min(fontsize, _MIN_PLACEHOLDER_SIZE))
    # line["dir"] is in page space (y down); insert_text turns counterclockwise
    angle = round(degrees(atan2(-sin, cos)), 3) % 360
    origin = Point(char["origin"])
    turn = {"rotate": int(angle)} if angle % 90 == 0 else {"morph": (origin, Matrix(angle))}
    shape.insert_text(
        origin, placeholder, fontname=fontname, fontsize=fontsize, color=sRGB_to_pdf(span["color"]), **turn
    )


# -- OCR ------------------------------------------------------------------------

def _pdf_ocr_lines(doc, page_number: int, has_text: bool, blackout_images: bool) -> list:
    """OCR lines (`line_ocr.OcrLine`, mapping into page coordinates) of what a
    page shows only as pixels or outlines: all of a page without text, else
    its images (logos, pasted scans) -- but not what blackout_images blacks
    out anyway. OCR reads a copy without text and annotations, which are
    obfuscated as text."""
    from .line_ocr import ocr_lines  # needs RapidOCR, which only OCR does

    with pdf_open() as shadow:
        shadow.insert_pdf(doc, from_page=page_number, to_page=page_number, links=False, annots=False, widgets=False)
        page = shadow[0]
        page.set_rotation(0)  # read in the coordinates of the page's text
        if not has_text:
            areas = [page.rect]
            if blackout_images and page.get_images():
                # The images go black anyway; the rest of the page (outlined
                # lettering, line art) still has to be read
                page.add_redact_annot(page.rect)
                page.apply_redactions(
                    images=PDF_REDACT_IMAGE_REMOVE, graphics=PDF_REDACT_LINE_ART_NONE, text=PDF_REDACT_TEXT_NONE
                )
                if not page.get_drawings():  # a plain scan: nothing left to read
                    areas = []
        else:
            areas = [] if blackout_images else _image_areas(page)
            if areas:
                page.add_redact_annot(page.rect)
                page.apply_redactions(
                    images=PDF_REDACT_IMAGE_NONE, graphics=PDF_REDACT_LINE_ART_NONE, text=PDF_REDACT_TEXT_REMOVE
                )
        zoom = _OCR_DPI / 72
        lines = []
        for area in areas:
            pixmap = page.get_pixmap(matrix=Matrix(zoom, zoom), clip=area, colorspace=csGRAY, annots=False)
            gray = frombuffer(pixmap.samples, dtype=uint8).reshape(pixmap.height, pixmap.width)
            to_page = array([[1 / zoom, 0, pixmap.x / zoom], [0, 1 / zoom, pixmap.y / zoom], [0, 0, 1]])
            lines += [line._replace(to_image=to_page @ line.to_image) for line in ocr_lines(gray)]
    return lines


def _image_areas(page) -> list:
    """The parts of a page its images cover, joined where they overlap -- the
    whole page if they cover most of it -- leaving out images too small to
    hold text."""
    rects = [Rect(info["bbox"]) & page.rect for info in page.get_image_info()]
    rects = [rect for rect in rects if rect.width >= _OCR_MIN_IMAGE and rect.height >= _OCR_MIN_IMAGE]
    if sum(abs(rect) for rect in rects) > _OCR_FULL_PAGE * abs(page.rect):
        return [page.rect]
    areas = []
    for rect in rects:
        overlapping = [area for area in areas if area.intersects(rect)]
        while overlapping:
            for area in overlapping:
                areas.remove(area)
                rect |= area
            overlapping = [area for area in areas if area.intersects(rect)]
        areas.append(rect)
    return areas


def _skeleton(text: str) -> Tuple[str, list]:
    """`text` reduced to what OCR tells apart reliably -- look-alikes made
    Latin, lower case, the letters it confuses with digits made digits,
    whitespace dropped -- and the index in `text` of each character left."""
    letters, indexes = [], []
    for index, char in enumerate(text):
        if not char.isspace():
            letters.append(char.translate(_LOOKALIKES).lower()[:1].translate(_CONFUSABLE))
            indexes.append(index)
    return "".join(letters), indexes


def _known_strings(rules: list, found: list) -> dict:
    """identity -> (type, original) of the strings known to be sensitive: the
    str targets, and whatever the rules match in `found` ([(text, regions)])."""
    known = {rule.literal: (rule.kind, rule.key) for rule in rules if rule.key}
    for text, regions in found:
        for start, end, kind in regions:
            identity = _identity(text[start:end])
            known_kind, original = known.get(identity, (kind, text[start:end]))
            known[identity] = ("blackout" if "blackout" in (kind, known_kind) else "placeholder", original)
    return known


def _misreadings(known: dict) -> list:
    """The known strings, ready to be looked for in OCR text: [(skeleton, its
    trigrams, how many of them a misreading keeps, fuzzy pattern -- None to
    match exactly --, type, original)]. In 6 characters and more, one may be
    misread, dropped or added, in 12 and more two; under 4 they are too short
    to tell from chance and left to the rules."""
    entries = []
    for identity, (kind, original) in known.items():
        skeleton = _skeleton(identity)[0]
        if len(skeleton) < 4:
            continue
        edits = 0 if len(skeleton) < 6 else 1 if len(skeleton) < 12 else 2
        trigrams = {skeleton[index:index + 3] for index in range(len(skeleton) - 2)}
        pattern = regex_compile(f"(?:{regex_escape(skeleton)}){{e<={edits}}}", BESTMATCH) if edits else None
        entries.append((skeleton, trigrams, len(trigrams) - 3 * edits, pattern, kind, original))
    return entries


def _misread_hits(text: str, misreadings: list) -> list:
    """Where the known strings occur in a line of OCR text, misreadings
    allowed: [(start, end, type, original)]."""
    skeleton, indexes = _skeleton(text)
    trigrams = {skeleton[index:index + 3] for index in range(len(skeleton) - 2)}
    hits = []
    for target, target_trigrams, needed, pattern, kind, original in misreadings:
        if pattern is None:
            spans = []
            start = skeleton.find(target)
            while start >= 0:
                spans.append((start, start + len(target)))
                start = skeleton.find(target, start + 1)
        elif len(target_trigrams & trigrams) >= needed:  # a cheap test before the slow one
            spans = [match.span() for match in pattern.finditer(skeleton)]
        else:
            continue
        spans = [(indexes[start], indexes[end - 1] + 1) for start, end in spans if end > start]
        if target.isdigit():  # not inside a longer number: 13923284 is no part of 0139232845
            spans = [(start, end) for start, end in spans
                     if not text[start - 1:start].isdigit() and not text[end:end + 1].isdigit()]
        hits += [(start, end, kind, original) for start, end in spans]
    return hits


def _ocr_regions(text: str, rules: list, misreadings: list) -> list:
    """The matches in a line of OCR text -- of the rules, in the text as read
    and with look-alikes made Latin, and of the strings known to be sensitive,
    misreadings allowed -- as [(start, end, type, original)]. Overlapping ones
    merge into one region, blacked out if any of them is, and standing for the
    original of the match that spans all of it, else for its own text."""
    plain = text.translate(_LOOKALIKES)
    hits = [(start, end, kind, text[start:end]) for start, end, kind in _match_regions(text, rules)]
    if plain != text:
        hits += [(start, end, kind, plain[start:end]) for start, end, kind in _match_regions(plain, rules)]
    # Last among equals: what OCR read and the rules match is taken as read --
    # 1644-0036 is not a misread 1644-0035.
    hits += _misread_hits(text, misreadings)
    hits.sort(key=lambda hit: (hit[0], -hit[1]))  # at each start the longest first; stable otherwise
    regions = []
    for start, end, kind, original in hits:
        if regions and start < regions[-1][1]:
            last_start, last_end, last_kind, last_original = regions[-1]
            if end > last_end:  # runs on past what the region's original spans
                last_end, last_original = end, plain[last_start:end]
            merged_kind = "blackout" if "blackout" in (kind, last_kind) else "placeholder"
            regions[-1] = (last_start, last_end, merged_kind, last_original)
        else:
            regions.append((start, end, kind, original))
    return regions


def _insert_ocr_placeholder(shape, corners, length: int, placeholder: str) -> None:
    """Write a placeholder where OCR read the `length` characters it replaces
    -- `corners`: top-left, top-right, bottom-right, bottom-left of them the
    right way up -- in Helvetica (all OCR tells of the font), at their size,
    shrunk if longer, and centred across their line."""
    top_left, top_right, _, bottom_left = (Point(*corner) for corner in corners)
    along, down = top_right - top_left, bottom_left - top_left
    fontsize = min(abs(along) / max(length, 1) / _OCR_PITCH, abs(down))
    width = get_text_length(placeholder, fontname="helv", fontsize=fontsize)
    if width > abs(along) > 0:
        fontsize = max(fontsize * abs(along) / width, min(fontsize, _MIN_PLACEHOLDER_SIZE))
    origin = top_left + down * 0.5 + down * (0.35 * fontsize / max(abs(down), 1e-6))  # cap height centred
    angle = degrees(atan2(-along.y, along.x)) % 360
    square = round(angle / 90) * 90 % 360
    if min(abs(angle - square), 360 - abs(angle - square)) < 3:  # a scan's skew of a degree or two
        turn = {"rotate": square}
    else:
        turn = {"morph": (origin, Matrix(angle))}
    shape.insert_text(origin, placeholder, fontname="helv", fontsize=fontsize, color=(0, 0, 0), **turn)


def _redact_page(
    doc, page_number: int, rules: list, replacements: _Replacements, blackout_images: bool,
    ocr_lines: list = (), misreadings: list = (),
) -> None:
    """Redact every match in a page's content: the characters are deleted,
    then a black box or the placeholder is drawn in their place. Matches in
    the page's OCR lines get their pixels blanked instead, under a black box,
    or a white one with the placeholder."""
    text, chars = _pdf_content_text(doc, page_number)
    regions = _match_regions(text, rules)
    page = doc[page_number]
    image_rects = [Rect(info["bbox"]) for info in page.get_image_info()] if blackout_images else []
    read = [(line, region) for line in ocr_lines for region in _ocr_regions(line.text, rules, misreadings)]
    if not regions and not image_rects and not read:
        return
    kept = _kept_boxes(chars, regions)
    removals = []
    covers = [(rect, (0, 0, 0)) for rect in image_rects]  # (rect, fill)
    to_write = []
    for start, end, _ in regions:
        replacement = replacements[text[start:end]]
        for index, segment in enumerate(_line_segments(chars[start:end])):
            cover = Rect(_union(map(_char_box, segment)))
            removals.append(_removal_rect(segment, kept))
            covers.append((cover, (0, 0, 0) if _is_blackout(replacement) else None))
            if index == 0 and not _is_blackout(replacement):
                to_write.append((cover, segment[0], replacement))  # all of it on the first line
    ocr_to_write = []
    for line, (start, end, _, original) in read:
        replacement = replacements[original]
        corners = line.quad(start, end)
        fill = (0, 0, 0) if _is_blackout(replacement) else (1, 1, 1)  # blanked pixels are white
        covers.append((Rect(*corners.min(axis=0), *corners.max(axis=0)), fill))
        if not _is_blackout(replacement):
            ocr_to_write.append((line.quad(start, end, pad=0), end - start, replacement))
    # First the matched characters go, and only they; then, with all text
    # kept, what lies under their boxes: image pixels (the visible text of a
    # scan behind its OCR layer) and line art a box covers completely (text
    # drawn as outlines; table borders run on beyond it and stay).
    if removals:
        for rect in removals:
            page.add_redact_annot(rect)
        page.apply_redactions(
            images=PDF_REDACT_IMAGE_NONE, graphics=PDF_REDACT_LINE_ART_NONE, text=PDF_REDACT_TEXT_REMOVE
        )
    for rect, fill in covers:
        page.add_redact_annot(rect, fill=fill)
    page.apply_redactions(
        images=PDF_REDACT_IMAGE_PIXELS, graphics=PDF_REDACT_LINE_ART_REMOVE_IF_COVERED, text=PDF_REDACT_TEXT_NONE
    )
    if to_write or ocr_to_write:
        found = {}

        def obstacles():  # only looked up if needed: listing line art can be slow
            if "boxes" not in found:
                found["boxes"] = _obstacles(page, chars)
            return found["boxes"]

        shape = page.new_shape()
        for rect, entry, placeholder in to_write:
            _insert_placeholder(shape, rect, entry, placeholder, obstacles)
        for corners, length, placeholder in ocr_to_write:
            _insert_ocr_placeholder(shape, corners, length, placeholder)
        shape.commit()
    doc.xref_set_key(page.xref, "Thumb", "null")  # a thumbnail would still show the original


def _pdf_info(doc) -> Tuple[int, list]:
    """The xref of the document information dictionary and [(key, value)] of
    its text entries -- the standard ones and custom ones (e.g. Company)."""
    kind, value = doc.xref_get_key(-1, "Info")
    if kind != "xref":
        return 0, []
    info_xref = int(value.split()[0])
    entries = []
    for key in doc.xref_get_keys(info_xref):
        kind, value = doc.xref_get_key(info_xref, key)
        if kind == "string" and key not in _PDF_INFO_SKIPPED:
            entries.append((key, value))
    return info_xref, entries


def _pdf_annotation_strings(page) -> list:
    """[(xref, field, value)] of the text fields of a page's annotations and
    form fields; field is ("choice_values", index) for a choice option."""
    entries = []
    for annot in page.annots():
        for field in ("content", "title", "subject"):
            entries.append((annot.xref, field, annot.info.get(field) or ""))
    for widget in page.widgets():
        if isinstance(widget.field_value, str):
            entries.append((widget.xref, "field_value", widget.field_value))
        for index, choice in enumerate(widget.choice_values or ()):
            if isinstance(choice, str):
                entries.append((widget.xref, ("choice_values", index), choice))
    return entries


def _obfuscate_pdf_annotations(doc, page_number: int, rules: list, replacements: _Replacements) -> None:
    """Obfuscate the text fields of a page's annotations and form fields and
    regenerate what they show. Rich text (RC) holding a match is dropped, an
    annotation or form field that still shows a match afterwards (a stamp,
    say) is deleted, and so is a link whose URI contains one. File attachments
    and signed signature fields go regardless."""
    page = doc[page_number]
    for annot in list(page.annots()):
        if annot.type[0] == PDF_ANNOT_FILE_ATTACHMENT:
            page.delete_annot(annot)
    for widget in list(page.widgets()):
        if widget.is_signed:
            page.delete_widget(widget)
    changed = {}
    for xref, field, value in _pdf_annotation_strings(page):
        regions = _match_regions(value, rules)
        if regions:
            changed.setdefault(xref, {})[field] = _replace_regions(value, regions, replacements)
    for annot in page.annots():
        rich_text = doc.xref_get_key(annot.xref, "RC")
        if rich_text[0] == "string" and _match_regions(unescape(rich_text[1]), rules):
            doc.xref_set_key(annot.xref, "RC", "null")
        if annot.xref in changed:
            annot.set_info(**changed[annot.xref])
            if annot.type[0] == PDF_ANNOT_FREE_TEXT:
                annot.update()
    for widget in page.widgets():
        if widget.xref in changed:
            choices = list(widget.choice_values or ())
            for field, value in changed[widget.xref].items():
                if field == "field_value":
                    widget.field_value = value
                else:
                    choices[field[1]] = value
            if choices:
                widget.choice_values = choices
            widget.update()
    placeholders = replacements.placeholders()
    for annot in list(page.annots()):
        if _leftover_matches(annot.get_text(), rules, placeholders):
            page.delete_annot(annot)
    # A widget's text is not to be had on its own; but the page content
    # under it is obfuscated by now, so what still matches there is the
    # widget's.
    for widget in list(page.widgets()):
        if _leftover_matches(page.get_text(clip=widget.rect), rules, placeholders):
            page.delete_widget(widget)
    for link in page.get_links():
        if _match_regions(link.get("uri") or "", rules):
            page.delete_link(link)


def _show_all_layers(doc):
    """Switch every optional-content layer on, in memory only (the saved layer
    settings stay as they are): text extraction skips a layer that is off,
    but its text is in the file all the same."""
    for layer in doc.layer_ui_configs():
        doc.set_layer_ui_config(layer["number"], 0)  # 0: on


def _obfuscate_pdf(
    filepath: str, output_path: str, rules: list, placeholders: Optional[dict], placeholder_format: str,
    blackout_images: bool, ocr: bool,
) -> dict:
    """Obfuscate a PDF: see the module docstring."""
    with pdf_open(filepath) as doc:
        if not doc.is_pdf:
            raise ValueError(f"{filepath} is not a PDF")
        if doc.needs_pass:
            raise ValueError(f"{filepath} is password protected")
        _show_all_layers(doc)

        # Find everything first: whether a string is blacked out, and the
        # numbering of the placeholders, depend on the whole document. (Only
        # the text is kept: the geometry of every character takes some 600
        # bytes, 1.5 GB for 1000 dense pages, so each page's is fetched again
        # as it is redacted.)
        texts = [_pdf_content_text(doc, page_number)[0] for page_number in range(doc.page_count)]
        ocr_lines = [
            _pdf_ocr_lines(doc, page_number, bool(texts[page_number].strip()), blackout_images) if ocr else []
            for page_number in range(doc.page_count)
        ]
        toc = [(item[3]["xref"], item[1]) for item in doc.get_toc(simple=False) if item[3].get("xref")]
        info_xref, info = _pdf_info(doc)
        texts += [value for page in doc for _, _, value in _pdf_annotation_strings(page)]
        texts += [title for _, title in toc] + [value for _, value in info]
        texts += [line.text for lines in ocr_lines for line in lines]
        rules = _every_occurrence(rules, texts)
        found = [(text, _match_regions(text, rules)) for text in texts]
        kinds = _kinds_by_identity(found)
        misreadings = _misreadings(_known_strings(rules, found)) if ocr else []
        for line in (line for lines in ocr_lines for line in lines):  # OCR text: matches allowing for misreadings
            for _, _, kind, original in _ocr_regions(line.text, rules, misreadings):
                if kinds.get(_identity(original)) != "blackout":
                    kinds[_identity(original)] = kind
        replacements = _Replacements(rules, kinds, placeholders, placeholder_format, "\n".join(texts))

        for page_number in range(doc.page_count):
            _redact_page(doc, page_number, rules, replacements, blackout_images, ocr_lines[page_number], misreadings)
            _obfuscate_pdf_annotations(doc, page_number, rules, replacements)
        for xref, title in toc:
            regions = _match_regions(title, rules)
            if regions:
                doc.xref_set_key(xref, "Title", get_pdf_str(_replace_regions(title, regions, replacements)))
        for key, value in info:
            regions = _match_regions(value, rules)
            if regions:
                doc.xref_set_key(info_xref, key, get_pdf_str(_replace_regions(value, regions, replacements)))
        # XMP metadata repeats the information dictionary as XML, plus edit
        # history and custom properties; with a match anywhere in it, it goes.
        if _match_regions(unescape(doc.get_xml_metadata()), rules):
            doc.del_xml_metadata()
        for name in doc.embfile_names():
            doc.embfile_del(name)
        for xref in [doc.pdf_catalog()] + [page.xref for page in doc]:
            if doc.xref_get_key(xref, "AF")[0] != "null":  # PDF/A-3 associated files
                doc.xref_set_key(xref, "AF", "null")
        catalog = doc.pdf_catalog()
        for xref in range(1, doc.xref_length()):
            # XMP of pages and images (the catalog's is handled above)
            kind, value = doc.xref_get_key(xref, "Metadata")
            if xref != catalog and kind == "xref":
                stream = doc.xref_stream(int(value.split()[0])) or b""
                if _match_regions(unescape(stream.decode("utf-8", "replace")), rules):
                    doc.xref_set_key(xref, "Metadata", "null")
        # garbage collection drops the replaced content streams, which still
        # hold the original text
        doc.save(output_path, garbage=3, deflate=True)

    leftovers = _pdf_leftovers(output_path, rules, replacements.placeholders())
    if leftovers:
        remove(output_path)
        raise ObfuscationError(f"Obfuscating {filepath} left matches readable in: {', '.join(leftovers)}")
    return replacements.mapping


def _leftover_matches(text: str, rules: list, placeholders: list) -> int:
    """How many rule matches obfuscated `text` still contains, not counting
    those inside the placeholders that were put there."""
    for placeholder in placeholders:
        text = text.replace(placeholder, "\0" * len(placeholder))
    return len(_match_regions(text, rules))


def _pdf_leftovers(path: str, rules: list, placeholders: list) -> list:
    """Where an obfuscated PDF still contains a match. Page text is read with
    annotations and form fields this time, since a viewer shows them."""
    leftovers = []
    with pdf_open(path) as doc:
        _show_all_layers(doc)
        for page in doc:
            if _leftover_matches(_pdf_page_text(page)[0], rules, placeholders):
                leftovers.append(f"page {page.number + 1}")
            if any(_leftover_matches(value, rules, placeholders) for _, _, value in _pdf_annotation_strings(page)):
                leftovers.append(f"page {page.number + 1} annotations")
            if any(_match_regions(link.get("uri") or "", rules) for link in page.get_links()):
                leftovers.append(f"page {page.number + 1} links")
        if any(_leftover_matches(title, rules, placeholders) for _, title, _ in doc.get_toc()):
            leftovers.append("bookmarks")
        if any(_leftover_matches(value, rules, placeholders) for _, value in _pdf_info(doc)[1]):
            leftovers.append("metadata")
        if _match_regions(unescape(doc.get_xml_metadata()), rules):
            leftovers.append("XMP metadata")
    return leftovers
