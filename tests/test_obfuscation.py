"""Obfuscation of PDFs and text files (`acd.obfuscation`).

Every PDF here is built on the fly with PyMuPDF, so each test shows the text
it starts from. Whether a string is gone is checked with pypdf as well: a
redaction that only fooled the library that made it would be no redaction.
"""

import importlib.util
import re

import numpy as np
import pymupdf
import pytest
from pypdf import PdfReader

from acd import obfuscation
from acd.obfuscation import BLACKOUT_CHAR
from acd.obfuscation import ObfuscationError
from acd.obfuscation import obfuscate_document
from acd.obfuscation import obfuscate_text


def make_pdf(path, lines, fontsize=10, leading=14, fontname="helv"):
    """A one-page PDF with `lines` written top-down from (40, 50)."""
    doc = pymupdf.open()
    page = doc.new_page(width=420, height=300)
    for number, line in enumerate(lines):
        page.insert_text((40, 50 + number * leading), line, fontsize=fontsize, fontname=fontname)
    doc.save(path)
    return str(path)


def text_of(path):
    """Everything a reader can extract, as PyMuPDF and as pypdf see it."""
    with pymupdf.open(path) as doc:
        mupdf_text = "".join(page.get_text(clip=pymupdf.INFINITE_RECT()) for page in doc)
    return mupdf_text + "\n" + "".join(page.extract_text() for page in PdfReader(path).pages)


def test_placeholders_replace_strings_consistently(tmp_path):
    source = make_pdf(tmp_path / "bom.pdf", [
        "Part Number : 1641B0100-01",
        "Cage Code : D9893",
        "Next assy 1641B0200-01 from D9893",
    ])

    output, mapping = obfuscate_document(source, {
        "D9893": "placeholder",
        re.compile(r"\b\d{4}B\d{4}-\d{2}\b"): "placeholder",
    })

    assert output == str(tmp_path / "bom_obfuscated.pdf")
    assert mapping == {"1641B0100-01": "PH-0001", "D9893": "PH-0002", "1641B0200-01": "PH-0003"}
    text = text_of(output)
    assert not any(original in text for original in mapping)
    assert text.count("PH-0002") == 4  # twice on the page, as each library sees it
    assert "Part Number :" in text and "Next assy" in text
    assert "D9893" in text_of(source)  # the original is left alone


def test_blackout_deletes_the_text_under_a_black_box(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["Created by Heike Wochner today"])

    output, mapping = obfuscate_document(source, {"Heike Wochner": "blackout"})

    assert mapping == {"Heike Wochner": "█████ ███████"}
    text = text_of(output)
    assert "Heike" not in text and "Wochner" not in text
    assert "Created by" in text and "today" in text
    with pymupdf.open(source) as doc:
        name = doc[0].search_for("Heike Wochner")[0]
    with pymupdf.open(output) as doc:
        pixmap = doc[0].get_pixmap(clip=name)
    assert pixmap.pixel(pixmap.width // 2, pixmap.height // 2) == (0, 0, 0)


@pytest.mark.parametrize("fontname", ["helv", "tiro", "cour"])
@pytest.mark.parametrize("kind", ["blackout", "placeholder"])
def test_tightly_set_lines_above_and_below_survive(tmp_path, kind, fontname):
    # 10 pt text on 8 pt leading: the boxes text extraction gives each glyph
    # reach into the lines above and below, and MuPDF deletes every glyph
    # whose box a redaction touches
    source = make_pdf(
        tmp_path / "tight.pdf",
        ["first line stays Apg", "Heike Wochner", "middle line stays", "HEIKE WOCHNER", "last line jjj stays"],
        leading=8, fontname=fontname,
    )

    output, _ = obfuscate_document(source, {"Heike Wochner": kind, "HEIKE WOCHNER": kind})

    with pymupdf.open(output) as doc:
        text = doc[0].get_text()
    assert "first line stays Apg" in text
    assert "middle line stays" in text
    assert "last line jjj stays" in text
    assert "Heike" not in text and "HEIKE" not in text


def test_neighbours_on_the_line_survive(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["PMC-LIAEREP30-D9893-05448-01 X(D9893)Y"])

    output, _ = obfuscate_document(source, {"D9893": "blackout"})

    with pymupdf.open(output) as doc:
        assert "".join(doc[0].get_text().split()) == "PMC-LIAEREP30--05448-01X()Y"


def test_literal_matches_across_line_breaks_and_missing_spaces(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((40, 50), "owner Heike", fontsize=10)
    page.insert_text((40, 62), "Wochner (2024)", fontsize=10)
    page.insert_text((40, 90), "HeikeWochner", fontsize=10)
    doc.save(tmp_path / "doc.pdf")

    output, mapping = obfuscate_document(str(tmp_path / "doc.pdf"), {"Heike Wochner": "placeholder"})

    assert mapping == {"Heike Wochner": "PH-0001"}  # one string, however it is spaced
    text = text_of(output)
    assert "Heike" not in text and "Wochner" not in text
    assert "owner" in text and "(2024)" in text
    assert text.count("PH-0001") == 4  # twice on the page, as each library sees it


def test_overlapping_matches_merge_and_blackout_wins(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["by Heike Wochner", "for Heike alone"])

    output, mapping = obfuscate_document(source, {"Heike": "placeholder", "Heike Wochner": "blackout"})

    assert mapping == {"Heike Wochner": "█████ ███████", "Heike": "PH-0001"}
    assert "Wochner" not in text_of(output)


def test_a_string_matched_once_is_obfuscated_everywhere(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["BaaN ID : 13922397", "Order 13922397", "Appendix_LLI_DCN-114380"])

    output, mapping = obfuscate_document(source, {
        re.compile(r"(?<=BaaN ID : )\d+"): "placeholder",  # finds the value only after its label
        "LLI_DCN-114380": "blackout",
    })

    assert mapping == {"13922397": "PH-0001", "LLI_DCN-114380": "██████████████"}
    text = text_of(output)
    assert "BaaN ID :" in text and "Order" in text and "Appendix_" in text
    assert "13922397" not in text and "114380" not in text
    assert text.count("PH-0001") == 4  # twice on the page, as each library sees it


def test_word_boundaries_do_not_hide_a_value_found_elsewhere(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["DCN LLI_DCN-114380", "See Appendix_LLI_DCN-114380"])

    output, mapping = obfuscate_document(source, {re.compile(r"\bLLI_DCN-\d+\b"): "placeholder"})

    assert mapping == {"LLI_DCN-114380": "PH-0001"}
    assert "114380" not in text_of(output)


def test_known_placeholders_carry_over_between_documents(tmp_path):
    first = make_pdf(tmp_path / "first.pdf", ["Cage D9893, assy 1641B0100-01"])
    second = make_pdf(tmp_path / "second.pdf", ["Assy 1641B0200-01 of 1641B0100-01, cage D9893"])
    targets = {"D9893": "placeholder", re.compile(r"\d{4}B\d{4}-\d{2}"): "placeholder"}

    _, first_mapping = obfuscate_document(first, targets, placeholders={"D9893": "CAGE1"})
    _, second_mapping = obfuscate_document(second, targets, placeholders=first_mapping)

    assert first_mapping == {"D9893": "CAGE1", "1641B0100-01": "PH-0001"}
    assert second_mapping == {"1641B0200-01": "PH-0002", "1641B0100-01": "PH-0001", "D9893": "CAGE1"}


def test_new_placeholders_avoid_text_already_in_the_document(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["See PH-0001 and PH-0002", "Cage D9893"])

    _, mapping = obfuscate_document(source, {"D9893": "placeholder"})

    assert mapping == {"D9893": "PH-0003"}


def test_annotations_forms_links_bookmarks_and_metadata(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((40, 40), "Body by Heike Wochner", fontsize=10)
    page.insert_text((40, 300), "mail heike@example.com", fontsize=10)
    note = page.add_freetext_annot(pymupdf.Rect(40, 60, 300, 90), "Checked by Heike Wochner", fontsize=10)
    note.set_info(title="Heike Wochner", subject="D9893")
    note.update()
    widget = pymupdf.Widget()
    widget.field_type = pymupdf.PDF_WIDGET_TYPE_TEXT
    widget.field_name = "cage"
    widget.field_value = "D9893"
    widget.rect = pymupdf.Rect(40, 100, 200, 120)
    page.add_widget(widget)
    page.insert_link({"kind": pymupdf.LINK_URI, "from": pymupdf.Rect(40, 290, 200, 305), "uri": "mailto:heike@example.com"})
    doc.set_toc([[1, "Report for D9893", 1]])
    doc.set_metadata({"author": "Heike Wochner", "title": "D9893 report", "creationDate": "D:20241205103000"})
    doc.set_xml_metadata("<x:xmpmeta xmlns:x='adobe:ns:meta/'><dc:creator>Heike Wochner</dc:creator></x:xmpmeta>")
    doc.save(tmp_path / "doc.pdf")

    output, mapping = obfuscate_document(str(tmp_path / "doc.pdf"), {
        "Heike Wochner": "blackout", "D9893": "placeholder", "heike@example.com": "placeholder",
    })

    assert mapping == {"Heike Wochner": "█████ ███████", "heike@example.com": "PH-0001", "D9893": "PH-0002"}
    with pymupdf.open(output) as doc:
        page = doc[0]
        note = next(page.annots())
        assert note.info["content"] == "Checked by █████ ███████"
        assert (note.info["title"], note.info["subject"]) == ("█████ ███████", "PH-0002")
        assert next(page.widgets()).field_value == "PH-0002"
        assert page.get_links() == []
        assert doc.get_toc() == [[1, "Report for PH-0002", 1]]
        assert doc.metadata["author"] == "█████ ███████"
        assert doc.metadata["title"] == "PH-0002 report"
        assert doc.metadata["creationDate"] == "D:20241205103000"
        assert doc.get_xml_metadata() == ""


def test_annotation_still_showing_a_match_is_deleted(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    stamp = page.add_freetext_annot(pymupdf.Rect(40, 60, 300, 90), "Approved D9893", fontsize=10)
    stamp.update()
    doc.xref_set_key(stamp.xref, "Contents", pymupdf.get_pdf_str("Approved"))  # appearance unchanged
    doc.save(tmp_path / "doc.pdf")

    output, _ = obfuscate_document(str(tmp_path / "doc.pdf"), {"D9893": "placeholder"})

    with pymupdf.open(output) as doc:
        assert list(doc[0].annots()) == []


def every_object(path):
    """All objects of a PDF, streams decompressed: whatever a reader of the
    file could get at."""
    with pymupdf.open(path) as doc:
        parts = []
        for xref in range(1, doc.xref_length()):
            parts.append(doc.xref_object(xref).encode("latin-1", "replace"))
            if doc.xref_is_stream(xref):
                parts.append(doc.xref_stream(xref))
    return b"\n".join(parts)


def test_attached_files_are_removed(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((40, 50), "Cage D9893", fontsize=10)
    doc.embfile_add("notes.txt", b"embedded: Heike Wochner")
    page.add_file_annot((40, 100), b"annotation: Heike Wochner", "annot.txt")
    stream = doc.get_new_xref()  # a PDF/A-3 associated file
    doc.update_object(stream, "<</Type/EmbeddedFile>>")
    doc.update_stream(stream, b"associated: Heike Wochner")
    spec = doc.get_new_xref()
    doc.update_object(spec, f"<</Type/Filespec/F(assoc.txt)/AFRelationship/Data/EF<</F {stream} 0 R>>>>")
    doc.xref_set_key(doc.pdf_catalog(), "AF", f"[{spec} 0 R]")
    doc.save(tmp_path / "doc.pdf", garbage=3)  # keeps only what is referenced: all three
    assert every_object(tmp_path / "doc.pdf").count(b"Heike Wochner") == 3

    output, _ = obfuscate_document(str(tmp_path / "doc.pdf"), {"D9893": "placeholder"})

    assert b"Heike Wochner" not in every_object(output)


def test_digital_signatures_are_removed(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((40, 50), "Approved D9893", fontsize=10)
    widget = pymupdf.Widget()
    widget.field_type = pymupdf.PDF_WIDGET_TYPE_SIGNATURE
    widget.field_name = "Signature1"
    widget.rect = pymupdf.Rect(40, 100, 240, 140)
    page.add_widget(widget)
    signature = doc.get_new_xref()  # the signer's certificate lives in /Contents
    doc.update_object(signature, f"<</Type/Sig/Name(Gregor Paulmann)/Contents<{b'CN=Gregor Paulmann'.hex()}>>>")
    doc.xref_set_key(next(page.widgets()).xref, "V", f"{signature} 0 R")
    doc.save(tmp_path / "doc.pdf")

    output, _ = obfuscate_document(str(tmp_path / "doc.pdf"), {"D9893": "placeholder"})

    with pymupdf.open(output) as doc:
        assert list(doc[0].widgets()) == []
    assert b"Paulmann" not in every_object(output)


def test_text_outside_the_page_is_removed_too(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page(width=300, height=200)
    page.insert_text((40, 50), "visible", fontsize=10)
    page.insert_text((40, 400), "hidden D9893 below the page", fontsize=10)
    doc.save(tmp_path / "doc.pdf")

    output, mapping = obfuscate_document(str(tmp_path / "doc.pdf"), {"D9893": "blackout"})

    assert mapping == {"D9893": "█████"}
    assert "D9893" not in text_of(output)


def test_placeholders_follow_rotated_text(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page(width=400, height=300)
    page.insert_text((50, 50), "flat D9893 end", fontsize=11)
    page.insert_text((300, 250), "up D9893 end", fontsize=11, rotate=90)
    page.insert_text((350, 100), "flip D9893 end", fontsize=11, rotate=180)
    page.set_rotation(90)
    doc.save(tmp_path / "doc.pdf")

    output, _ = obfuscate_document(str(tmp_path / "doc.pdf"), {"D9893": "placeholder"})

    def lines(path):
        with pymupdf.open(path) as doc:
            return [
                ("".join(span["text"] for span in line["spans"]).strip(), line)
                for block in doc[0].get_text("dict")["blocks"] for line in block["lines"]
            ]

    placeholders = [line for text, line in lines(output) if text == "PH-0001"]
    for text, original in lines(tmp_path / "doc.pdf"):
        # written in the same direction, where the original was
        assert any(
            line["dir"] == original["dir"] and pymupdf.Rect(line["bbox"]).intersects(original["bbox"])
            for line in placeholders
        ), text
    assert len(placeholders) == 3
    assert "D9893" not in text_of(output)


@pytest.mark.parametrize("font, flags, expected", [
    ("Arial,Bold", 20, "hebo"),  # Crystal Reports flags Arial as serifed
    ("TimesNewRomanPS-ItalicMT", 0, "tiit"),
    ("CourierNewPSMT", 0, "cour"),
    ("ABCDEE+CorporateText-Regular", 4, "tiro"),  # nothing in the name: the flags decide
    ("ABCDEE+CorporateText-Regular", 0, "helv"),
])
def test_placeholder_font_follows_the_original(font, flags, expected):
    assert obfuscation._base14_font({"font": font, "flags": flags}) == expected


def test_blackout_images(tmp_path):
    doc = pymupdf.open()
    page = doc.new_page(width=300, height=200)
    logo = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 20), False)
    logo.set_rect(logo.irect, (0, 60, 140))
    page.insert_image(pymupdf.Rect(100, 50, 200, 100), pixmap=logo)
    doc.save(tmp_path / "doc.pdf")

    output, mapping = obfuscate_document(str(tmp_path / "doc.pdf"), {}, blackout_images=True)

    assert mapping == {}
    with pymupdf.open(output) as doc:
        pixmap = doc[0].get_pixmap(clip=pymupdf.Rect(100, 50, 200, 100))
        assert {pixmap.pixel(x, y) for x in (5, 50, 90) for y in (5, 30, 45)} == {(0, 0, 0)}
        # not just painted over: no pixel of the logo is left in the file
        for xref, *_ in doc[0].get_images():
            image = pymupdf.Pixmap(doc, xref)
            assert (0, 60, 140) not in {image.pixel(x, y) for x in range(image.width) for y in range(image.height)}


def test_obfuscated_copy_is_checked(tmp_path, monkeypatch):
    source = make_pdf(tmp_path / "doc.pdf", ["Cage D9893"])
    monkeypatch.setattr(obfuscation, "_redact_page", lambda *args: None)  # a redaction that fails silently

    with pytest.raises(ObfuscationError, match="page 1"):
        obfuscate_document(source, {"D9893": "placeholder"})
    assert not (tmp_path / "doc_obfuscated.pdf").exists()


def test_xml_file(tmp_path):
    source = tmp_path / "DMC-TEST.XML"
    source.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<dmodule><pmCode pmIssuer="D9893"/><para>Approved by Heike Wochner for D9893</para></dmodule>\n',
        encoding="utf-8",
    )

    output, mapping = obfuscate_document(source, {"D9893": "placeholder", "Heike Wochner": "blackout"})

    assert output == str(tmp_path / "DMC-TEST_obfuscated.XML")
    assert mapping == {"D9893": "PH-0001", "Heike Wochner": "█████ ███████"}
    assert (tmp_path / "DMC-TEST_obfuscated.XML").read_text(encoding="utf-8") == (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<dmodule><pmCode pmIssuer="PH-0001"/><para>Approved by █████ ███████ for PH-0001</para></dmodule>\n'
    )


def test_text_file_keeps_encoding_and_line_endings(tmp_path):
    source = tmp_path / "notes.txt"
    source.write_bytes("Prüfer: Heike Wochner\r\nCage: D9893\r\n".encode("cp1252"))

    output, mapping = obfuscate_document(str(source), {"Heike Wochner": "blackout", "D9893": "placeholder"})

    assert mapping == {"Heike Wochner": f"{BLACKOUT_CHAR * 5} {BLACKOUT_CHAR * 7}", "D9893": "PH-0001"}
    # cp1252 has no block character: X stands in for it
    assert (tmp_path / "notes_obfuscated.txt").read_bytes() == "Prüfer: XXXXX XXXXXXX\r\nCage: PH-0001\r\n".encode("cp1252")


def test_obfuscate_text_replaces_in_a_str():
    text, mapping = obfuscate_text(
        "Prüfer: Heike\nWochner, Cage D9893, PN 1641B0100-01 (D9893)",
        {"Heike Wochner": "blackout", "D9893": "placeholder", re.compile(r"\b\d{4}B\d{4}-\d{2}\b"): "placeholder"},
    )

    assert text == "Prüfer: █████\n███████, Cage PH-0001, PN PH-0002 (PH-0001)"
    assert mapping == {"Heike Wochner": "█████ ███████", "D9893": "PH-0001", "1641B0100-01": "PH-0002"}


def test_obfuscate_text_reuses_placeholders_and_blackout_char():
    _, first = obfuscate_text("Cage D9893", {"D9893": "placeholder"})
    text, mapping = obfuscate_text(
        "PH-0001 and D9893 by Heike Wochner",
        {"D9893": "placeholder", "Heike Wochner": "blackout"},
        placeholders=first,
        blackout_char="X",
    )

    # PH-0001 is D9893's from the first text, even though it occurs here already
    assert text == "PH-0001 and PH-0001 by XXXXX XXXXXXX"
    assert mapping == {"D9893": "PH-0001", "Heike Wochner": f"{BLACKOUT_CHAR * 5} {BLACKOUT_CHAR * 7}"}


def test_obfuscate_text_rejects_bad_arguments():
    with pytest.raises(ValueError):
        obfuscate_text("D9893", {"D9893": "redact"})
    with pytest.raises(ValueError):
        obfuscate_text("D9893", {"D9893": "placeholder"}, placeholder_format="PH")
    with pytest.raises(TypeError):
        obfuscate_text("D9893", {9893: "placeholder"})


needs_ocr = pytest.mark.skipif(
    importlib.util.find_spec("rapidocr") is None, reason="OCR needs rapidocr: pip install rapidocr onnxruntime"
)


def make_scan(path, lines, dpi=300):
    """A one-page PDF that is only an image: `lines` of (point, text,
    fontsize, rotate) rendered to pixels."""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=420)
    for point, text, fontsize, rotate in lines:
        page.insert_text(point, text, fontsize=fontsize, rotate=rotate)
    pixmap = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
    scan = pymupdf.open()
    scan.new_page(width=595, height=420).insert_image(pymupdf.Rect(0, 0, 595, 420), pixmap=pixmap)
    scan.save(path)
    return str(path)


def read_back(path):
    """What OCR reads on the first page of a PDF."""
    from acd.line_ocr import ocr_lines
    with pymupdf.open(path) as doc:
        pixmap = doc[0].get_pixmap(dpi=200, colorspace=pymupdf.csGRAY)
    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width)
    return [line.text for line in ocr_lines(image)]


@needs_ocr
def test_ocr_obfuscates_a_scan(tmp_path):
    source = make_scan(tmp_path / "scan.pdf", [
        ((60, 80), "Created by Heike Wochner on 16.11.99", 14, 0),
        ((60, 130), "Part 1644-0035 cage D9893 Werkstoff", 14, 0),
        ((60, 180), "keep this line as it is", 14, 0),
        ((40, 380), "16440035 MARGIN", 12, 90),  # bottom to top, as on drawings
    ])

    output, mapping = obfuscate_document(source, {
        "Heike Wochner": "blackout",
        "D9893": "placeholder",
        re.compile(r"\b1644-?\d{4}\b"): "placeholder",
    }, ocr=True)

    assert mapping == {
        "Heike Wochner": "█████ ███████", "1644-0035": "PH-0001", "D9893": "PH-0002", "16440035": "PH-0003",
    }
    read = " | ".join(read_back(output))
    for original in mapping:
        assert original not in read, read
    for kept in ("Created by", "Part PH-0001", "cage PH-0002", "keep this line as it is", "PH-0003 MARGIN"):
        assert kept in read, read
    with pymupdf.open(source) as before, pymupdf.open(output) as after:
        name = pymupdf.Rect(130, 66, 260, 84)  # where "Heike Wochner" was
        assert after[0].get_pixmap(clip=name).pixel(30, 9) == (0, 0, 0)
        assert before[0].get_pixmap(clip=name).pixel(30, 9) != (0, 0, 0)


@needs_ocr
def test_ocr_reads_text_in_images_on_text_pages(tmp_path):
    logo = pymupdf.open()
    logo_page = logo.new_page(width=200, height=40)
    logo_page.insert_text((10, 28), "LIEBHERR", fontsize=24, fontname="hebo")
    pixmap = logo_page.get_pixmap(dpi=300)
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=420)
    page.insert_image(pymupdf.Rect(380, 20, 580, 60), pixmap=pixmap)
    page.insert_text((60, 120), "Liebherr-Aerospace Lindenberg GmbH", fontsize=11)
    doc.save(tmp_path / "doc.pdf")

    output, mapping = obfuscate_document(str(tmp_path / "doc.pdf"), {re.compile("liebherr", re.I): "placeholder"}, ocr=True)

    assert mapping == {"Liebherr": "PH-0001", "LIEBHERR": "PH-0002"}
    with pymupdf.open(output) as doc:
        assert "Liebherr" not in doc[0].get_text()
    assert "LIEBHERR" not in " ".join(read_back(output))


def test_ocr_matching_allows_for_misreadings():
    rules = obfuscation._compile_targets({"Klippel": "blackout", re.compile(r"\b\d{4}-\d{4}\b"): "placeholder"})
    found = [("Drawing 1644-0035, checked by Klippel", obfuscation._match_regions(
        "Drawing 1644-0035, checked by Klippel", rules))]
    misreadings = obfuscation._misreadings(obfuscation._known_strings(rules, found))

    def originals(text):
        return [(text[start:end], original) for start, end, _, original in
                obfuscation._ocr_regions(text, rules, misreadings)]

    assert originals("Bearb.16.11.99KLippeι") == [("KLippeι", "Klippel")]  # case, Greek iota
    assert originals("Gepr. Klippe") == [("Klippe", "Klippel")]  # a letter dropped
    assert originals("Zeichnungsnummer 1644-O035") == [("1644-O035", "1644-0035")]  # O for 0
    assert originals("Idnr．：1644－0035") == [("1644－0035", "1644-0035")]  # full-width forms
    assert originals("Klebt am Rand, Kapitel 7") == []  # nothing like it
    # read as a pattern match, a sibling part number is not taken for a misread one
    assert originals("Next assy 1644-0036") == [("1644-0036", "1644-0036")]


def test_ocr_misreadings_of_a_number_are_not_taken_inside_a_longer_one():
    rules = obfuscation._compile_targets({"13923284": "placeholder"})
    misreadings = obfuscation._misreadings(obfuscation._known_strings(rules, []))

    def originals(text):
        return [(text[start:end], original) for start, end, _, original in
                obfuscation._ocr_regions(text, rules, misreadings)]

    assert originals("Order 13923284 of") == [("13923284", "13923284")]
    assert originals("Order 13928284 of") == [("13928284", "13923284")]  # one digit misread
    assert originals("Ref 913928284") == []  # a misread could be anywhere in a longer number
    assert originals("Ref 13928284 7") == [("13928284", "13923284")]  # a number next to it is no part of it


def test_ocr_matches_keep_the_known_string_they_contain():
    # a pattern match inside a misread known string must not give it a placeholder of its own
    rules = obfuscation._compile_targets({
        "Liebherr-Aerospace Lindenberg GmbH": "placeholder", re.compile("liebherr", re.I): "placeholder",
    })
    misreadings = obfuscation._misreadings(obfuscation._known_strings(rules, []))

    regions = obfuscation._ocr_regions("LIEBHERR-AEROSPACE LINDÈNBERG GmbH", rules, misreadings)

    assert [(start, end, original) for start, end, _, original in regions] == [
        (0, 34, "Liebherr-Aerospace Lindenberg GmbH"),
    ]


@pytest.mark.parametrize("targets, error", [
    ({"D9893": "redact"}, ValueError),
    ({"  ": "blackout"}, ValueError),
    ({42: "blackout"}, TypeError),
])
def test_invalid_targets(tmp_path, targets, error):
    source = make_pdf(tmp_path / "doc.pdf", ["Cage D9893"])
    with pytest.raises(error):
        obfuscate_document(source, targets)


def test_invalid_arguments(tmp_path):
    source = make_pdf(tmp_path / "doc.pdf", ["Cage D9893"])
    (tmp_path / "doc.docx").write_bytes(b"")
    with pytest.raises(ValueError, match="Cannot obfuscate .docx"):
        obfuscate_document(str(tmp_path / "doc.docx"), {"D9893": "blackout"})
    with pytest.raises(ValueError, match="never overwritten"):
        obfuscate_document(source, {"D9893": "blackout"}, output_path=source)
    with pytest.raises(ValueError, match="running number"):
        obfuscate_document(source, {"D9893": "placeholder"}, placeholder_format="XXX")


def test_long_placeholder_shrinks_to_fit_before_the_next_word(tmp_path):
    # an 18-character project placeholder for a 6-character code, the next word close behind
    source = make_pdf(tmp_path / "doc.pdf", ["code D9893A end"])
    output, mapping = obfuscate_document(str(source), {"D9893A": "placeholder"}, placeholder_format="PH-E3B1Q-{:09d}")
    assert mapping == {"D9893A": "PH-E3B1Q-000000001"}
    with pymupdf.open(source) as doc:
        code = next(pymupdf.Rect(word[:4]) for word in doc[0].get_text("words") if word[4] == "D9893A")
    with pymupdf.open(output) as doc:
        page = doc[0]
        words = {word[4]: pymupdf.Rect(word[:4]) for word in page.get_text("words")}
        sizes = {span["text"].strip(): span["size"] for block in page.get_text("dict")["blocks"]
                 for line in block["lines"] for span in line["spans"]}
    assert sizes["PH-E3B1Q-000000001"] < 5  # below half the original size: it had to, to fit
    assert words["PH-E3B1Q-000000001"].x1 <= code.x1 + 0.5  # within the code's place: not over the space and "end"
