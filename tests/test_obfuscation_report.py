"""Tests for acd.obfuscation_report: the original vs obfuscated HTML report."""

import pytest

pymupdf = pytest.importorskip("pymupdf")
pytest.importorskip("cv2")

from acd.obfuscation_report import build_report


def two_page_pdf(path, name):
    doc = pymupdf.open()
    for text in (f"Inspected by {name}", "Nothing to hide on this page"):
        page = doc.new_page()
        page.insert_text((72, 100), text, fontsize=14)
    doc.save(path)


def test_report_pairs_pages_and_boxes_changes(tmp_path):
    originals, obfuscated, report = tmp_path / "in", tmp_path / "out", tmp_path / "report"
    originals.mkdir()
    obfuscated.mkdir()
    two_page_pdf(originals / "sheet.pdf", "Heike Wochner")
    two_page_pdf(obfuscated / "sheet_obfuscated.pdf", "█████ ███████")
    two_page_pdf(obfuscated / "orphan_obfuscated.pdf", "x")  # no original: left out
    summary = {"sheet.pdf": {"mapping": {"Heike Wochner": "█" * 5 + " " + "█" * 7},
                             "logos": [{"page": 1, "box": [10, 10, 50, 20], "score": 0.91, "reference": "logo.png"}]}}
    path = build_report(originals, obfuscated, report, summary, progress=lambda line: None)
    page = path.read_text(encoding="utf-8")
    assert "d_sheet" in page and "orphan" not in page
    assert "change(s)</span>" in page
    assert "unchanged" in page  # page 2
    assert "Heike Wochner" in page and "1 logo(s) removed" in page
    images = sorted(p.name for p in (report / "img").iterdir())
    assert "d_sheet_p1_original.jpg" in images and "d_sheet_p2_obfuscated.jpg" in images
    assert "d_sheet_p1_obfuscated_boxed.jpg" in images
