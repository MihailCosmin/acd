"""HTML report comparing obfuscated documents with their originals, page by page.

    python -m acd.obfuscation_report <originals folder> <obfuscated folder> <report folder> [--summary summary.json]

Every PDF in the obfuscated folder is paired with the original of the same
name (without the `_obfuscated` suffix). Each page is rendered from both,
side by side; what changed is boxed in red on the original, so a reviewer
sees what was removed and can check that nothing sensitive is left. The
optional summary (written by `python -m acd.obfuscation_ai apply --summary`)
adds the strings replaced and the logos removed per document.

The report is report.html with one tab per document, and the page images in
img/ beside it, loaded when their tab is opened. It shows the original
documents: keep it where the originals may be.
"""

import html
import json
import re
from argparse import ArgumentParser
from pathlib import Path

from numpy import abs as np_abs
from numpy import asarray
from numpy import int16
from numpy import uint8

REPORT_MAX_PIXELS = 1800  # longest side of a page image
REPORT_DIFF_LEVEL = 48  # grey levels a pixel must change by to count as changed
RED = (220, 0, 0)

CSS = """
:root { --bg:#f4f5f7; --panel:#fff; --text:#1d2329; --muted:#5d6670; --line:#d6d9dd; --tab:#e8eaed; --red:#c40000; --ok:#167a2e; }
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) { --bg:#16191d; --panel:#1f2329; --text:#e6e8eb; --muted:#9aa3ad; --line:#353b43; --tab:#2a2f36; --red:#ff6b6b; --ok:#5fd17a; } }
:root[data-theme="dark"] { --bg:#16191d; --panel:#1f2329; --text:#e6e8eb; --muted:#9aa3ad; --line:#353b43; --tab:#2a2f36; --red:#ff6b6b; --ok:#5fd17a; }
* { box-sizing:border-box; } body { margin:0; font:14px/1.45 system-ui,Segoe UI,sans-serif; background:var(--bg); color:var(--text); }
header { padding:16px 24px 8px; } h1 { font-size:20px; margin:0 0 4px; } h2 { font-size:17px; margin:0 0 6px; } .muted { color:var(--muted); }
.tabs { display:flex; flex-wrap:wrap; gap:4px; padding:8px 24px 0; border-bottom:1px solid var(--line); position:sticky; top:0; background:var(--bg); z-index:2; }
.tab { border:1px solid var(--line); border-bottom:none; background:var(--tab); color:var(--text); padding:6px 12px; border-radius:6px 6px 0 0; cursor:pointer; font-size:12.5px; }
.tab.active { background:var(--panel); font-weight:600; } .tab .n { color:var(--muted); font-weight:400; }
.panel { display:none; padding:16px 24px 40px; } .panel.active { display:block; }
.tools { display:flex; gap:16px; align-items:center; flex-wrap:wrap; margin:8px 0 14px; font-size:13px; }
.pair { background:var(--panel); border:1px solid var(--line); border-radius:8px; margin:0 0 18px; padding:12px; }
.pair h3 { margin:0 0 8px; font-size:14px; } .sides { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.side img { width:100%; border:1px solid var(--line); background:#fff; display:block; } .side .label { font-size:12px; color:var(--muted); margin-bottom:4px; }
.side .ov { display:none; }
.panel.overlay .side.after .plain { display:none; } .panel.overlay .side.after .ov { display:block; }
.unchanged { opacity:.75; } .hide-unchanged .unchanged { display:none; }
.badge { font-size:11.5px; border-radius:10px; padding:1px 8px; border:1px solid var(--line); margin-left:6px; font-weight:400; }
.badge.chg { color:var(--red); border-color:var(--red); } .badge.same { color:var(--muted); }
details { margin:6px 0 14px; } summary { cursor:pointer; color:var(--muted); }
table { border-collapse:collapse; margin:6px 0; } td, th { border:1px solid var(--line); padding:4px 10px; text-align:left; vertical-align:top; } th { background:var(--tab); }
td.num { text-align:right; } .warn { color:var(--red); font-weight:600; } .good { color:var(--ok); }
code { font-size:12.5px; } a { color:inherit; }
@media (max-width:800px) { .sides { grid-template-columns:1fr; } header, .panel, .tabs { padding-left:16px; padding-right:16px; } }
"""

JS = """
function show(id){document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active',t.dataset.id===id));
document.querySelectorAll('.panel').forEach(p=>{const on=p.id===id;p.classList.toggle('active',on);
if(on)p.querySelectorAll('img[data-src]').forEach(i=>{i.src=i.dataset.src;i.removeAttribute('data-src');});});
history.replaceState(null,'','#'+id);}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>show(t.dataset.id));
document.querySelectorAll('input[data-act]').forEach(c=>c.onchange=()=>{const p=c.closest('.panel');
p.classList.toggle(c.dataset.act,c.checked);});
show(location.hash.slice(1)||'overview');
"""


def _render(page):
    from pymupdf import Matrix
    from PIL import Image

    zoom = min(1.4, REPORT_MAX_PIXELS / max(page.rect.width, page.rect.height))
    pixmap = page.get_pixmap(matrix=Matrix(zoom, zoom))
    return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def _changed_boxes(before, after) -> list:
    """Boxes (x0, y0, x1, y1) around the pixels that changed between two renders of a page."""
    import cv2

    if before.size != after.size:
        after = after.resize(before.size)
    a = asarray(before.convert("L"), dtype=int16)
    b = asarray(after.convert("L"), dtype=int16)
    mask = (np_abs(a - b) > REPORT_DIFF_LEVEL).astype(uint8)
    mask = cv2.dilate(mask, cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9)))
    count, _, stats, _ = cv2.connectedComponentsWithStats(mask)
    return [(x, y, x + w, y + h) for x, y, w, h, area in stats[1:count] if area >= 30]


def _boxed(image, boxes):
    from PIL import Image
    from PIL import ImageDraw

    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    for x0, y0, x1, y1 in boxes:
        draw.rectangle([x0 - 2, y0 - 2, x1 + 2, y1 + 2], fill=RED + (45,), outline=RED + (255,), width=2)
    return Image.alpha_composite(image.convert("RGBA"), overlay).convert("RGB")


def _original_name(name: str) -> str:
    return re.sub(r"_obfuscated(?=\.[^.]+$)", "", name)


def _stem_id(name: str) -> str:
    """An HTML id for a document: its path (folders included), without the extension."""
    return "d_" + re.sub(r"[^A-Za-z0-9]+", "_", str(Path(name).with_suffix(""))).strip("_")


def compare_document(original: Path, obfuscated: Path, img_dir: Path, stem: str) -> list:
    """[{page, changes, img_a, img_b, img_ov}] for every page of the obfuscated document."""
    from pymupdf import open as pdf_open

    pages = []
    with pdf_open(original) as doc_a, pdf_open(obfuscated) as doc_b:
        for number in range(max(len(doc_a), len(doc_b))):
            entry = {"page": number + 1, "changes": 0}
            image_a = _render(doc_a[number]) if number < len(doc_a) else None
            image_b = _render(doc_b[number]) if number < len(doc_b) else None
            boxes = _changed_boxes(image_a, image_b) if image_a and image_b else []
            entry["changes"] = len(boxes)
            if image_a:
                entry["img_a"] = f"img/{stem}_p{number + 1}_original.jpg"
                _boxed(image_a, boxes).save(img_dir / Path(entry["img_a"]).name, "JPEG", quality=78, optimize=True)
            if image_b:
                entry["img_b"] = f"img/{stem}_p{number + 1}_obfuscated.jpg"
                image_b.save(img_dir / Path(entry["img_b"]).name, "JPEG", quality=78, optimize=True)
            if image_a and image_b:  # the same boxes on the obfuscated page, for the overlay view
                entry["img_ov"] = f"img/{stem}_p{number + 1}_obfuscated_boxed.jpg"
                _boxed(image_b.resize(image_a.size), boxes).save(img_dir / Path(entry["img_ov"]).name, "JPEG",
                                                                 quality=78, optimize=True)
            pages.append(entry)
    return pages


def _side(label, src, overlay=None, after=False):
    if not src:
        return f"<div class='side'><div class='label'>{label}</div><p class='muted'>no such page</p></div>"
    extra = (f"<a class='ov' href='{overlay}' target='_blank'><img data-src='{overlay}' alt='{label}, changes boxed'></a>"
             if overlay else "")
    return (f"<div class='side{' after' if after else ''}'><div class='label'>{label}</div>"
            f"<a class='plain' href='{src}' target='_blank'><img data-src='{src}' alt='{label}'></a>{extra}</div>")


def _details(info: dict) -> str:
    if not info:
        return ""
    parts = []
    mapping = info.get("mapping") or {}
    if mapping:
        rows = "".join(f"<tr><td><code>{html.escape(original)}</code></td><td><code>{html.escape(new)}</code></td></tr>"
                       for original, new in sorted(mapping.items(), key=lambda item: item[0].lower()))
        parts.append(f"<details><summary>{len(mapping)} string(s) obfuscated</summary>"
                     f"<table><tr><th>Original</th><th>Replaced by</th></tr>{rows}</table></details>")
    logos = info.get("logos") or []
    if logos:
        rows = "".join(f"<tr><td class='num'>{logo['page']}</td><td>{html.escape(logo['reference'])}</td>"
                       f"<td class='num'>{logo['score']:.2f}</td><td><code>{', '.join(f'{v:.0f}' for v in logo['box'])}</code></td></tr>"
                       for logo in logos)
        parts.append(f"<details><summary>{len(logos)} logo(s) removed</summary><table><tr><th>Page</th>"
                     f"<th>Like reference</th><th>Similarity</th><th>Box (pt)</th></tr>{rows}</table></details>")
    for key, what in (("signatures", "signature area(s)"), ("stamps", "stamp(s)")):
        areas = info.get(key) or []
        if areas:
            rows = "".join(f"<tr><td class='num'>{area['page']}</td><td><code>"
                           f"{', '.join(f'{v:.0f}' for v in area['box'])}</code></td></tr>" for area in areas)
            parts.append(f"<details><summary>{len(areas)} {what} blacked out</summary><table><tr>"
                         f"<th>Page</th><th>Box (pt)</th></tr>{rows}</table></details>")
    if info.get("error"):
        parts.append(f"<p class='warn'>Failed: {html.escape(info['error'])}</p>")
    return "".join(parts)


def build_report(originals, obfuscated, report_dir, summary: dict = None, title: str = "Obfuscation check",
                 progress=print) -> Path:
    """Write report_dir/report.html; returns its path."""
    originals, obfuscated, report_dir = Path(originals), Path(obfuscated), Path(report_dir)
    img_dir = report_dir / "img"
    img_dir.mkdir(parents=True, exist_ok=True)
    summary = summary or {}
    rows, tabs, panels = [], [], []
    for result in sorted(obfuscated.rglob("*.pdf")):
        if result.name.startswith("."):
            continue  # a step's temporary file
        relative = result.relative_to(obfuscated).with_name(_original_name(result.name))
        original = originals / relative
        if not original.exists():
            continue
        stem = _stem_id(relative.as_posix())
        pages = compare_document(original, result, img_dir, stem)
        changed = sum(1 for page in pages if page["changes"])
        info = summary.get(relative.as_posix(), summary.get(original.name, {}))
        strings, logos = len(info.get("mapping") or {}), len(info.get("logos") or [])
        signatures, stamps = len(info.get("signatures") or []), len(info.get("stamps") or [])
        progress(f"{relative.as_posix()}: {len(pages)} page(s), {changed} changed")
        name = html.escape(relative.with_suffix("").as_posix())
        rows.append(f"<tr><td><a href='#{stem}' onclick=\"show('{stem}');return false\">{name}</a></td>"
                    f"<td class='num'>{len(pages)}</td><td class='num'>{changed}</td>"
                    f"<td class='num'>{strings if info else '–'}</td><td class='num'>{logos if info else '–'}</td>"
                    f"<td class='num'>{signatures if info else '–'}</td><td class='num'>{stamps if info else '–'}</td>"
                    f"<td class='num {'warn' if changed == 0 else ''}'>{len(pages) - changed}</td></tr>")
        tabs.append(f"<button class='tab' data-id='{stem}'>{name} <span class='n'>{changed}/{len(pages)}</span></button>")
        body = [f"<h2>{name}</h2><p class='muted'>{len(pages)} page(s), {changed} with changes. Red boxes: what "
                f"differs between the original and the obfuscated page.</p>", _details(info),
                "<div class='tools'><label><input type='checkbox' data-act='hide-unchanged'> hide unchanged pages</label>"
                "<label><input type='checkbox' data-act='overlay'> box the changes on the obfuscated side too</label></div>"]
        for page in pages:
            badge = (f"<span class='badge chg'>{page['changes']} change(s)</span>" if page["changes"]
                     else "<span class='badge same'>unchanged</span>")
            body.append(f"<div class='pair{'' if page['changes'] else ' unchanged'}'><h3>Page {page['page']}{badge}</h3>"
                        f"<div class='sides'>{_side('Original', page.get('img_a'))}"
                        f"{_side('Obfuscated', page.get('img_b'), page.get('img_ov'), after=True)}</div></div>")
        panels.append(f"<section class='panel' id='{stem}'>{''.join(body)}</section>")
    overview = ("<section class='panel' id='overview'><h2>Overview</h2>"
                "<p class='muted'>Unchanged pages are worth a look: a page with nothing obfuscated may still hold "
                "something sensitive the targets missed (scans are only read with OCR).</p>"
                "<table><tr><th>Document</th><th>Pages</th><th>Pages changed</th><th>Strings obfuscated</th>"
                f"<th>Logos removed</th><th>Signature areas</th><th>Stamps</th><th>Pages unchanged</th></tr>{''.join(rows)}</table></section>")
    page = (f"<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' "
            f"content='width=device-width,initial-scale=1'><title>{html.escape(title)}</title><style>{CSS}</style></head><body>"
            f"<header><h1>{html.escape(title)}: original vs obfuscated</h1>"
            f"<p class='muted'>Originals: {html.escape(str(originals))}<br>Obfuscated: {html.escape(str(obfuscated))}</p></header>"
            f"<nav class='tabs'><button class='tab' data-id='overview'>Overview</button>{''.join(tabs)}</nav>"
            f"{overview}{''.join(panels)}<script>{JS}</script></body></html>")
    path = report_dir / "report.html"
    path.write_text(page, encoding="utf-8")
    return path


def _command_line(arguments=None) -> None:
    parser = ArgumentParser(prog="python -m acd.obfuscation_report", description=__doc__.split("\n\n")[0])
    parser.add_argument("originals")
    parser.add_argument("obfuscated")
    parser.add_argument("report_dir")
    parser.add_argument("--summary", help="JSON written by `obfuscation_ai apply --summary`")
    parser.add_argument("--title", default="Obfuscation check")
    args = parser.parse_args(arguments)
    summary = json.loads(Path(args.summary).read_text(encoding="utf-8")) if args.summary else {}
    print(build_report(args.originals, args.obfuscated, args.report_dir, summary, args.title))


if __name__ == "__main__":
    _command_line()
