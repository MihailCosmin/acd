"""Read illustrated parts lists from S1000D IPD data modules and ATA iSpec 2200 IPLs.

Two markups, one neutral result (plain dicts, no Qt, standard library only):

- S1000D IPD data modules (info code 941), issues 1.7 to 6, SGML or XML: ``catalogSeqNumber``
  (``csn``) with its ``itemSeqNumber`` (``isn``) entries: part number, manufacturer code,
  quantity per next higher assembly, indenture, description, attaching parts, applicability,
  references to other catalogue sequence numbers.
- ATA iSpec 2200 CMM IPLs, SGML or XML: ``dplist/figure`` with ``prtlist/itemdata`` (item number,
  indenture, illustrated, attaching, ``pnr``, ``iplnom`` with ``nom``/``kwd``/``adt``, ``mfr`` and
  the remark elements, ``effcode``, ``upa``) and the vendor list ``vendlist/vendata``.

``read_ipl(path_or_text)`` detects the markup. The result::

    {"format": "s1000d-ipd" | "ispec2200",
     "document": {"dmc": ..., "title": ...},           # what identifies the source
     "figures": [{"number", "title", "sheets", "at", "items": [item, ...]}],
     "vendors": [{"code", "name", "address"}],
     "notes": [{"level": "error" | "warning", "at", "message"}]}

    item = {"item", "part_number", "nomenclature", "indenture" (int or None),
            "quantity" (as written), "manufacturer" (as written), "effectivity" (codes),
            "illustrated", "attaching", "remarks" (lines), "at"}

``at`` is the path of the source element below the catalogue (``catalogSeqNumber[3]/itemSeqNumber[2]``,
``figure[2]/prtlist/itemdata[5]``). Every guess is a note; nothing is dropped silently. The SGML is read without
its DTD: an element that is never closed in the file ends at the next start tag once it has text,
at the next element of its own name, or at the end tag of an element around it.
"""

from __future__ import annotations

import html.entities
import re
from dataclasses import dataclass, field
from pathlib import Path
from collections.abc import Iterator

__all__ = ["Node", "parse_markup", "read_ipl", "read_s1000d_ipd", "read_ispec2200_ipl",
           "detect", "IplReadError"]

S1000D = "s1000d-ipd"
ISPEC = "ispec2200"
LETTERS = "ABCDEFGHJKLMNPQRSTUVWXYZ"  # variant letters without I and O


class IplReadError(ValueError):
    """The text holds no IPL markup this module reads."""


# -- tolerant SGML/XML tree


@dataclass(eq=False)
class Node:
    tag: str  # as written
    attrs: dict[str, str] = field(default_factory=dict)  # lower-case names
    parts: list[str | Node] = field(default_factory=list)
    parent: Node | None = None

    @property
    def name(self) -> str:
        """Lower-case local name (no namespace prefix)."""
        return self.tag.rsplit(":", 1)[-1].lower()

    @property
    def children(self) -> list[Node]:
        return [p for p in self.parts if isinstance(p, Node)]

    def get(self, attr: str, default: str = "") -> str:
        return self.attrs.get(attr.lower(), default)

    def own_text(self) -> str:
        """The text directly in this element, whitespace collapsed."""
        return _ws(" ".join(p for p in self.parts if isinstance(p, str)))

    def text(self) -> str:
        """All text in this element and below, whitespace collapsed."""
        out: list[str] = []

        def walk(n: Node) -> None:
            for p in n.parts:
                if isinstance(p, str):
                    out.append(p)
                else:
                    walk(p)

        walk(self)
        return _ws(" ".join(out))

    def iter(self, *names: str, stop: tuple[str, ...] = ()) -> Iterator[Node]:
        """Descendants named one of `names` (all with none), document order; does not go into
        elements named in `stop` (other than self)."""
        want = {n.lower() for n in names}
        stops = {s.lower() for s in stop}
        for c in self.children:
            if not want or c.name in want:
                yield c
            if c.name not in stops:
                yield from c.iter(*names, stop=stop)

    def find(self, *names: str, stop: tuple[str, ...] = ()) -> Node | None:
        return next(self.iter(*names, stop=stop), None)

    def child(self, *names: str) -> Node | None:
        want = {n.lower() for n in names}
        return next((c for c in self.children if c.name in want), None)

    def path(self, base: Node | None = None) -> str:
        """`a/b[2]/c` from below `base` (or the root) to this element; the index only when
        several siblings share the name."""
        steps: list[str] = []
        n: Node | None = self
        while n is not None and n is not base and n.parent is not None:
            same = [c for c in n.parent.children if c.name == n.name]
            steps.append(n.tag if len(same) == 1 else f"{n.tag}[{same.index(n) + 1}]")
            n = n.parent
        return "/".join(reversed(steps))


def _ws(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip()


_ENTITY = re.compile(r"&(#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9._-]*);?")


def _entities(s: str, extra: dict[str, str]) -> str:
    def rep(m: re.Match[str]) -> str:
        e = m[1]
        if e.startswith("#"):
            try:
                return chr(int(e[2:], 16) if e[1] in "xX" else int(e[1:]))
            except (ValueError, OverflowError):
                return m[0]
        if e in extra:
            return extra[e]
        got = html.entities.html5.get(e + ";")
        return got if got is not None else m[0]

    return _ENTITY.sub(rep, s) if "&" in s else s


_TOKEN = re.compile(
    r"<!--.*?-->"
    r"|<!\[CDATA\[(?P<cdata>.*?)\]\]>"
    r"|<!\[\s*(?P<ms>[A-Za-z]+|%[\w.-]+;)\s*\["  # marked section start
    r"|\]\]>"
    r"|<!DOCTYPE[^\[>]*(?:\[(?P<subset>.*?)\]\s*)?>"
    r"|<![^>]*>"
    r"|<\?[^>]*>"
    r"|</\s*(?P<end>[A-Za-z_][\w:.-]*)?\s*>"
    r"|<(?P<start>[A-Za-z_][\w:.-]*)(?P<attrs>(?:\s+[^\s=>/]+(?:\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s>]+))?)*)\s*(?P<empty>/?)>",
    re.DOTALL,
)
_ATTR = re.compile(r"([^\s=>/]+)(?:\s*=\s*(\"[^\"]*\"|'[^']*'|[^\s>]+))?")
_DECL = re.compile(r"<!ENTITY\s+([A-Za-z][\w.-]*)\s+(?:CDATA\s+|SDATA\s+)?(\"[^\"]*\"|'[^']*')\s*>", re.I)


# elements declared EMPTY in the S1000D and iSpec 2200 DTDs: never containers, even unclosed
EMPTY = frozenset({"ipp", "rfs", "mov", "nsn", "asp", "csnref", "sym", "initialprovisioningproject",
                   "reasonforselection", "partref", "attachstoreshippart", "notillustrated",
                   "catalogseqnumberref", "functionalitemref", "applicref", "sheet-ref", "refint",
                   "issno", "issdate", "security", "rpc", "model", "version", "dmcode", "issueinfo",
                   "language", "issuedate", "graphic-ref"})


def parse_markup(text: str, empty: frozenset[str] = EMPTY) -> Node:
    """A tree of an SGML or XML document, tolerant of omitted end tags, upper-case names,
    SGML processing instructions and undeclared entities. Elements named in `empty` (lower case)
    never hold content. The returned root has the tag `#root`."""
    entities: dict[str, str] = {}
    closed = {m.lower().rsplit(":", 1)[-1] for m in re.findall(r"</\s*([A-Za-z_][\w:.-]*)", text)}
    root = Node("#root")
    stack = [root]
    pos = 0
    skip_to: int | None = None

    def add_text(s: str) -> None:
        if s:
            stack[-1].parts.append(_entities(s, entities))

    for m in _TOKEN.finditer(text):
        if skip_to is not None:
            if m.start() < skip_to:
                continue
            skip_to = None
        add_text(text[pos:m.start()])
        pos = m.end()
        if m["cdata"] is not None:
            stack[-1].parts.append(m["cdata"])
        elif m["ms"] is not None:
            if m["ms"].upper() == "IGNORE":
                end = text.find("]]>", m.end())
                skip_to = len(text) if end < 0 else end + 3
                pos = skip_to
        elif m["subset"] is not None:
            for d in _DECL.finditer(m["subset"]):
                entities[d[1]] = d[2][1:-1]
        elif m["start"] is not None:
            node = Node(m["start"])
            for a in _ATTR.finditer(m["attrs"] or ""):
                v = a[2] or a[1]
                if v[:1] in "\"'" and v[-1:] == v[:1] and len(v) > 1:
                    v = v[1:-1]
                node.attrs[a[1].lower()] = _entities(v, entities).strip()
            name = node.name
            # implied end tags of elements that are never closed in this file
            while len(stack) > 1 and stack[-1].name not in closed and stack[-1].own_text():
                stack.pop()
            if name not in closed:
                for i in range(len(stack) - 1, 0, -1):
                    if stack[i].name == name:
                        del stack[i:]
                        break
            node.parent = stack[-1]
            stack[-1].parts.append(node)
            if not m["empty"] and not (name in empty and name not in closed):
                stack.append(node)
        elif m.group(0).startswith("</"):
            end = (m["end"] or "").lower().rsplit(":", 1)[-1]
            if not end:  # SGML empty end tag </>
                if len(stack) > 1:
                    stack.pop()
                continue
            for i in range(len(stack) - 1, 0, -1):
                if stack[i].name == end:
                    del stack[i:]
                    break
    add_text(text[pos:])
    return root


# -- reading


def _read_text(source: str | Path | bytes) -> str:
    if isinstance(source, bytes):
        raw = source
    elif isinstance(source, Path) or (isinstance(source, str) and "<" not in source[:4096]
                                      and len(source) < 1024):
        raw = Path(source).read_bytes()
    else:
        return str(source)
    head = raw[:200].decode("ascii", "replace")
    enc = re.search(r"encoding\s*=\s*[\"']([\w.-]+)", head)
    for e in ([enc[1]] if enc else []) + ["utf-8-sig", "cp1252"]:
        try:
            return raw.decode(e)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("latin-1")


def detect(root: Node) -> str | None:
    """`s1000d-ipd`, `ispec2200` or None."""
    if root.find("catalogseqnumber", "csn") is not None:
        return S1000D
    if root.find("itemdata") is not None:
        return ISPEC
    return None


def read_ipl(source: str | Path | bytes) -> dict:
    """Read an S1000D IPD data module or an iSpec 2200 IPL (a path, the bytes or the text)."""
    root = parse_markup(_read_text(source))
    kind = detect(root)
    if kind == S1000D:
        return _s1000d(root)
    if kind == ISPEC:
        return _ispec(root)
    raise IplReadError("no IPL markup found: neither S1000D catalogSeqNumber/csn nor iSpec 2200 itemdata")


def read_s1000d_ipd(text: str) -> dict:
    return _s1000d(parse_markup(text))


def read_ispec2200_ipl(text: str) -> dict:
    return _ispec(parse_markup(text))


class _Out:
    def __init__(self, fmt: str):
        self.data: dict = {"format": fmt, "document": {}, "figures": [], "vendors": [], "notes": []}

    def note(self, level: str, at: str, message: str) -> None:
        self.data["notes"].append({"level": level, "at": at, "message": message})

    def figure(self, number: str, at: str) -> dict:
        for f in self.data["figures"]:
            if f["number"] == number:
                return f
        f = {"number": number, "title": "", "sheets": 0, "at": [at] if at else [], "items": []}
        self.data["figures"].append(f)
        return f


def _number(s: str) -> str:
    """`010` -> `10`, `01A` -> `1A`, `000` -> `0`."""
    s = s.strip().upper()
    m = re.fullmatch(r"0*([0-9]+)([A-Z]?)", s)
    if not m:
        return s
    return (m[1].lstrip("0") or "0") + m[2]


def _int(s: str) -> int | None:
    s = s.strip()
    return int(s) if re.fullmatch(r"[0-9]{1,2}", s) else None


def _item(at: str, number: str, pn: str, nom: str, indenture: int | None) -> dict:
    return {"item": number, "part_number": pn, "nomenclature": nom, "indenture": indenture,
            "quantity": "", "manufacturer": "", "effectivity": [], "illustrated": True,
            "attaching": False, "remarks": [], "at": at}


# -- S1000D IPD

_CSN_VALUE = re.compile(r"([0-9]{2})([A-Z]?)([0-9]{3})([A-Z]?)$")


def _s1000d(root: Node) -> dict:
    out = _Out(S1000D)
    out.data["document"] = _dm_ident(root)
    cat = root.find("illustratedpartscatalog", "ipc") or root
    applic = _applics(root)
    csns = list(cat.iter("catalogseqnumber", "csn", stop=("catalogseqnumber", "csn")))
    # figure number and variant of each CSN first: a variant counts only when the module has two
    rows: list[tuple[Node, str, str, str]] = []
    last_fig = ""
    for csn in csns:
        fig, var, item = _csn_ident(csn)
        if not fig:
            fig = last_fig
        if not fig:
            fig = "1"
            out.note("warning", csn.path(cat), "no figure number: figure 1 given")
        last_fig = fig
        rows.append((csn, fig, var, item))
    variants: dict[str, set[str]] = {}
    for _, fig, var, _ in rows:
        if var:
            variants.setdefault(fig, set()).add(var)
    figure_nodes = [f for f in cat.iter("figure", stop=("catalogseqnumber", "csn"))]
    zero: set[str] = set()
    no_descr: list[str] = []
    for csn, fig, var, item in rows:
        number = fig + (var if len(variants.get(fig, ())) > 1 else "")
        if re.sub(r"[A-Z]+$", "", number) in ("", "0"):
            if number not in zero:
                zero.add(number)
                out.note("warning", csn.path(cat), f"figure number {number!r} is zero: figure 1 given")
            number = "1"
        at_csn = csn.path(cat)
        f = out.figure(number, at_csn)
        isns = list(csn.iter("itemseqnumber", "itemsequencenumber", "isn", stop=("itemseqnumber", "itemsequencenumber", "isn")))
        if not isns:
            out.note("warning", at_csn, f"CSN {fig}-{item}: no item sequence number, skipped")
            continue
        ind = _int(csn.get("indenture") or csn.get("ind"))
        if ind is None:
            out.note("warning", at_csn, f"CSN {fig}-{item}: no indenture")
        if len(isns) > 1:
            names = ", ".join(i.get("itemseqnumbervalue") or i.get("isn") or i.get("id") or "?" for i in isns)
            out.note("warning", at_csn, f"item {number}-{item}: {len(isns)} item sequence numbers ({names}): the alternatives get the variant letters {', '.join(item + LETTERS[k] for k in range(len(isns) - 1))}, check them")  # noqa: E501
        for k, isn in enumerate(isns):
            at = isn.path(cat)
            no = item if k == 0 else item + LETTERS[min(k - 1, len(LETTERS) - 1)]
            it = _item(at, no, *_s1000d_part(isn), ind)
            if not it["part_number"]:
                out.note("warning", at, f"item {number}-{no}: no part number, skipped")
                continue
            if not it["nomenclature"]:
                no_descr.append(it["part_number"])
            qna = isn.find("quantitypernexthigherassy", "qna")
            it["quantity"] = qna.own_text() or qna.text() if qna is not None else ""
            ref = isn.find("partref")
            mfr = ref.get("manufacturercodevalue") if ref is not None else ""
            if not mfr:
                m = isn.find("manufacturercode", "mfc")
                mfr = m.text() if m is not None else ""
            it["manufacturer"] = mfr
            asp = isn.find("attachstoreshippart", "asp")
            if asp is not None and (asp.get("attachstoreshippartcode") or asp.get("asp") or asp.own_text()) == "1":
                it["attaching"] = True
            if isn.find("notillustrated") is not None or csn.find("notillustrated") is not None:
                it["illustrated"] = False
            for loc in isn.iter("descrforlocation"):
                if loc.text():
                    it["remarks"].append(loc.text())
            for r in isn.iter("catalogseqnumberref"):
                rf = _number(r.get("figurenumber")) if r.get("figurenumber") else ""
                ri = _number(r.get("item")) if r.get("item") else ""
                tgt = f"FIG {rf} ITEM {ri}".strip() if (rf or ri) else (r.get("catalogseqnumbervalue") or "").strip()
                if tgt:
                    it["remarks"].append(f"SEE {tgt}")
            for r in isn.iter("csnref"):
                if r.get("refcsn"):
                    it["remarks"].append(f"SEE CSN {_ws(r.get('refcsn'))}")
            for r in isn.iter("functionalitemref"):
                fin = r.get("functionalitemnumber") or r.text()
                if fin:
                    it["remarks"].append(f"FIN {fin}")
            for u in isn.iter("usableoncodeassy", "usableoncodeequip", "uca", "uce"):
                it["effectivity"] += [c for c in re.split(r"[\s,;]+", u.text().upper()) if c]
            for node in (isn, csn):
                rid = node.get("applicrefid")
                if rid:
                    shown = applic.get(rid, "")
                    it["remarks"].append(f"APPLIC {shown or rid}")
                    out.note("warning", at, f"{it['part_number']}: applicability {rid}{' (' + shown + ')' if shown else ''} put into the remarks, check the effect code")  # noqa: E501
                    break
            f["items"].append(it)
    _s1000d_titles(out, figure_nodes, cat)
    if no_descr:
        named = ", ".join(no_descr[:10]) + (f" and {len(no_descr) - 10} more" if len(no_descr) > 10 else "")
        out.note("warning", "", f"{len(no_descr)} parts without descrForPart (from S1000D Issue 5.0 the part data can be in a separate parts data module): {named}")  # noqa: E501
    if not csns:
        out.note("error", "", "no catalogSeqNumber found")
    return out.data


def _s1000d_part(isn: Node) -> tuple[str, str]:
    ref = isn.find("partref")
    pn = ref.get("partnumbervalue") if ref is not None else ""
    if not pn:
        p = isn.find("partnumber", "pnr")
        pn = p.text() if p is not None else ""
    d = isn.find("descrforpart", "dfp")
    return pn.strip(), d.text() if d is not None else ""


def _csn_ident(csn: Node) -> tuple[str, str, str]:
    """(figure, figure variant, item with its variant) of a catalogue sequence number."""
    fig = csn.get("figurenumber")
    if fig:
        return (_number(fig), csn.get("figurenumbervariant").strip().upper(),
                _number(csn.get("item") or "0") + csn.get("itemvariant").strip().upper())
    value = re.sub(r"\s+", "", (csn.get("catalogseqnumbervalue") or csn.get("csn")).upper())
    m = _CSN_VALUE.search(value)
    item_attr = csn.get("item").strip()
    if m:
        return _number(m[1]), m[2], _number(item_attr or m[3]) + m[4]
    return "", "", _number(item_attr) if item_attr else ""


def _s1000d_titles(out: _Out, figure_nodes: list[Node], cat: Node) -> None:
    figs = out.data["figures"]
    titled = [(fn, fn.child("title")) for fn in figure_nodes]
    for k, f in enumerate(figs):
        fn = titled[k][0] if k < len(titled) else (titled[0][0] if len(titled) == 1 else None)
        if fn is None:
            out.note("warning", "", f"figure {f['number']}: no figure element, no title")
            continue
        t = fn.child("title")
        f["title"] = t.text() if t is not None else ""
        f["sheets"] = len(list(fn.iter("graphic")))
        f["at"] = [fn.path(cat)]
    if len(figs) > 1 and len(titled) == 1:
        out.note("warning", "", f"{len(figs)} figure numbers but one figure element: all get its title")


def _dm_ident(root: Node) -> dict:
    code = root.find("dmcode")
    if code is not None:
        keys = ["modelidentcode", "systemdiffcode", "systemcode", "subsystemcode", "subsubsystemcode",
                "assycode", "disassycode", "disassycodevariant", "infocode", "infocodevariant",
                "itemlocationcode"]
        v = [code.get(k) for k in keys]
        dmc = "DMC-{}-{}-{}-{}{}-{}-{}{}-{}{}-{}".format(*v)
    else:
        avee = root.find("avee")
        if avee is None:
            return {}
        keys = ["modelic", "sdc", "chapnum", "section", "subsect", "subject", "discode", "discodev",
                "incode", "incodev", "itemloc"]
        v = [(avee.find(k).own_text() if avee.find(k) is not None else "") for k in keys]
        dmc = "DMC-{}-{}-{}-{}{}-{}-{}{}-{}{}-{}".format(*v)
    tech, info = root.find("techname"), root.find("infoname")
    title = " - ".join(x.own_text() or x.text() for x in (tech, info) if x is not None)
    return {"dmc": dmc, "title": title}


def _applics(root: Node) -> dict[str, str]:
    out = {}
    for a in root.iter("applic"):
        if a.get("id"):
            d = a.find("displaytext")
            out[a.get("id")] = d.text() if d is not None else ""
    return out


# -- iSpec 2200

# remark elements of iplnom whose parts print in a fixed form (iSpec 2200 CMM)
_FORMS = {"afpmfr": "MODIFIED FROM {0} MADE BY {1}", "alprtmfr": "ALT TO {0} {1}",
          "optmfr": "OPT TO {0} {1}"}


def _ispec(root: Node) -> dict:
    out = _Out(ISPEC)
    title = root.find("title")
    cmm = root.find("cmm") or root
    out.data["document"] = {k: cmm.get(k) for k in ("chapnbr", "sectnbr", "subjnbr", "spl", "model", "tsn")
                            if cmm.get(k)}
    if title is not None:
        out.data["document"]["title"] = title.text()
    dplist = root.find("dplist") or root
    base = dplist.parent if dplist is not root else root
    figs = [f for f in dplist.iter("figure") if f.find("itemdata") is not None]
    by_number: dict[str, list[Node]] = {}
    for fn in figs:
        by_number.setdefault(_number(fn.get("fignbr") or ""), []).append(fn)
    for k, fn in enumerate(figs, 1):
        at_fig = fn.path(base)
        number = _number(fn.get("fignbr") or "")
        if not number:
            number = str(k)
            out.note("warning", at_fig, f"figure without fignbr: {number} given")
        elif len(by_number.get(number, [])) > 1 and fn.get("confltr") and not number[-1].isalpha():
            number += fn.get("confltr").upper()
        f = out.figure(number, at_fig)
        t = fn.child("title")
        f["title"] = t.text() if t is not None else ""
        f["sheets"] = len(list(fn.iter("sheet"))) or len(list(fn.iter("graphic")))
        eff = fn.child("effect")
        if eff is not None and eff.text():
            out.note("warning", at_fig, f"figure {number}: effectivity {eff.text()[:80]!r} not carried over")
        for item in fn.iter("itemdata", stop=("itemdata",)):
            it = _ispec_item(item, base, out, number)
            if it is None:
                continue

            f["items"].append(it)
        if not f["items"]:
            out.note("warning", at_fig, f"figure {number}: no item with a part number")
    for v in (root.find("vendlist") or Node("#")).iter("vendata"):
        mfr, mad = v.child("mfr"), v.child("mad")
        code = mfr.text() if mfr is not None else ""
        if not code:
            continue
        addr = mad.text() if mad is not None else ""
        name, _, rest = addr.partition(",")
        out.data["vendors"].append({"code": code, "name": name.strip(), "address": rest.strip()})
    if not figs:
        out.note("error", "", "no figure with itemdata found")
    return out.data


def _ispec_item(item: Node, root: Node, out: _Out, fig: str) -> dict | None:
    at = item.path(root)
    raw = (item.get("itemnbr") or "").strip()
    pnr = item.child("pnr")
    pn = pnr.text() if pnr is not None else ""
    if not pn:
        if raw or item.text():
            out.note("warning", at, f"item {fig}-{raw or '?'}: no part number, skipped")
        return None
    dash = raw.startswith(("-", "–"))
    number = _number(raw.lstrip("-– "))
    if not number:
        out.note("warning", at, f"{pn}: no itemnbr")
    ind = _int(item.get("indent"))
    it = _item(at, number, pn, "", ind)
    if dash or item.get("illusind") == "0":
        it["illustrated"] = False
    if item.get("attach") == "1":
        it["attaching"] = True
    if item.get("delitem") == "1":
        it["remarks"].append("DELETED")
        out.note("warning", at, f"{pn}: deleted item (delitem=1)")
    nom = item.child("iplnom")
    noms: list[str] = []
    for c in (nom.children if nom is not None else []):
        if c.name == "nom":
            kwd, adt = c.child("kwd"), c.child("adt")
            k = kwd.text() if kwd is not None else ""
            a = adt.text() if adt is not None else ""
            if a and k and not a.upper().startswith(k.upper()):
                noms.append(f"{k}, {a}")
            else:
                noms.append(a or k)
        elif c.name == "mfr":
            it["manufacturer"] = c.text()
        elif c.name in _FORMS:
            parts = [x.text() for x in c.children]
            if len(parts) >= 2:
                it["remarks"].append(_FORMS[c.name].format(*parts[:2]).strip())
            elif c.text():
                it["remarks"].append(c.text())
        elif c.text():
            it["remarks"].append(c.text())
    noms = [n for n in noms if n]
    if noms:
        it["nomenclature"] = noms[0]
        it["remarks"][:0] = noms[1:]
    else:
        out.note("warning", at, f"{pn}: no nomenclature (iplnom/nom/kwd)")
    it["effectivity"] = [c for e in item.iter("effcode") for c in re.split(r"[\s,;]+", e.text().upper()) if c]
    upa = item.child("upa")
    it["quantity"] = upa.text() if upa is not None else ""
    return it
