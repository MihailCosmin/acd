"""XML syntax and schema checking, reported the way `BrexChecker` reports BREX
violations.

`BrexChecker` answers "does this object obey its BREX", but it can only answer
that for an object that parses and means what it says. This module answers the
questions that come first, in the order a failure actually cascades:

1. **encoding** -- is the byte stream something an XML parser can read at all
   (non-empty, decodable as its declared encoding, free of characters XML 1.0
   forbids)? A BOM is reported too: it parses, but breaks downstream
   publishing tools that concatenate files.
2. **syntax** -- is it well-formed? Every entry of the parser's error log is
   reported, not just the first, so one pass shows every mismatched tag and
   undeclared entity instead of one per re-run.
3. **structure** -- does the object make sense as an S1000D CSDB object: does
   its root element match what the filename claims, does it declare a schema
   at all, and does the code in the filename match the code inside the file?
4. **DTD** -- if a DOCTYPE is declared, does the object validate against it?
   (S1000D 2.x/3.x and CMM content; 4.x+ is XSD-driven and skips this.)
5. **schema** -- does it validate against its
   `xsi:noNamespaceSchemaLocation`, again reporting the full error log?

A layer that cannot run is skipped rather than guessed at: a file that is not
well-formed is never schema-checked, because every schema error it produced
would be an artefact of the syntax error.

Schema resolution never depends on the network being up more than once. A
declared schema is looked for in the caller's search paths first, then in an
on-disk cache, and only then fetched -- and a fetch pulls the whole
`include`/`import` tree into the cache, so the second run of a folder is
fully offline. Compiled schemas are additionally cached in memory for the
duration of a run, so a folder of 90 data modules sharing one schema compiles
it once.

The reports (`to_excel_report`, `to_html_report`) are built on the same
`acd.report` layer as `BrexChecker`'s, so a reviewer reads one layout.
"""

from datetime import datetime
from json import dumps
from os import listdir
from os import makedirs
from os.path import basename
from os.path import dirname
from os.path import isdir
from os.path import isfile
from os.path import join
from os.path import expanduser
from re import search as re_search
from urllib.parse import urljoin
from urllib.parse import urlparse
from urllib.request import urlopen

from lxml import etree

from .filepath import clean_path
from .prompt import ask_for_folder
from .report import ExcelReport
from .report import HtmlReport
from .resources import find_schema
from .resources import register_catalog
from .xml_processing import get_schema_from_xml


# The five layers, in the order they run and the order they are reported in.
# A finding's `Check` is always one of these.
CHECK_LAYERS = ("encoding", "syntax", "structure", "dtd", "schema")

# Every rule this module can report, with the human-readable summary shown in
# the reports. Keeping them in one table means a reader of a report and a
# reader of the code see the same list, and a caller can suppress a rule by
# name without knowing which layer raised it.
CHECK_RULES = {
    "ENC-EMPTY": "File is empty or contains only whitespace.",
    "ENC-BOM": "File starts with a UTF-8 byte order mark.",
    "ENC-DECL": "File does not decode as the encoding it declares.",
    "ENC-CTRL": "File contains a control character XML 1.0 forbids.",
    "XML-SYNTAX": "The file is not well-formed XML.",
    "STR-ROOT": "Root element does not match the object type the filename declares.",
    "STR-NOSCHEMA": "Object declares neither a schema nor a DOCTYPE.",
    "STR-IDENT": "The code inside the file does not match the code in the filename.",
    "DTD-VALID": "The object does not validate against its DTD.",
    "DTD-MISSING": "A DOCTYPE is declared but its DTD could not be loaded.",
    "XSD-VALID": "The object does not validate against its schema.",
    "XSD-UNRESOLVED": "The declared schema could not be found or fetched.",
    "XSD-BROKEN": "The declared schema could not be compiled.",
}

# Root element each S1000D filename prefix implies. Anything not listed here
# (ICN, SCO, ...) is not checked for a root element.
_ROOT_BY_PREFIX = {
    "DMC": "dmodule",
    "PMC": "pm",
    "DML": "dml",
    "DDN": "ddn",
    "COM": "comment",
}

# Codes the filename of a DMC/PMC/DDN object encodes, and where the matching
# code element lives inside the object. Used by the STR-IDENT check, which is
# what catches an object that was renamed without its ident being updated (or
# the reverse).
#
# The path matters: an object carries other objects' codes too -- every
# `dmRef`, and the `brexDmRef` naming its BREX, contains a `dmCode` element.
# Only the one under the ident is this object's own.
_IDENT_ELEMENT = {
    "DMC": (".//dmIdent/dmCode", "dmCode", (
        "modelIdentCode", "systemDiffCode", "systemCode", "subSystemCode",
        "subSubSystemCode", "assyCode", "disassyCode", "disassyCodeVariant",
        "infoCode", "infoCodeVariant", "itemLocationCode",
    )),
    "PMC": (".//pmIdent/pmCode", "pmCode",
            ("modelIdentCode", "pmIssuer", "pmNumber", "pmVolume")),
    "DDN": (".//ddnIdent/ddnCode", "ddnCode", (
        "modelIdentCode", "senderIdent", "receiverIdent", "yearOfDataIssue",
        "seqNumber",
    )),
}

# Characters XML 1.0 forbids outright: C0 controls other than tab, newline and
# carriage return. A file carrying one parses nowhere, and the parser's own
# message ("PCDATA invalid Char value 3") does not say where it came from.
_ILLEGAL_CONTROL_BYTES = (
    set(range(0x00, 0x09)) | {0x0B, 0x0C} | set(range(0x0E, 0x20))
)

_DEFAULT_SCHEMA_CACHE = join(expanduser("~"), ".acd", "schema-cache")


class XmlChecker():
    """Check XML objects for encoding, well-formedness, structural sanity, DTD
    and schema validity, and report the result as JSON, an Excel workbook or a
    self-contained HTML page.

    The API deliberately mirrors `BrexChecker`: `set_xml`/`set_xml_dir` choose
    what to check, `validate()` returns a result dict of the same shape
    (single-object, or `{filename: result}` in directory mode), and
    `run_summary`/`to_excel_report`/`to_html_report` read that result.

    Typical use::

        checker = XmlChecker()
        checker.set_xml_dir(r"C:\\CSDB\\CRP 67-34-02")
        result = checker.validate()
        checker.to_excel_report(result, "check.xlsx")
        checker.to_html_report(result, "check.html")
    """

    def __init__(self):
        # Register the bundled entity catalog before anything parses: libxml2
        # reads XML_CATALOG_FILES once per process, on its first catalog
        # lookup, so this has to happen at construction rather than at parse
        # time. Without it, a data module referencing the ISO character
        # entities by their s1000d.org URL needs network access to parse.
        register_catalog()

        self._xml_path = None
        self._xml_dir = None

        self._schema_search_paths = []
        self._schema_cache_dir = _DEFAULT_SCHEMA_CACHE
        self._allow_network = True

        # Compiled `etree.XMLSchema` per resolved location, so a folder whose
        # objects share a schema compiles it once per run rather than once per
        # object. `False` marks a schema that failed to compile, so a broken
        # schema is not recompiled for every object that declares it.
        self._schema_cache = {}
        # Resolved location per declared location, same reasoning.
        self._resolved_schemas = {}

    # ------------------------------------------------------------------
    # Inputs
    # ------------------------------------------------------------------

    def set_xml(self, xml: str) -> None:
        """Check a single object.

        Args:
            xml (str): path to the XML file
        """
        self._xml_path = xml
        self._xml_dir = None

    def set_xml_dir(self, dir_path: str) -> None:
        """Check every `.xml` file in a directory (not recursive), the way
        `BrexChecker.set_xml_dir` does.

        Args:
            dir_path (str): directory holding the objects to check
        """
        self._xml_dir = dir_path
        self._xml_path = None

    def add_schema_search_path(self, dir_path: str) -> None:
        """Add a local directory to look in before the cache or the network.

        A declared schema is matched against these by basename, so a local
        copy of `proced.xsd` satisfies a
        `http://www.s1000d.org/S1000D_4-1/xml_schema_flat/proced.xsd`
        declaration. Search paths are tried in the order they are added.

        Args:
            dir_path (str): directory holding schema files
        """
        if dir_path and dir_path not in self._schema_search_paths:
            self._schema_search_paths.append(dir_path)

    def set_schema_cache(self, dir_path: str) -> None:
        """Choose where fetched schemas are cached, or disable disk caching.

        A fetch pulls the schema's whole `include`/`import` tree into this
        directory, mirroring the URL's path layout, so a later run resolves
        the same schema without touching the network at all. Defaults to
        `~/.acd/schema-cache`.

        Args:
            dir_path (str): cache directory, created on demand; `None`
                disables disk caching (schemas are then resolved over the
                network on every run)
        """
        self._schema_cache_dir = dir_path

    def set_allow_network(self, enabled: bool = True) -> None:
        """Allow or forbid fetching a schema that is not already local.

        With this off, an object whose schema is neither in a search path nor
        in the cache is reported as `XSD-UNRESOLVED` rather than validated --
        which is the honest outcome, and much better than a run that appears
        to pass because nothing was actually checked.

        Args:
            enabled (bool): whether a missing schema may be fetched
        """
        self._allow_network = enabled

    # ------------------------------------------------------------------
    # Layer 1 -- encoding
    # ------------------------------------------------------------------

    @staticmethod
    def _finding(check: str, rule: str, message: str, status: str = "Error",
                 line=None, column=None, element=None, detail=None) -> dict:
        """One finding record. Every check produces these and nothing else, so
        the reports need to understand exactly one shape."""
        return {
            "Check": check,
            "Rule": rule,
            "Status": status,
            "Line": line,
            "Column": column,
            "Message": message,
            "Element": element,
            "Detail": detail,
        }

    def _check_encoding(self, raw: bytes) -> tuple:
        """Check the raw bytes before any parser sees them.

        Args:
            raw (bytes): the file's contents

        Returns:
            tuple: `(findings, fatal)` -- `fatal` is True when the file cannot
                meaningfully be parsed at all, so the caller stops here
        """
        findings = []

        if not raw.strip():
            findings.append(self._finding(
                "encoding", "ENC-EMPTY", CHECK_RULES["ENC-EMPTY"],
                detail=f"{len(raw)} bytes",
            ))
            return findings, True

        body = raw
        if raw.startswith(b"\xef\xbb\xbf"):
            findings.append(self._finding(
                "encoding", "ENC-BOM", CHECK_RULES["ENC-BOM"], status="Warning",
                line=1,
                detail="Harmless to a parser, but it corrupts the XML declaration "
                       "when files are concatenated or included verbatim.",
            ))
            body = raw[3:]

        declared = re_search(rb'''<\?xml[^>]*?encoding\s*=\s*["']([^"']+)["']''', body[:200])
        encoding = declared.group(1).decode("ascii", "replace") if declared else "utf-8"
        try:
            body.decode(encoding)
        except (UnicodeDecodeError, LookupError) as exc:
            findings.append(self._finding(
                "encoding", "ENC-DECL", CHECK_RULES["ENC-DECL"],
                detail=f"declared {encoding!r}: {exc}",
            ))
            return findings, True

        for offset, byte in enumerate(body):
            if byte in _ILLEGAL_CONTROL_BYTES:
                findings.append(self._finding(
                    "encoding", "ENC-CTRL", CHECK_RULES["ENC-CTRL"],
                    line=body.count(b"\n", 0, offset) + 1,
                    detail=f"0x{byte:02X} at byte offset {offset}",
                ))
                # One is enough to make the point and to stop the report
                # filling with a row per byte of a corrupted file.
                break

        return findings, False

    # ------------------------------------------------------------------
    # Layer 2 -- syntax
    # ------------------------------------------------------------------

    def _check_syntax(self, path: str) -> tuple:
        """Parse the object, reporting every entry of the parser's error log.

        External entities and the DTD are deliberately not loaded here: this
        layer answers "is the markup well-formed", and loading a DTD off a
        network share would make that answer depend on the share being up. An
        *undeclared* entity is still reported -- that is a well-formedness
        error whether or not entities are resolved.

        Args:
            path (str): path to the XML file

        Returns:
            tuple: `(tree or None, findings)`
        """
        parser = etree.XMLParser(
            resolve_entities=False, load_dtd=False, no_network=True,
            recover=False, huge_tree=True,
        )
        try:
            return etree.parse(clean_path(path), parser), []
        except etree.XMLSyntaxError as exc:
            entries = list(parser.error_log)
            if not entries:
                return None, [self._finding("syntax", "XML-SYNTAX", str(exc))]
            return None, [
                self._finding(
                    "syntax", "XML-SYNTAX", entry.message,
                    line=entry.line or None, column=entry.column or None,
                )
                for entry in entries
            ]
        except OSError as exc:
            return None, [self._finding(
                "syntax", "XML-SYNTAX", f"File could not be read: {exc}",
            )]

    # ------------------------------------------------------------------
    # Layer 3 -- structure
    # ------------------------------------------------------------------

    def _check_structure(self, path: str, tree, declared_schema: str) -> list:
        """Sanity-check the object against what its filename claims it is.

        Three things, all of which a schema-valid file can still get wrong:
        the root element for its type, the presence of any schema or DOCTYPE
        at all, and -- the one that matters most in a real CSDB -- whether the
        code in the filename matches the code inside the file.

        Args:
            path (str): path to the XML file
            tree: the parsed tree
            declared_schema (str): the object's `xsi:noNamespaceSchemaLocation`,
                or None

        Returns:
            list: findings
        """
        findings = []
        name = basename(str(path))
        prefix = name.split("-")[0].upper()
        root = tree.getroot()

        expected_root = _ROOT_BY_PREFIX.get(prefix)
        if expected_root and root.tag != expected_root:
            findings.append(self._finding(
                "structure", "STR-ROOT", CHECK_RULES["STR-ROOT"], status="Warning",
                line=root.sourceline, element=str(root.tag),
                detail=f"filename says {prefix}, so the root should be "
                       f"<{expected_root}>, but it is <{root.tag}>",
            ))

        if not declared_schema and not tree.docinfo.doctype:
            findings.append(self._finding(
                "structure", "STR-NOSCHEMA", CHECK_RULES["STR-NOSCHEMA"],
                status="Warning", line=root.sourceline,
                detail="Nothing declares how this object should be validated, "
                       "so no schema or DTD check could run.",
            ))

        findings.extend(self._check_ident(name, prefix, tree))
        return findings

    def _check_ident(self, name: str, prefix: str, tree) -> list:
        """Compare the code encoded in the filename with the code element
        inside the object.

        A mismatch means one of the two was edited without the other, which is
        how a data module ends up referencing a BREX, an applicability or a
        publication entry that belongs to a different object entirely.

        Args:
            name (str): the object's filename
            prefix (str): its type prefix (DMC/PMC/DDN/...)
            tree: the parsed tree

        Returns:
            list: at most one finding, listing every field that differs
        """
        if prefix not in _IDENT_ELEMENT:
            return []
        ident_path, element_name, attributes = _IDENT_ELEMENT[prefix]

        root = tree.getroot()
        code_node = root.find(ident_path)
        if code_node is None:
            # An object whose ident section is missing or differently nested
            # (2.x/3.x markup) still has exactly one code element before any
            # reference to another object, so document order stands in.
            code_node = next(iter(root.iter(element_name)), None)
        if code_node is None:
            return []

        try:
            from .s1000d import get_dm_code_from_filename
            from_name = get_dm_code_from_filename(name)
        except (IndexError, ValueError, ImportError):
            from_name = None
        if not from_name:
            return []

        differences = []
        for attribute in attributes:
            expected = from_name.get(attribute)
            found = code_node.get(attribute)
            if expected is None or found is None:
                continue
            if str(expected) != str(found):
                differences.append(f"@{attribute}: filename {expected!r} vs file {found!r}")
        if not differences:
            return []
        return [self._finding(
            "structure", "STR-IDENT", CHECK_RULES["STR-IDENT"],
            line=code_node.sourceline, element=element_name,
            detail="; ".join(differences),
        )]

    # ------------------------------------------------------------------
    # Layer 4 -- DTD
    # ------------------------------------------------------------------

    @staticmethod
    def _has_content_model(tree) -> bool:
        """Whether the object's DOCTYPE actually declares a content model to
        validate against.

        This distinction matters more than it looks. S1000D 4.x content is
        XSD-driven but still carries a DOCTYPE, because that is where the ICN
        graphic entities and the ISO character entity set are declared::

            <!DOCTYPE dmodule [
              <!ENTITY ICN-... SYSTEM "ICN-....cgm" NDATA cgm>
              <!ENTITY % ISOEntities PUBLIC "..." "...">
            ]>

        That internal subset declares entities and nothing else. Handing it to
        a validating parser reports every element and attribute in the
        document as undeclared -- thousands of findings, all of them
        artefacts. A DTD is only worth validating against when it names an
        external DTD file, or when its internal subset declares at least one
        element.

        Args:
            tree: the parsed tree

        Returns:
            bool: whether a DTD validation pass would mean anything
        """
        info = tree.docinfo
        if not info.doctype:
            return False
        if info.system_url or info.public_id:
            return True
        internal = info.internalDTD
        return internal is not None and any(True for _ in internal.iterelements())

    def _check_dtd(self, path: str, tree) -> list:
        """Validate against the DOCTYPE-declared DTD, when there is a content
        model to validate against (see `_has_content_model`).

        S1000D 2.x/3.x content and CMM deliverables are the cases this exists
        for; 4.x+ is XSD-driven and its entity-only DOCTYPE is skipped.

        Args:
            path (str): path to the XML file
            tree: the parsed tree, used only to detect the DOCTYPE

        Returns:
            list: findings
        """
        if not self._has_content_model(tree):
            return []

        parser = etree.XMLParser(
            dtd_validation=True, load_dtd=True, resolve_entities=True,
            no_network=not self._allow_network, recover=True, huge_tree=True,
        )
        try:
            etree.parse(clean_path(path), parser)
        except (etree.XMLSyntaxError, OSError):
            pass

        findings = []
        for entry in parser.error_log:
            # A DTD that could not be loaded at all is a different problem
            # from content that breaks it, and needs a different fix.
            if "failed to load external entity" in entry.message or \
                    "no DTD found" in entry.message:
                findings.append(self._finding(
                    "dtd", "DTD-MISSING", CHECK_RULES["DTD-MISSING"], status="Warning",
                    line=entry.line or None, detail=entry.message,
                ))
            else:
                findings.append(self._finding(
                    "dtd", "DTD-VALID", entry.message,
                    line=entry.line or None, column=entry.column or None,
                ))
        return findings

    # ------------------------------------------------------------------
    # Layer 5 -- schema
    # ------------------------------------------------------------------

    def _cache_path_for(self, url: str) -> str:
        """Where a fetched schema URL lives in the cache, mirroring the URL's
        host and path so its relative `include`s resolve as they would on the
        server."""
        parsed = urlparse(url)
        parts = [_ for _ in parsed.path.split("/") if _ and _ not in (".", "..")]
        return join(self._schema_cache_dir, parsed.netloc, *parts)

    def _fetch_schema_tree(self, url: str, seen: set = None) -> str:
        """Fetch a schema and everything it includes or imports into the cache.

        Fetching only the entry-point schema would leave lxml resolving its
        `include`s back over the network on every run, which is exactly the
        cost the cache exists to remove. Walking the tree once means the
        second run of a folder is fully offline.

        Args:
            url (str): absolute URL of the schema
            seen (set): URLs already fetched in this walk (cycle guard)

        Returns:
            str: local path of the cached entry-point schema

        Raises:
            OSError: if the schema could not be fetched
        """
        seen = seen if seen is not None else set()
        if url in seen:
            return self._cache_path_for(url)
        seen.add(url)

        local = self._cache_path_for(url)
        if not isfile(local):
            with urlopen(url, timeout=30) as response:
                content = response.read()
            makedirs(dirname(local), exist_ok=True)
            with open(clean_path(local), "wb") as cached:
                cached.write(content)
        else:
            with open(clean_path(local), "rb") as cached:
                content = cached.read()

        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError:
            return local
        for node in root.iter():
            tag = etree.QName(node).localname if node.tag is not etree.Comment else ""
            if tag not in ("include", "import", "redefine"):
                continue
            location = node.get("schemaLocation")
            if not location:
                continue
            try:
                self._fetch_schema_tree(urljoin(url, location), seen)
            except OSError:
                # A missing include is the schema author's problem; report it
                # when the schema fails to compile rather than aborting here.
                continue
        return local

    def _resolve_schema(self, declared: str, xml_dir: str) -> tuple:
        """Find the declared schema, in the order a reader would expect:

        1. beside the object being checked, for a relative declaration;
        2. any directory registered with `add_schema_search_path` -- which is
           also where a folder chosen at a prompt lands;
        3. the bundled `res/xsd/<issue>/` tree, found by `resources`, so the
           common case needs no configuration at all;
        4. the on-disk cache of a previous fetch;
        5. the network, if allowed, caching the whole include tree as it goes.

        A location that survives all five is left unresolved and reported --
        `validate` gathers those and asks about them once, at the end of the
        run.

        Args:
            declared (str): the `xsi:noNamespaceSchemaLocation` value
            xml_dir (str): directory of the object, for a relative declaration

        Returns:
            tuple: `(location or None, source)` -- `source` is one of
                `"local"`, `"bundled"`, `"cache"`, `"network"` or
                `"unresolved"`
        """
        if declared in self._resolved_schemas:
            return self._resolved_schemas[declared]

        name = basename(urlparse(declared).path) or basename(declared)
        is_url = urlparse(declared).scheme in ("http", "https")
        resolved = (None, "unresolved")

        if not is_url:
            candidate = join(xml_dir or "", declared)
            if isfile(candidate):
                resolved = (candidate, "local")

        if resolved[0] is None:
            for search_path in self._schema_search_paths:
                candidate = join(search_path, name)
                if isfile(candidate):
                    resolved = (candidate, "local")
                    break

        if resolved[0] is None:
            bundled = find_schema(declared)
            if bundled:
                resolved = (bundled, "bundled")

        if resolved[0] is None and is_url:
            # `_cache_path_for` needs a cache directory to build a path from,
            # so it must not be called at all when caching is switched off.
            cached = self._cache_path_for(declared) if self._schema_cache_dir else None
            if cached and isfile(cached):
                resolved = (cached, "cache")
            elif self._allow_network:
                try:
                    if self._schema_cache_dir:
                        resolved = (self._fetch_schema_tree(declared), "network")
                    else:
                        resolved = (declared, "network")
                except OSError:
                    resolved = (None, "unresolved")

        self._resolved_schemas[declared] = resolved
        return resolved

    def _compiled_schema(self, location: str):
        """Compile a schema once per run.

        Args:
            location (str): local path or URL of the schema

        Returns:
            the compiled `etree.XMLSchema`, or `False` if it would not compile
        """
        if location in self._schema_cache:
            return self._schema_cache[location]
        try:
            compiled = etree.XMLSchema(etree.parse(location))
        except (etree.XMLSchemaParseError, etree.XMLSyntaxError, OSError):
            compiled = False
        self._schema_cache[location] = compiled
        return compiled

    def _check_schema(self, tree, declared: str, xml_dir: str) -> tuple:
        """Validate the object against its declared schema.

        Args:
            tree: the parsed tree
            declared (str): the `xsi:noNamespaceSchemaLocation` value, or None
            xml_dir (str): directory of the object

        Returns:
            tuple: `(findings, schema info dict)`
        """
        info = {"declared": declared, "resolved": None, "source": None}
        if not declared:
            # Already reported by the structure layer as STR-NOSCHEMA; there
            # is nothing further to say here.
            return [], info

        location, source = self._resolve_schema(declared, xml_dir)
        info["resolved"] = location
        info["source"] = source
        if location is None:
            return [self._finding(
                "schema", "XSD-UNRESOLVED", CHECK_RULES["XSD-UNRESOLVED"],
                status="Warning", detail=(
                    f"{declared} -- not in any search path or the cache"
                    + ("" if self._allow_network else ", and network access is off")
                ),
            )], info

        schema = self._compiled_schema(location)
        if schema is False:
            return [self._finding(
                "schema", "XSD-BROKEN", CHECK_RULES["XSD-BROKEN"], status="Warning",
                detail=f"{location} could not be compiled as an XML Schema",
            )], info

        if schema.validate(tree):
            return [], info
        return [
            self._finding(
                "schema", "XSD-VALID", entry.message,
                line=entry.line or None, column=entry.column or None,
                element=entry.path,
            )
            for entry in schema.error_log
        ], info

    # ------------------------------------------------------------------
    # Driver
    # ------------------------------------------------------------------

    def _check_object(self, path: str, check_encoding: bool, check_syntax: bool,
                      check_structure: bool, check_dtd: bool,
                      check_schema: bool) -> dict:
        """Run every enabled layer over one object, stopping at the first that
        makes the next one meaningless."""
        findings = []
        schema_info = {"declared": None, "resolved": None, "source": None}

        try:
            with open(clean_path(path), "rb") as handle:
                raw = handle.read()
        except OSError as exc:
            findings.append(self._finding(
                "syntax", "XML-SYNTAX", f"File could not be read: {exc}",
            ))
            return self._object_result(path, findings, schema_info)

        if check_encoding:
            encoding_findings, fatal = self._check_encoding(raw)
            findings.extend(encoding_findings)
            if fatal:
                return self._object_result(path, findings, schema_info)

        if not check_syntax:
            return self._object_result(path, findings, schema_info)

        tree, syntax_findings = self._check_syntax(path)
        findings.extend(syntax_findings)
        if tree is None:
            # Nothing below this layer can say anything true about a document
            # that does not parse.
            return self._object_result(path, findings, schema_info)

        declared = get_schema_from_xml(raw)

        if check_structure:
            findings.extend(self._check_structure(path, tree, declared))
        if check_dtd:
            findings.extend(self._check_dtd(path, tree))
        if check_schema:
            schema_findings, schema_info = self._check_schema(
                tree, declared, dirname(str(path))
            )
            findings.extend(schema_findings)

        return self._object_result(path, findings, schema_info)

    @staticmethod
    def _object_result(path: str, findings: list, schema_info: dict) -> dict:
        """Wrap one object's findings with its own totals, so a per-document
        summary never has to re-walk the findings list."""
        errors = sum(1 for _ in findings if _["Status"] == "Error")
        warnings = len(findings) - errors
        return {
            "Document": str(path),
            "Findings": findings,
            "Schema": schema_info,
            "Summary": {
                "Errors": errors,
                "Warnings": warnings,
                "Status": "Failed" if errors else "Passed",
            },
        }

    def validate(self, check_encoding: bool = True, check_syntax: bool = True,
                 check_structure: bool = True, check_dtd: bool = True,
                 check_schema: bool = True, progress_callback=None) -> dict:
        """Check the object set by `set_xml`, or every `.xml` file in the
        directory set by `set_xml_dir`.

        Layers run in order and a layer that cannot produce a truthful answer
        is skipped: an unreadable or undecodable file is not parsed, and a
        file that is not well-formed is not structure-, DTD- or
        schema-checked.

        Args:
            check_encoding (bool): run the byte-level layer (empty file, BOM,
                declared encoding, illegal control characters)
            check_syntax (bool): run the well-formedness layer. With this off
                nothing below it runs either, since they all need a tree.
            check_structure (bool): run the root-element, schema-declaration
                and filename-vs-ident layer
            check_dtd (bool): validate against the DOCTYPE-declared DTD, if any
            check_schema (bool): validate against
                `xsi:noNamespaceSchemaLocation`, if declared
            progress_callback (callable): optional
                `progress_callback(current, total, stage)`, called once per
                file with `stage="files"` in directory mode

        Returns:
            dict: in single-object mode, one object result --
                `{"Document", "Findings", "Schema", "Summary"}`; in directory
                mode, `{filename: object result}`

        Raises:
            ValueError: if neither `set_xml` nor `set_xml_dir` was called
        """
        layers = (check_encoding, check_syntax, check_structure, check_dtd, check_schema)

        if self._xml_dir:
            files = [
                _ for _ in listdir(self._xml_dir)
                if _.lower().endswith(".xml") and isfile(join(self._xml_dir, _))
            ]
            results = {}
            for index, name in enumerate(sorted(files)):
                results[name] = self._check_object(join(self._xml_dir, name), *layers)
                if progress_callback is not None:
                    progress_callback(index + 1, len(files), "files")
            return self._retry_unresolved_schemas(results, layers)

        if not self._xml_path:
            raise ValueError(
                "nothing to check: call set_xml(path) or set_xml_dir(path) first"
            )
        result = self._check_object(self._xml_path, *layers)
        return self._retry_unresolved_schemas(result, layers)

    def unresolved_schemas(self, result: dict) -> list:
        """Every declared schema a run could not resolve, once each.

        Args:
            result (dict): a `validate()` return value

        Returns:
            list: declared locations, sorted, empty when everything resolved
        """
        missing = {
            document["Schema"]["declared"]
            for document in self._result_documents(result).values()
            if document["Schema"].get("source") == "unresolved"
            and document["Schema"].get("declared")
        }
        return sorted(missing)

    def _retry_unresolved_schemas(self, result: dict, layers: tuple) -> dict:
        """Ask once for a folder holding whatever schema could not be found,
        then re-check only the objects that were affected.

        This is the last line of defence, and it is deliberately last: by the
        time it runs, the object's own folder, every registered search path,
        the bundled `res/xsd` tree, the cache and the network have all been
        tried. Asking beats reporting a run as clean when nothing was actually
        validated.

        One prompt covers the whole run -- a folder of ninety objects sharing
        one missing schema has one thing wrong with it -- and a declined or
        unavailable prompt simply leaves the findings as they are, so a batch
        job behaves exactly as it did before (see `prompt.prompting_enabled`).

        Args:
            result (dict): the run's result, single-object or directory-mode
            layers (tuple): the check flags the run was made with

        Returns:
            dict: the result, with re-checked objects replaced
        """
        missing = self.unresolved_schemas(result)
        if not missing:
            return result

        folder = ask_for_folder("schemas", missing)
        if not folder:
            return result

        self.add_schema_search_path(folder)
        # The failed lookups are cached as failures; drop them so the new
        # search path is actually consulted.
        self._resolved_schemas = {
            key: value for key, value in self._resolved_schemas.items()
            if value[1] != "unresolved"
        }

        if self._is_single_object_result(result):
            return self._check_object(self._xml_path, *layers)
        for name, document in list(result.items()):
            if document["Schema"].get("source") == "unresolved":
                result[name] = self._check_object(
                    join(self._xml_dir, name), *layers
                )
        return result

    # ------------------------------------------------------------------
    # Results
    # ------------------------------------------------------------------

    @staticmethod
    def _is_single_object_result(result: dict) -> bool:
        """Tell a single-object result from a directory-mode one, the way
        `BrexChecker` does: only the former carries its own `Findings`."""
        return "Findings" in result

    def _result_documents(self, result: dict) -> dict:
        """Normalise either result shape to `{name: object result}`."""
        if self._is_single_object_result(result):
            return {basename(str(result.get("Document") or "")): result}
        return {
            name: value for name, value in result.items()
            if isinstance(value, dict) and self._is_single_object_result(value)
        }

    def findings(self, result: dict) -> list:
        """Every finding of a run, flattened and tagged with its document --
        the row set both formatted reports are built from.

        Args:
            result (dict): a `validate()` return value

        Returns:
            list: dicts keyed by the report's column labels, in document then
                layer order
        """
        rows = []
        for name, document in self._result_documents(result).items():
            for finding in document["Findings"]:
                rows.append({
                    "Document": name,
                    "Check": finding["Check"],
                    "Rule": finding["Rule"],
                    "Status": finding["Status"],
                    "Line": finding["Line"],
                    "Column": finding["Column"],
                    "Message": finding["Message"],
                    "Element": finding["Element"],
                    "Detail": finding["Detail"],
                })
        rows.sort(key=lambda row: (
            row["Document"], CHECK_LAYERS.index(row["Check"]), row["Line"] or 0
        ))
        return rows

    def document_stats(self, result: dict) -> list:
        """Per-document pass/fail tallies, the counterpart of `run_summary`'s
        run-wide totals.

        Args:
            result (dict): a `validate()` return value

        Returns:
            list: `{"document", "errors", "warnings", "status", "schema"}` dicts
        """
        rows = []
        for name, document in self._result_documents(result).items():
            summary = document["Summary"]
            rows.append({
                "document": name,
                "errors": summary["Errors"],
                "warnings": summary["Warnings"],
                "status": summary["Status"],
                "schema": document["Schema"].get("declared"),
            })
        return rows

    def run_summary(self, result: dict) -> dict:
        """Run-wide totals, the same shape `BrexChecker.run_summary` returns
        plus a per-layer breakdown.

        Args:
            result (dict): a `validate()` return value

        Returns:
            dict::

                {
                    "DocumentsChecked": int,
                    "DocumentsPassed": int,
                    "DocumentsFailed": int,
                    "DocumentsSkipped": int,
                    "Errors": int,
                    "Warnings": int,
                    "FindingsByCheck": dict,  # layer -> count
                    "FindingsByRule": dict,   # rule code -> count
                }
        """
        documents = self._result_documents(result)
        errors = warnings = passed = failed = 0
        by_check = {}
        by_rule = {}
        for document in documents.values():
            summary = document["Summary"]
            errors += summary["Errors"]
            warnings += summary["Warnings"]
            if summary["Status"] == "Failed":
                failed += 1
            else:
                passed += 1
            for finding in document["Findings"]:
                by_check[finding["Check"]] = by_check.get(finding["Check"], 0) + 1
                by_rule[finding["Rule"]] = by_rule.get(finding["Rule"], 0) + 1
        return {
            "DocumentsChecked": len(documents),
            "DocumentsPassed": passed,
            "DocumentsFailed": failed,
            "DocumentsSkipped": 0,
            "Errors": errors,
            "Warnings": warnings,
            "FindingsByCheck": by_check,
            "FindingsByRule": by_rule,
        }

    def _report_source(self) -> str:
        """What was checked, for a report header."""
        return self._xml_dir or self._xml_path or ""

    # ------------------------------------------------------------------
    # Reports
    # ------------------------------------------------------------------

    def to_json_report(self, result: dict, path: str = None) -> str:
        """Serialise a run as JSON: the run summary, then one entry per
        document with its schema resolution and findings.

        Args:
            result (dict): a `validate()` return value
            path (str): optional destination, written as UTF-8

        Returns:
            str: the JSON document
        """
        payload = {
            "source": self._report_source(),
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": self.run_summary(result),
            "documents": [
                {
                    "document": name,
                    "status": document["Summary"]["Status"],
                    "errors": document["Summary"]["Errors"],
                    "warnings": document["Summary"]["Warnings"],
                    "schema": document["Schema"],
                    "findings": document["Findings"],
                }
                for name, document in self._result_documents(result).items()
            ],
        }
        text = dumps(payload, indent=4)
        if path:
            with open(clean_path(path), "w", encoding="utf-8") as report_file:
                report_file.write(text)
        return text

    def to_excel_report(self, result: dict, path: str) -> str:
        """Convert a run into a formatted Excel workbook, laid out like
        `BrexChecker.to_excel_report`'s.

        Sheets: **Summary** (run totals, findings by layer and by rule, and a
        per-document table), **Findings** (one row per finding), and
        **Schemas** (what each declared schema resolved to and where it came
        from -- the sheet that shows whether a "clean" run actually validated
        anything).

        Args:
            result (dict): a `validate()` return value
            path (str): destination `.xlsx` path; parent directory must exist

        Raises:
            ImportError: if `openpyxl` is not installed

        Returns:
            str: `path`, for convenience
        """
        report = ExcelReport("XML check report", self._report_source())
        palette = report.palette

        totals = self.run_summary(result)
        finding_rows = self.findings(result)
        document_rows = self.document_stats(result)

        row = report.write_totals(4, [
            ("Documents checked", totals["DocumentsChecked"], None),
            ("Documents passed", totals["DocumentsPassed"], "ok"),
            ("Documents failed", totals["DocumentsFailed"], "error"),
            ("Errors", totals["Errors"], "error"),
            ("Warnings", totals["Warnings"], "warning"),
        ])

        row += 1
        row = report.label(row, "Findings by check")
        row = report.write_table(
            report.summary, ["Check", "Count"],
            [
                {"Check": check, "Count": totals["FindingsByCheck"][check]}
                for check in CHECK_LAYERS if check in totals["FindingsByCheck"]
            ] or [{"Check": "(none)", "Count": 0}],
            [34, 16], start_row=row, autofilter=False, freeze=False,
        )

        row = report.label(row, "Findings by rule")
        row = report.write_table(
            report.summary, ["Rule", "Count", "Meaning"],
            [
                {"Rule": rule, "Count": count, "Meaning": CHECK_RULES.get(rule, "")}
                for rule, count in sorted(totals["FindingsByRule"].items())
            ] or [{"Rule": "(none)", "Count": 0, "Meaning": ""}],
            [34, 16, 80], start_row=row, autofilter=False, freeze=False,
        )

        row = report.label(row, "Documents")
        report.write_table(
            report.summary,
            ["Document", "Status", "Errors", "Warnings", "Schema"],
            [
                {
                    "Document": entry["document"],
                    "Status": entry["status"],
                    "Errors": entry["errors"],
                    "Warnings": entry["warnings"],
                    "Schema": entry["schema"],
                }
                for entry in document_rows
            ],
            [52, 14, 12, 12, 60],
            tint=lambda entry: (
                "error" if entry["Status"] == "Failed"
                else "warning" if entry["Warnings"] else "ok"
            ),
            start_row=row, freeze=False,
        )

        report.write_table(
            report.add_sheet("Findings", palette["error"]),
            ["Document", "Check", "Rule", "Status", "Line", "Column", "Message",
             "Element", "Detail"],
            finding_rows,
            [30, 12, 16, 10, 8, 8, 70, 30, 60],
            tint=lambda entry: "error" if entry["Status"] == "Error" else "warning",
        )

        schema_rows = self._schema_rows(result)
        if schema_rows:
            report.write_table(
                report.add_sheet("Schemas", palette["muted"]),
                ["Declared", "Resolved", "Source", "Documents"],
                schema_rows,
                [64, 64, 14, 14],
            )

        return report.save(path)

    def _schema_rows(self, result: dict) -> list:
        """One row per distinct declared schema: what it resolved to, where
        that came from, and how many objects declared it.

        This is the antidote to a run that looks clean because nothing was
        checked -- a `source` of `unresolved` says so in one glance.

        Args:
            result (dict): a `validate()` return value

        Returns:
            list: dicts keyed by the Schemas sheet's column labels
        """
        rows = {}
        for document in self._result_documents(result).values():
            schema = document["Schema"]
            declared = schema.get("declared")
            if not declared:
                continue
            entry = rows.setdefault(declared, {
                "Declared": declared,
                "Resolved": schema.get("resolved"),
                "Source": schema.get("source") or "unresolved",
                "Documents": 0,
            })
            entry["Documents"] += 1
        return [rows[key] for key in sorted(rows)]

    def to_html_report(self, result: dict, path: str = None,
                       title: str = "XML check report") -> str:
        """Convert a run into a self-contained HTML report -- the
        browser-readable counterpart of `to_excel_report`, built from the same
        rows and the same `acd.report` layer as the BREX report, so the two
        read as one document set.

        Args:
            result (dict): a `validate()` return value
            path (str): optional destination, written as UTF-8
            title (str): heading shown at the top of the report

        Returns:
            str: the complete HTML document
        """
        totals = self.run_summary(result)
        finding_rows = self.findings(result)
        document_rows = self.document_stats(result)

        report = HtmlReport(title, self._report_source())
        cell = report.cell
        table = report.table
        section = report.section

        report.add_cards([
            ("Documents checked", totals["DocumentsChecked"], "neutral"),
            ("Passed", totals["DocumentsPassed"], "ok" if totals["DocumentsPassed"] else "neutral"),
            ("Failed", totals["DocumentsFailed"], "error" if totals["DocumentsFailed"] else "neutral"),
            ("Errors", totals["Errors"], "error" if totals["Errors"] else "ok"),
            ("Warnings", totals["Warnings"], "warning" if totals["Warnings"] else "neutral"),
        ])
        report.add_chips([
            (check, totals["FindingsByCheck"][check])
            for check in CHECK_LAYERS if check in totals["FindingsByCheck"]
        ])

        if len(document_rows) > 1:
            report.add(section(
                "Documents", len(document_rows),
                table(
                    ["Document", "Status", "Errors", "Warnings", "Schema"],
                    [
                        cell(entry["document"], mono=True)
                        + report.status_cell(entry["status"])
                        + cell(entry["errors"], "num")
                        + cell(entry["warnings"], "num")
                        + cell(entry["schema"], mono=True)
                        for entry in document_rows
                    ],
                ),
            ))

        if finding_rows:
            body_rows = []
            for entry in finding_rows:
                details = []
                if entry["Detail"]:
                    details.append(
                        f'<p class="muted">{report.escape(entry["Detail"])}</p>'
                    )
                if entry["Element"]:
                    details.append(
                        f'<p><span class="muted">at:</span> '
                        f'<code>{report.escape(str(entry["Element"]))}</code></p>'
                    )
                body_rows.append(
                    f'<tr data-status="{entry["Status"].lower()}">'
                    + cell(entry["Document"], mono=True)
                    + cell(entry["Check"])
                    + cell(entry["Rule"], mono=True)
                    + report.status_cell(entry["Status"])
                    + cell(entry["Line"], "num")
                    + cell(entry["Message"])
                    + f'<td class="details">{"".join(details)}</td>'
                    + "</tr>"
                )
            report.add(section(
                "Findings", len(finding_rows),
                report.filter_controls(
                    "Filter findings (document, rule, message...)",
                    "finding-filter", "errors-only", "finding-count",
                )
                + table(
                    ("Document", "Check", "Rule", "Status", "Line", "Message", "Details"),
                    body_rows,
                    table_attrs='id="findings" data-filterable data-noun="findings"',
                ),
            ))
        else:
            report.add_empty_state("No syntax or schema problems found.")

        schema_rows = self._schema_rows(result)
        if schema_rows:
            report.add(section(
                "Schemas", len(schema_rows),
                table(
                    ["Declared", "Resolved", "Source", "Documents"],
                    [
                        cell(entry["Declared"], mono=True)
                        + cell(entry["Resolved"], mono=True)
                        + cell(entry["Source"])
                        + cell(entry["Documents"], "num")
                        for entry in schema_rows
                    ],
                ),
                open_by_default=False,
            ))

        return report.render(path)
