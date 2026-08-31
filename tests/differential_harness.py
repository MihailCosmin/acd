"""Differential harness comparing `BrexChecker` against the reference
`s1kd-brexcheck` C tool (brex_checker_rework.md §4.5), for the same real
S1000D CSDB objects/BREX.

Both tools can emit an XML report in the same `brexCheck` shape --
`BrexChecker.to_xml_report` (category D6) was deliberately built compatible
with `s1kd-brexcheck -x`'s documented shape (see its man page's EXAMPLE
section). `parse_brex_check_xml_report` is the single parser used for both
sides, so the two tools' output goes through identical extraction logic
before being compared: only the underlying check results can differ, not
how the harness reads them.

This module works, and is unit-tested (see test_differential_harness.py),
with no `s1kd-brexcheck` binary present at all -- only `run_s1kd_brexcheck`
(and the one integration test that calls it) needs a real build. Not a
`test_*.py` module itself so pytest does not try to collect it directly.
"""

from dataclasses import dataclass
from os.path import basename
from os import environ
from shutil import which
from subprocess import run
from typing import Optional

from lxml import etree


# XPath 2.0-only functions used by 6 of the 919 real rules in the plan
# document's evidence base (brex_checker_rework.md §1, "XPath engine
# comparison"): a *default* s1kd-brexcheck build (no XPath 2.0 engine
# compiled in -- `xpath2_engine=NONE` in the tool's own Makefile) cannot
# compile these, so a rule whose objectPath contains one of them is an
# expected, documented divergence, not a real parity gap.
KNOWN_DIVERGENT_XPATH_MARKERS = ("matches(", "tokenize(", "lower-case(")


@dataclass(frozen=True)
class ParsedViolation:
    """One violation record read out of a `brexCheck` XML report, either
    ours (`BrexChecker.to_xml_report`) or the reference tool's (`s1kd-brexcheck -x`).
    """
    document: str
    brex: str
    flag: Optional[str]
    object_path: Optional[str]
    line: Optional[str]
    node_xpath: Optional[str]

    def key(self, include_brex: bool = True) -> tuple:
        """Normalized identity for cross-tool comparison: which rule fired,
        against which document (and, by default, from which BREX file).

        Deliberately excludes `line`/`node_xpath`: our canonical-XPath/
        line-number computation (elementpath) and s1kd's (libxml2's
        `xmlGetLineNo`/its own `xpath_of`) are independent implementations,
        not expected to produce byte-identical strings even for the exact
        same real violation.

        `include_brex=False` is available because a BREX layer resolved via
        each tool's own built-in-default fallback may be reported under a
        different path/label by the two tools (ours: our bundled `acd/brex/`
        copy's path; s1kd's: whatever its compiled-in `brex.h` table
        reports) even when the rule content is identical -- if a real
        differential run shows spurious-looking divergence concentrated on
        one BREX layer, re-diffing with `include_brex=False` separates "the
        rule itself disagrees" from "only the BREX label differs".

        Args:
            include_brex (bool): include the (basename of the) BREX file
                path in the key

        Returns:
            tuple: comparison key
        """
        parts = (basename(self.document), self.flag, self.object_path)
        if include_brex:
            return (basename(self.brex),) + parts
        return parts


def find_s1kd_brexcheck() -> Optional[str]:
    """Locate an `s1kd-brexcheck` executable.

    Checks the `S1KD_BREXCHECK_BIN` environment variable first (a full path,
    for a locally-built binary not on `PATH`), then falls back to a `PATH`
    lookup. Returns `None` if neither resolves -- the expected case on a
    machine that has not built the reference C tool (it is not vendored or
    compiled as part of this package; see
    `s1kd-tools-master/tools/s1kd-brexcheck/Makefile`, which needs
    `libxml2`/`libxslt`/`libexslt` development headers).

    Returns:
        Optional[str]: path to the executable, or None
    """
    override = environ.get("S1KD_BREXCHECK_BIN")
    if override:
        return override
    return which("s1kd-brexcheck") or which("s1kd-brexcheck.exe")


def parse_brex_check_xml_report(xml_text) -> list:
    """Parse a `brexCheck` XML report into a flat list of `ParsedViolation`.

    Handles both our own `BrexChecker.to_xml_report` output and a real
    `s1kd-brexcheck -x` report -- deliberately compatible shapes (category
    D6). An `<error>` with several `<object>` children (s1kd groups every
    node matched by the same rule under one `<error>`; we always emit at
    most one `<object>` per `<error>`/violation record, see
    `BrexChecker._append_error_node`) is split into one `ParsedViolation`
    per `<object>`, so both shapes normalize to "one record per matched
    node". An `<error>` with no `<object>` at all (a flag-1 "required but
    missing" violation, or a boolean-valued flag-0 result) becomes one
    record with `line`/`node_xpath` both `None`.

    Args:
        xml_text: a `brexCheck` XML report, as `str` or `bytes`

    Returns:
        list: `ParsedViolation` records, in document order
    """
    data = xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text
    root = etree.fromstring(data)
    records = []
    for document_node in root.findall(".//document"):
        docname = document_node.get("path", "")
        for brex_node in document_node.findall("brex"):
            brex_path = brex_node.get("path", "")
            for error_node in brex_node.findall("error"):
                object_path_node = error_node.find("objectPath")
                flag = object_path_node.get("allowedObjectFlag") if object_path_node is not None else None
                object_path = object_path_node.text if object_path_node is not None else None
                object_nodes = error_node.findall("object")
                if object_nodes:
                    for object_node in object_nodes:
                        records.append(ParsedViolation(
                            document=docname, brex=brex_path, flag=flag,
                            object_path=object_path,
                            line=object_node.get("line"),
                            node_xpath=object_node.get("xpath"),
                        ))
                else:
                    records.append(ParsedViolation(
                        document=docname, brex=brex_path, flag=flag,
                        object_path=object_path, line=None, node_xpath=None,
                    ))
    return records


def run_s1kd_brexcheck(binary: str, object_paths: list, brex_dir: str = None,
                        extra_args: list = None, timeout: int = 900) -> str:
    """Run the reference `s1kd-brexcheck` tool and return its XML report.

    Runs the checklist's literal `s1kd-brexcheck -clnST` invocation (`-c`
    object-value checking, `-l` layered BREX, `-n` notation rules, `-S` SNS,
    `-T` summary), with `-x` added so the output is the same structured
    `brexCheck` XML both `parse_brex_check_xml_report` and our own
    `to_xml_report` produce. `-clnST` alone only controls which *checks*
    run; `-x` only changes how the result is *reported* -- adding it changes
    nothing about what is being compared.

    Args:
        binary (str): path to the `s1kd-brexcheck` executable (see `find_s1kd_brexcheck`)
        object_paths (list): CSDB object files to check
        brex_dir (str): `-d` search directory for referenced BREX data modules
        extra_args (list): additional CLI flags, e.g. `["-B"]`
        timeout (int): subprocess timeout in seconds

    Returns:
        str: stdout (the XML report)
    """
    args = [binary, "-c", "-l", "-n", "-S", "-T", "-x"]
    if brex_dir:
        args += ["-d", brex_dir]
    if extra_args:
        args += list(extra_args)
    args += list(object_paths)
    completed = run(args, capture_output=True, text=True, timeout=timeout, check=False)
    return completed.stdout


def diff_violation_sets(ours: list, theirs: list,
                         ignore_markers: tuple = KNOWN_DIVERGENT_XPATH_MARKERS,
                         include_brex: bool = True) -> dict:
    """Diff two `ParsedViolation` lists by their normalized `.key()`.

    A key present on only one side is classified as an intentional,
    documented divergence (`ignored_*`) when its `object_path` contains one
    of `ignore_markers`, otherwise as a real parity gap (`only_*`).

    Args:
        ours (list): `ParsedViolation` records from our own report
        theirs (list): `ParsedViolation` records from the reference tool's report
        ignore_markers (tuple): substrings of `object_path` that mark a
            known, documented divergence rather than a parity gap
        include_brex (bool): forwarded to `ParsedViolation.key`

    Returns:
        dict: `{"common", "only_ours", "only_theirs", "ignored_ours",
            "ignored_theirs"}`, each a set of `.key()` tuples
    """
    def is_ignored(key):
        object_path = key[-1] or ""
        return any(marker in object_path for marker in ignore_markers)

    our_keys = {v.key(include_brex) for v in ours}
    their_keys = {v.key(include_brex) for v in theirs}

    only_ours_raw = our_keys - their_keys
    only_theirs_raw = their_keys - our_keys

    return {
        "common": our_keys & their_keys,
        "only_ours": {k for k in only_ours_raw if not is_ignored(k)},
        "only_theirs": {k for k in only_theirs_raw if not is_ignored(k)},
        "ignored_ours": {k for k in only_ours_raw if is_ignored(k)},
        "ignored_theirs": {k for k in only_theirs_raw if is_ignored(k)},
    }
