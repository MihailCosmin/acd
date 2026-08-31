"""Differential harness tests for brex_checker_rework.md §4.5: "Add a
differential harness that runs both brex_checker.py and
s1kd-brexcheck -clnST over the whole CMP 21-77-05 folder and diffs the
violation sets, so parity is measurable and intentional divergences are
recorded."

The parsing/diffing logic (differential_harness.py) is real and fully unit
tested below with no external binary needed. The actual cross-tool run
(test_differential_parity_over_the_evidence_folder) additionally needs a
real `s1kd-brexcheck` build, which this environment does not have (no
prebuilt binary, and building it needs libxml2/libxslt/libexslt development
headers this machine's mingw64 toolchain does not have either -- see
`differential_harness.py`'s module docstring) -- it is skipped rather than
faked, and will actually run wherever both the binary and the evidence
folder are available (set S1KD_BREXCHECK_BIN / ACD_BREX_EVIDENCE_DIR).
"""

import os
from os.path import isdir, isfile, join

import pytest

from acd.brex_checker import BrexChecker
from differential_harness import (
    ParsedViolation,
    diff_violation_sets,
    find_s1kd_brexcheck,
    parse_brex_check_xml_report,
    run_s1kd_brexcheck,
)

SAMPLE_REPORT = """<?xml version="1.0"?>
<brexCheck>
<document path="DMC-EX-A-00-00-00-00A-040A-D_000-01_EN-CA.XML">
<brex path="DMC-S1000D-G-04-10-0301-00A-022A-D_001-00_EN-US.XML">
<error fail="yes">
<brDecisionRef brDecisionIdentNumber="BREX-S1-00052"/>
<objectPath allowedObjectFlag="0">//internalRef[@internalRefTargetType != 'irtt08']</objectPath>
<objectUse>Only when the reference target is a step...</objectUse>
<object line="52" xpath="/dmodule[1]/content[1]/description[1]/para[2]/internalRef[1]">
<internalRef internalRefTargetType="irtt08" internalRefId="stp-0001"/>
</object>
</error>
<error fail="yes">
<objectPath allowedObjectFlag="1">//missingRequired</objectPath>
<objectUse>missingRequired must be present.</objectUse>
</error>
</brex>
</document>
</brexCheck>
"""


# ---------------------------------------------------------------------------
# parse_brex_check_xml_report -- no binary needed
# ---------------------------------------------------------------------------

def test_parse_reads_the_documented_report_shape():
    records = parse_brex_check_xml_report(SAMPLE_REPORT)

    assert len(records) == 2
    with_object = next(r for r in records if r.flag == "0")
    assert with_object.document == "DMC-EX-A-00-00-00-00A-040A-D_000-01_EN-CA.XML"
    assert with_object.brex == "DMC-S1000D-G-04-10-0301-00A-022A-D_001-00_EN-US.XML"
    assert with_object.object_path == "//internalRef[@internalRefTargetType != 'irtt08']"
    assert with_object.line == "52"
    assert with_object.node_xpath == "/dmodule[1]/content[1]/description[1]/para[2]/internalRef[1]"

    without_object = next(r for r in records if r.flag == "1")
    assert without_object.line is None
    assert without_object.node_xpath is None


def test_parse_splits_one_error_with_multiple_objects_into_one_record_each():
    # s1kd-brexcheck groups every node matched by the same rule under one
    # <error> with several <object> children; our own to_xml_report always
    # emits at most one <object> per <error>. The parser must normalize
    # both shapes the same way so key()-based comparison is meaningful
    # regardless of which tool produced the report.
    xml_text = """<brexCheck><document path="d.xml"><brex path="b.xml">
    <error fail="yes">
    <objectPath allowedObjectFlag="0">//forbidden</objectPath>
    <objectUse>forbidden must not be present.</objectUse>
    <object line="10" xpath="/root[1]/forbidden[1]"><forbidden/></object>
    <object line="20" xpath="/root[1]/forbidden[2]"><forbidden/></object>
    </error>
    </brex></document></brexCheck>"""

    records = parse_brex_check_xml_report(xml_text)

    assert len(records) == 2
    assert {r.line for r in records} == {"10", "20"}
    assert all(r.object_path == "//forbidden" for r in records)


def test_our_own_to_xml_report_round_trips_through_the_shared_parser(tmp_path):
    # Confirms the "same parser for both sides" design against our own real
    # to_xml_report output, not just a hand-written sample -- the only half
    # of the harness fully verifiable without the reference binary.
    brex_content = """<brex>
    <contextRules><structureObjectRuleGroup>
    <structureObjectRule id="SOR-1">
    <objectPath allowedObjectFlag="0">//forbidden</objectPath>
    <objectUse>forbidden must not be present.</objectUse>
    </structureObjectRule>
    </structureObjectRuleGroup></contextRules>
    </brex>
    """
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text("<root><forbidden/></root>\n", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(brex_path)])
    result = checker.validate()

    records = parse_brex_check_xml_report(checker.to_xml_report(result))

    assert len(records) == 1
    assert records[0].object_path == "//forbidden"
    assert records[0].flag == "0"


# ---------------------------------------------------------------------------
# diff_violation_sets -- no binary needed
# ---------------------------------------------------------------------------

def test_diff_matches_identical_sets():
    v = ParsedViolation("d.xml", "b.xml", "0", "//foo", "1", "/root[1]/foo[1]")
    diff = diff_violation_sets([v], [v])
    assert diff["common"] == {v.key()}
    assert not diff["only_ours"] and not diff["only_theirs"]


def test_diff_reports_real_gaps_on_both_sides():
    ours_only = ParsedViolation("d.xml", "b.xml", "0", "//onlyOurs", None, None)
    theirs_only = ParsedViolation("d.xml", "b.xml", "0", "//onlyTheirs", None, None)

    diff = diff_violation_sets([ours_only], [theirs_only])

    assert diff["only_ours"] == {ours_only.key()}
    assert diff["only_theirs"] == {theirs_only.key()}
    assert not diff["ignored_ours"] and not diff["ignored_theirs"]


def test_diff_classifies_known_xpath2_only_functions_as_ignored():
    # Ref brex_checker_rework.md §1: 6 of 919 real rules use matches()/
    # tokenize()/lower-case(), which a default (XPath-1-only) s1kd-brexcheck
    # build cannot compile -- an expected, documented divergence rather
    # than a real parity gap.
    xpath2_only = ParsedViolation("d.xml", "b.xml", "0", "//x[matches(., '[0-9]+')]", None, None)

    diff = diff_violation_sets([xpath2_only], [])

    assert diff["only_ours"] == set()
    assert diff["ignored_ours"] == {xpath2_only.key()}


def test_diff_ignore_markers_can_be_overridden():
    v = ParsedViolation("d.xml", "b.xml", "0", "//custom-marker-rule", None, None)
    diff = diff_violation_sets([v], [], ignore_markers=("custom-marker",))
    assert diff["ignored_ours"] == {v.key()}
    assert diff["only_ours"] == set()


def test_diff_can_ignore_the_brex_label_and_still_match():
    # A BREX layer resolved via each tool's own built-in-default fallback
    # may be reported under a different path by the two tools even when the
    # rule content is identical -- include_brex=False isolates that case.
    ours = ParsedViolation("d.xml", "C:/ours/acd/brex/DMC-S1000D-F.XML", "0", "//foo", None, None)
    theirs = ParsedViolation("d.xml", "S1000D_default_F", "0", "//foo", None, None)

    assert diff_violation_sets([ours], [theirs])["only_ours"] == {ours.key()}
    diff_ignoring_brex = diff_violation_sets([ours], [theirs], include_brex=False)
    assert diff_ignoring_brex["only_ours"] == set()
    assert diff_ignoring_brex["common"] == {ours.key(include_brex=False)}


# ---------------------------------------------------------------------------
# find_s1kd_brexcheck -- no binary needed
# ---------------------------------------------------------------------------

def test_find_s1kd_brexcheck_honours_explicit_env_override(monkeypatch, tmp_path):
    fake_binary = tmp_path / "s1kd-brexcheck"
    fake_binary.write_text("", encoding="utf-8")
    monkeypatch.setenv("S1KD_BREXCHECK_BIN", str(fake_binary))
    assert find_s1kd_brexcheck() == str(fake_binary)


def test_find_s1kd_brexcheck_returns_none_when_not_available(monkeypatch):
    monkeypatch.delenv("S1KD_BREXCHECK_BIN", raising=False)
    monkeypatch.setattr("differential_harness.which", lambda name: None)
    assert find_s1kd_brexcheck() is None


# ---------------------------------------------------------------------------
# Real integration: actually shells out to s1kd-brexcheck over the full
# CMP 21-77-05 evidence folder and diffs the violation sets. Skipped
# whenever either the reference binary or the evidence folder is
# unavailable -- expected on most machines/CI (see module docstring).
# ---------------------------------------------------------------------------
EVIDENCE_DIR = os.environ.get(
    "ACD_BREX_EVIDENCE_DIR",
    r"C:\Users\munte\Develop\TD\SITEC\Seventh Delivery\CMP 21-77-05",
)
S1KD_BREXCHECK_BIN = find_s1kd_brexcheck()

requires_differential_fixtures = pytest.mark.skipif(
    not (S1KD_BREXCHECK_BIN and isdir(EVIDENCE_DIR)),
    reason="needs both a real s1kd-brexcheck binary (set S1KD_BREXCHECK_BIN) "
           "and the CMP 21-77-05 evidence folder (set ACD_BREX_EVIDENCE_DIR)",
)


@requires_differential_fixtures
@pytest.mark.evidence
def test_differential_parity_over_the_evidence_folder():
    object_paths = sorted(
        join(EVIDENCE_DIR, name) for name in os.listdir(EVIDENCE_DIR)
        if name.lower().endswith(".xml") and isfile(join(EVIDENCE_DIR, name))
    )

    theirs_xml = run_s1kd_brexcheck(S1KD_BREXCHECK_BIN, object_paths, brex_dir=EVIDENCE_DIR)
    theirs = parse_brex_check_xml_report(theirs_xml)

    checker = BrexChecker()
    checker.set_xml_dir(EVIDENCE_DIR)
    # `run_s1kd_brexcheck` passes `-S` and `-n`, so opt in on our side too --
    # both checks are off by default here as they are in the C original.
    results = checker.validate(check_sns=True, check_notations=True)
    ours = parse_brex_check_xml_report(checker.to_xml_report(results))

    diff = diff_violation_sets(ours, theirs)
    assert not diff["only_ours"], f"rules we flag that s1kd-brexcheck does not: {sorted(diff['only_ours'])[:20]}"
    assert not diff["only_theirs"], f"rules s1kd-brexcheck flags that we do not: {sorted(diff['only_theirs'])[:20]}"
