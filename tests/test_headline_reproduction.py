"""Regression test for brex_checker_rework.md §1.1 / §3.1.

Before the P0 fixes, a minimal BREX with 10 deliberately-violated rules (5 x
flag 0, 2 x flag 1, 3 x flag 2) reported "3 Errors" instead of 10: the
error-key collision in `_check_rules` (§3.1) meant only the last violation of
each flag survived. This pins the exact scenario from the plan's evidence
base so that regression cannot silently come back.
"""

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-F0-1">
        <objectPath allowedObjectFlag="0">//forbidden1</objectPath>
        <objectUse>forbidden1 must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F0-2">
        <objectPath allowedObjectFlag="0">//forbidden2</objectPath>
        <objectUse>forbidden2 must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F0-3">
        <objectPath allowedObjectFlag="0">//forbidden3</objectPath>
        <objectUse>forbidden3 must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F0-4">
        <objectPath allowedObjectFlag="0">//forbidden4</objectPath>
        <objectUse>forbidden4 must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F0-5">
        <objectPath allowedObjectFlag="0">//forbidden5</objectPath>
        <objectUse>forbidden5 must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F1-1">
        <objectPath allowedObjectFlag="1">//required1</objectPath>
        <objectUse>required1 must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F1-2">
        <objectPath allowedObjectFlag="1">//required2</objectPath>
        <objectUse>required2 must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-F2-1">
        <objectPath allowedObjectFlag="2">//constrained1/@code</objectPath>
        <objectUse>constrained1/@code must be "aa".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-F2-2">
        <objectPath allowedObjectFlag="2">//constrained2/@code</objectPath>
        <objectUse>constrained2/@code must be "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-F2-3">
        <objectPath allowedObjectFlag="2">//constrained3/@code</objectPath>
        <objectUse>constrained3/@code must be "cc".</objectUse>
        <objectValue valueForm="single" valueAllowed="cc"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

# Violates every rule above: forbidden1..5 are present (5 flag-0 violations);
# required1/required2 are absent (2 flag-1 violations); constrained1..3 are
# present but carry the wrong value (3 flag-2 violations).
XML_CONTENT = (
    '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <forbidden1/>\n"
    "  <forbidden2/>\n"
    "  <forbidden3/>\n"
    "  <forbidden4/>\n"
    "  <forbidden5/>\n"
    '  <constrained1 code="zz"/>\n'
    '  <constrained2 code="zz"/>\n'
    '  <constrained3 code="zz"/>\n'
    "</dmodule>\n"
)


@pytest.fixture
def result(tmp_path):
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(brex_path)])

    return checker.validate()


def test_exactly_ten_violations_are_reported(result):
    # The historical bug reported "3 Errors" (1 per flag, last rule only).
    assert result["Summary"] == "10 Errors"


def test_all_five_flag_0_violations_survive_with_distinct_rules(result):
    brex_path = next(k for k in result if k not in ("Summary", "brexFallback", "sns", "notations", "nonContextRules"))
    violations = result[brex_path]["0"]
    assert len(violations) == 5
    descriptions_by_xpath = {v["Xpath"]: v["Description"] for v in violations}
    assert descriptions_by_xpath == {
        "//forbidden1": "forbidden1 must not be present.",
        "//forbidden2": "forbidden2 must not be present.",
        "//forbidden3": "forbidden3 must not be present.",
        "//forbidden4": "forbidden4 must not be present.",
        "//forbidden5": "forbidden5 must not be present.",
    }


def test_both_flag_1_violations_survive_with_distinct_rules(result):
    brex_path = next(k for k in result if k not in ("Summary", "brexFallback", "sns", "notations", "nonContextRules"))
    violations = result[brex_path]["1"]
    assert len(violations) == 2
    assert {v["Xpath"] for v in violations} == {"//required1", "//required2"}


def test_all_three_flag_2_violations_survive_with_distinct_rules_and_values(result):
    brex_path = next(k for k in result if k not in ("Summary", "brexFallback", "sns", "notations", "nonContextRules"))
    violations = result[brex_path]["2"]
    assert len(violations) == 3
    allowed_by_xpath = {v["Xpath"]: v["Single Values"] for v in violations}
    assert allowed_by_xpath == {
        "//constrained1/@code": [["aa"]],
        "//constrained2/@code": [["bb"]],
        "//constrained3/@code": [["cc"]],
    }
