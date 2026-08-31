"""End-to-end malformed-`objectPath` fixture for brex_checker_rework.md
§4.5 / §3.12: a BREX containing one unparseable rule alongside otherwise
normal, violated rules must still let the run complete, reporting an
`xpathError` for the broken rule only -- every other rule is checked and
reported normally.

test_xpath_error_handling.py already covers this at the
`_check_object_flag_0/1/2` unit level, with an empty violations dict and no
other rules in play. This exercises the same defect end-to-end through
`BrexChecker.validate()`, alongside real, correctly-firing flag 0/1/2 rules
that must be unaffected by the broken one.
"""

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BAD_XPATH = "][invalid xpath("

BREX_CONTENT = f"""<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-MALFORMED">
        <objectPath allowedObjectFlag="0">{BAD_XPATH}</objectPath>
        <objectUse>This rule's objectPath cannot be compiled.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-GOOD-0">
        <objectPath allowedObjectFlag="0">//forbidden</objectPath>
        <objectUse>forbidden must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-GOOD-1">
        <objectPath allowedObjectFlag="1">//required</objectPath>
        <objectUse>required must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-GOOD-2">
        <objectPath allowedObjectFlag="2">//constrained/@code</objectPath>
        <objectUse>constrained/@code must be "aa".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

# Violates SOR-GOOD-0 (forbidden present) and SOR-GOOD-2 (wrong value);
# omits //required so SOR-GOOD-1 also fires. 3 real violations total.
XML_CONTENT = (
    '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <forbidden/>\n"
    '  <constrained code="zz"/>\n'
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

    return checker.validate(), str(brex_path)


def test_run_completes_without_raising(result):
    # The mere fact this fixture ran to completion (via the `result`
    # fixture) already proves no exception propagated out of validate().
    validated_result, brex_path = result
    assert brex_path in validated_result


def test_exactly_one_xpath_error_is_reported_for_the_malformed_rule_only(result):
    validated_result, brex_path = result
    errors = validated_result[brex_path]["xpathError"]
    assert len(errors) == 1
    assert errors[0]["Xpath"] == BAD_XPATH
    assert errors[0]["Description"] == "This rule's objectPath cannot be compiled."


def test_the_other_three_rules_are_still_checked_and_report_correctly(result):
    validated_result, brex_path = result
    assert [v["Xpath"] for v in validated_result[brex_path]["0"]] == ["//forbidden"]
    assert [v["Xpath"] for v in validated_result[brex_path]["1"]] == ["//required"]
    assert [v["Xpath"] for v in validated_result[brex_path]["2"]] == ["//constrained/@code"]


def test_summary_counts_only_the_real_violations_not_the_xpath_error(result):
    validated_result, _brex_path = result
    assert validated_result["Summary"] == "3 Errors"
