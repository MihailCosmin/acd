import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <brDecisionRef brDecisionIdentNumber="BR-GEN-00001"/>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="0">//alsoForbidden</objectPath>
        <objectUse>alsoForbidden must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-3">
        <brDecisionRef brDecisionIdentNumber="BR-GEN-00002"/>
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-4">
        <brDecisionRef brDecisionIdentNumber="BR-GEN-00003"/>
        <objectPath allowedObjectFlag="2">//@constrainedAttr</objectPath>
        <objectUse>constrainedAttr can only have codes "aa" or "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-5">
        <brDecisionRef brDecisionIdentNumber="BR-GEN-00004"/>
        <objectPath allowedObjectFlag="0">][invalid xpath(</objectPath>
        <objectUse>Deliberately malformed rule to exercise xpathError reporting.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <forbiddenElement/>\n"
    "  <alsoForbidden/>\n"
    '  <someElement constrainedAttr="zz"/>\n'
    "</dml>\n"
)


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def result(tmp_path, brex_path):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker._check_rules()[brex_path]


def _entry_for(entries, xpath):
    matches = [e for e in entries if e['Xpath'] == xpath]
    assert len(matches) == 1, f"expected exactly one entry for {xpath!r}, found {len(matches)}"
    return matches[0]


def test_flag_0_violation_carries_br_decision_ident_number(result):
    entry = _entry_for(result['0'], '//forbiddenElement')
    assert entry['BrDecisionIdentNumber'] == "BR-GEN-00001"


def test_flag_0_violation_without_br_decision_ref_reports_none(result):
    entry = _entry_for(result['0'], '//alsoForbidden')
    assert entry['BrDecisionIdentNumber'] is None


def test_flag_1_violation_carries_br_decision_ident_number(result):
    assert len(result['1']) == 1
    assert result['1'][0]['BrDecisionIdentNumber'] == "BR-GEN-00002"


def test_flag_2_violation_carries_br_decision_ident_number(result):
    assert len(result['2']) == 1
    assert result['2'][0]['BrDecisionIdentNumber'] == "BR-GEN-00003"


def test_xpath_error_carries_br_decision_ident_number(result):
    assert len(result['xpathError']) == 1
    assert result['xpathError'][0]['BrDecisionIdentNumber'] == "BR-GEN-00004"
