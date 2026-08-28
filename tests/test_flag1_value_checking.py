import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="1">//requiredAttr/@code</objectPath>
        <objectUse>requiredAttr must be present and its code must be "aa" or "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="1">//missingAttr/@code</objectPath>
        <objectUse>missingAttr must be present and its code must be "aa" or "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-3">
        <objectPath allowedObjectFlag="1">//valuelessAttr/@code</objectPath>
        <objectUse>valuelessAttr must be present, no value constraint.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <requiredAttr code="zz"/>\n'
    '  <valuelessAttr code="anything"/>\n'
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


def test_flag_1_present_but_invalid_value_reports_no_flag1_violation(result):
    assert all(e['Xpath'] != '//requiredAttr/@code' for e in result['1'])


def test_flag_1_present_but_invalid_value_reports_value_violation(result):
    assert len(result['2']) == 1
    entry = result['2'][0]
    assert entry['Xpath'] == '//requiredAttr/@code'
    assert 'zz' in entry['Description']


def test_flag_1_absent_still_reports_only_presence_violation(result):
    assert len(result['1']) == 1
    assert result['1'][0]['Xpath'] == '//missingAttr/@code'
    assert result['2'] == [] or all(e['Xpath'] != '//missingAttr/@code' for e in result['2'])


def test_flag_1_present_without_object_value_children_is_valid(result):
    assert all(e['Xpath'] != '//valuelessAttr/@code' for e in result['1'])
    assert all(e['Xpath'] != '//valuelessAttr/@code' for e in result['2'])
