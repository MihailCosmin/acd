import pytest

from lxml import etree

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="2">//@constrainedAttr</objectPath>
        <objectUse>constrainedAttr can only have codes "aa" or "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-3">
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-4">
        <objectPath allowedObjectFlag="0">count(//tooMany) &gt; 1</objectPath>
        <objectUse>at most one tooMany element is allowed.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <keep/>\n'
    '  <forbiddenElement id="f1"><child/>text</forbiddenElement>\n'
    '  <forbiddenElement id="f2"/>\n'
    '  <someElement constrainedAttr="zz"/>\n'
    '  <tooMany/>\n'
    '  <tooMany/>\n'
    "</dml>\n"
)


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def _make_checker(tmp_path, brex_path):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def _entry_for(entries, xpath):
    matches = [e for e in entries if e['Xpath'] == xpath]
    assert len(matches) >= 1, f"expected at least one entry for {xpath!r}, found {len(matches)}"
    return matches


def test_flag_0_element_violation_reports_canonical_xpath_and_shallow_copy(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entries = sorted(_entry_for(result['0'], '//forbiddenElement'), key=lambda e: e['NodeXpath'])

    assert entries[0]['NodeXpath'] == '/dml[1]/forbiddenElement[1]'
    assert entries[1]['NodeXpath'] == '/dml[1]/forbiddenElement[2]'

    # Shallow copy (default): own tag/attributes only, no children or text.
    copied = etree.fromstring(entries[0]['Object'].encode('utf-8'))
    assert copied.tag == 'forbiddenElement'
    assert copied.get('id') == 'f1'
    assert len(copied) == 0
    assert not (copied.text or '').strip()


def test_flag_0_element_violation_deep_copy_includes_subtree(tmp_path, brex_path):
    checker = _make_checker(tmp_path, brex_path)
    result = checker._check_rules(deep_copy_nodes=True)[brex_path]
    entry = [e for e in result['0'] if e['NodeXpath'] == '/dml[1]/forbiddenElement[1]'][0]

    copied = etree.fromstring(entry['Object'].encode('utf-8'))
    assert copied.tag == 'forbiddenElement'
    assert copied.get('id') == 'f1'
    assert len(copied) == 1
    assert copied[0].tag == 'child'
    assert copied[0].tail == 'text'


def test_flag_2_attribute_violation_reports_owning_element(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entry = _entry_for(result['2'], '//@constrainedAttr')[0]

    assert entry['NodeXpath'] == '/dml[1]/someElement[1]/@constrainedAttr'
    # The attribute itself has no subtree: the copy is of its owning element.
    copied = etree.fromstring(entry['Object'].encode('utf-8'))
    assert copied.tag == 'someElement'
    assert copied.get('constrainedAttr') == 'zz'


def test_flag_1_missing_element_has_no_backing_node(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entry = _entry_for(result['1'], '//requiredElement')[0]

    assert entry['NodeXpath'] is None
    assert entry['Object'] is None


def test_flag_0_boolean_rule_has_no_backing_node(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entry = _entry_for(result['0'], 'count(//tooMany) > 1')[0]

    assert entry['NodeXpath'] is None
    assert entry['Object'] is None


def test_validate_threads_deep_copy_nodes_end_to_end(tmp_path, brex_path):
    checker = _make_checker(tmp_path, brex_path)
    result = checker.validate(deep_copy_nodes=True)
    entry = [e for e in result[brex_path]['0'] if e['NodeXpath'] == '/dml[1]/forbiddenElement[1]'][0]

    copied = etree.fromstring(entry['Object'].encode('utf-8'))
    assert len(copied) == 1
