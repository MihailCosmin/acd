import pytest

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
        <objectPath allowedObjectFlag="2">//constrainedElement</objectPath>
        <objectUse>constrainedElement can only contain "aa" or "bb".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="single" valueAllowed="bb"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

# Line numbers below are 1-based and deliberately padded with blank/comment
# lines so a text-scan-based line number (which would find the *first*
# textual occurrence of the attribute/element name) would disagree with the
# real, per-occurrence parsed-tree line number asserted here.
XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '                    # line 1
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <keep/>\n'                                                                     # line 2
    '  <!-- constrainedAttr placeholder, not a real match -->\n'                      # line 3
    '  <someElement constrainedAttr="zz"/>\n'                                         # line 4
    '\n'                                                                              # line 5
    '\n'                                                                              # line 6
    '  <someElement constrainedAttr="zz"/>\n'                                         # line 7
    '  <forbiddenElement id="f1"/>\n'                                                 # line 8
    '\n'                                                                              # line 9
    '  <forbiddenElement id="f2"/>\n'                                                 # line 10
    '  <constrainedElement>zz</constrainedElement>\n'                                 # line 11
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


def _entries_for(entries, xpath):
    return sorted((e for e in entries if e['Xpath'] == xpath), key=lambda e: e['Line'])


def test_flag_0_element_violations_report_real_per_occurrence_line(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entries = _entries_for(result['0'], '//forbiddenElement')

    assert [e['Line'] for e in entries] == [8, 10]


def test_flag_2_attribute_violations_report_owning_elements_real_line(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entries = _entries_for(result['2'], '//@constrainedAttr')

    assert [e['Line'] for e in entries] == [4, 7]


def test_flag_2_element_text_violation_reports_real_line(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entries = _entries_for(result['2'], '//constrainedElement')

    assert [e['Line'] for e in entries] == [11]


def test_line_number_is_int_not_string(tmp_path, brex_path):
    result = _make_checker(tmp_path, brex_path)._check_rules()[brex_path]
    entry = _entries_for(result['0'], '//forbiddenElement')[0]

    assert isinstance(entry['Line'], int)


# `and`-in-XPath rules used to be exempted from real line numbers by a pair of
# heuristics that substituted a "(Origin traced back to multiple lines -> ...)"
# placeholder for `Line`. Violations are one-per-node, so the violating node --
# and therefore its real line -- is always known.
AND_BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="AND-1">
        <objectPath allowedObjectFlag="0">//bad[@x and @y]</objectPath>
        <objectUse>bad with both x and y must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="AND-2">
        <objectPath allowedObjectFlag="2">//val[@a and @b]</objectPath>
        <objectUse>val with both a and b can only contain "ok".</objectUse>
        <objectValue valueForm="single" valueAllowed="ok"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

AND_XML_CONTENT = (
    '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '                # line 1
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <bad x="1" y="2"/>\n'                                                         # line 2
    '\n'                                                                             # line 3
    '  <bad x="1" y="2"/>\n'                                                         # line 4
    '  <val a="1" b="2">zz</val>\n'                                                  # line 5
    '\n'                                                                             # line 6
    '  <val a="1" b="2">zz</val>\n'                                                  # line 7
    "</dmodule>\n"
)


@pytest.fixture
def and_brex_path(tmp_path):
    path = tmp_path / "and_brex.xml"
    path.write_text(AND_BREX_CONTENT, encoding="utf-8")
    return str(path)


def _make_and_checker(tmp_path, and_brex_path):
    xml_path = tmp_path / "and_object.xml"
    xml_path.write_text(AND_XML_CONTENT, encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([and_brex_path])
    return checker


def test_flag_0_and_in_xpath_reports_real_lines(tmp_path, and_brex_path):
    result = _make_and_checker(tmp_path, and_brex_path)._check_rules()[and_brex_path]
    entries = _entries_for(result['0'], '//bad[@x and @y]')

    assert [e['Line'] for e in entries] == [2, 4]
    assert [e['NodeXpath'] for e in entries] == ['/dmodule[1]/bad[1]', '/dmodule[1]/bad[2]']


def test_flag_2_and_in_xpath_reports_real_lines(tmp_path, and_brex_path):
    result = _make_and_checker(tmp_path, and_brex_path)._check_rules()[and_brex_path]
    entries = _entries_for(result['2'], '//val[@a and @b]')

    assert [e['Line'] for e in entries] == [5, 7]
