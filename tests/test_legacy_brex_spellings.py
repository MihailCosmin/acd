import pytest

from acd.brex_checker import BrexChecker
from acd.default_brex import default_brex_path

# S1000D <= 3.0 spelling: contextrules/structrules/objrule/objpath|objuse|objval,
# @objappl (not @allowedObjectFlag), @valtype (not @valueForm), @val1[/@val2]
# (not @valueAllowed). Mirrors the real bundled DMC-AE-A-...-022A-D BREX
# (acd/brex/DMC-AE-A-04-10-0301-00A-022A-D_003-00.XML), which uses exactly
# this structure throughout.
BREX_CONTENT = """<brex>
  <contextrules>
    <structrules>
      <objrule>
        <objpath objappl="0">//forbiddenElement</objpath>
        <objuse>forbiddenElement must not be present.</objuse>
      </objrule>
      <objrule>
        <objpath objappl="1">//requiredElement</objpath>
        <objuse>requiredElement must be present.</objuse>
      </objrule>
      <objrule>
        <objpath objappl="2">//@constrainedAttr</objpath>
        <objuse>constrainedAttr can only have codes "aa" or "bb".</objuse>
        <objval valtype="single" val1="aa"/>
        <objval valtype="single" val1="bb"/>
      </objrule>
      <objrule>
        <objpath>//@accpnltype</objpath>
        <objuse>Type of access panel (no @objappl at all, like most rules in the real S1000D 3.0 default BREX).</objuse>
        <objval valtype="single" val1="accpnl01"/>
        <objval valtype="single" val1="accpnl02"/>
      </objrule>
      <objrule>
        <objpath>//@rangedAttr</objpath>
        <objuse>rangedAttr must be between 20 and 100.</objuse>
        <objval valtype="range" val1="20" val2="100"/>
      </objrule>
      <objrule>
        <objpath>//@patternedAttr</objpath>
        <objuse>patternedAttr must be exactly two digits.</objuse>
        <objval valtype="pattern" val1="[0-9]{2}"/>
      </objrule>
      <objrule>
        <objpath>//@undefinedAttr</objpath>
        <objuse>Informational only: no objval children, so this is never a violation.</objuse>
      </objrule>
    </structrules>
  </contextrules>
  <contextrules context="http://example.com/qualified-schema.xsd">
    <structrules>
      <objrule>
        <objpath objappl="0">//onlyForQualifiedSchema</objpath>
        <objuse>onlyForQualifiedSchema must not be present, but only under the qualified schema.</objuse>
      </objrule>
    </structrules>
  </contextrules>
</brex>
"""

# No xsi:noNamespaceSchemaLocation, so get_schema_from_xml resolves to None,
# matching a real DTD-based S1000D <= 3.0 content object. onlyForQualifiedSchema
# is present but its rule only applies under "http://example.com/qualified-schema.xsd",
# which this object does not declare, so that rule must not fire.
XML_CONTENT = """<dml>
  <forbiddenElement/>
  <someElement constrainedAttr="zz" accpnltype="badcode" rangedAttr="150" patternedAttr="X1" undefinedAttr="whatever"/>
  <onlyForQualifiedSchema/>
</dml>
"""

VALID_XML_CONTENT = """<dml>
  <requiredElement/>
  <someElement constrainedAttr="aa" accpnltype="accpnl02" rangedAttr="50" patternedAttr="42" undefinedAttr="whatever"/>
</dml>
"""


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def _check(tmp_path, brex_path, xml_content):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker._check_rules()[brex_path]


@pytest.fixture
def result(tmp_path, brex_path):
    return _check(tmp_path, brex_path, XML_CONTENT)


@pytest.fixture
def valid_result(tmp_path, brex_path):
    return _check(tmp_path, brex_path, VALID_XML_CONTENT)


def _entry_for(entries, xpath):
    matches = [e for e in entries if e['Xpath'] == xpath]
    assert len(matches) == 1, f"expected exactly one entry for {xpath!r}, found {len(matches)}"
    return matches[0]


def test_legacy_flag_0_violation(result):
    entry = _entry_for(result['0'], '//forbiddenElement')
    assert entry['Description'] == "forbiddenElement must not be present."


def test_legacy_flag_1_violation(result):
    entry = _entry_for(result['1'], '//requiredElement')
    assert entry['Description'] == "requiredElement must be present."


def test_legacy_flag_2_single_value_violation(result):
    entry = _entry_for(result['2'], '//@constrainedAttr')
    assert entry['Single Values'] == [['aa', 'bb']]


def test_legacy_rule_without_objappl_still_value_checked(result):
    # The majority shape in the real S1000D <= 3.0 default BREX: no @objappl
    # at all, just objval children constraining the matched value.
    entry = _entry_for(result['2'], '//@accpnltype')
    assert entry['Single Values'] == [['accpnl01', 'accpnl02']]


def test_legacy_range_valtype(result):
    entry = _entry_for(result['2'], '//@rangedAttr')
    assert entry['Range Values'] == [['20~100']]


def test_legacy_pattern_valtype(result):
    entry = _entry_for(result['2'], '//@patternedAttr')
    assert len(entry['Pattern Values'][0]) == 1


def test_legacy_rule_without_objval_is_never_a_violation(result):
    assert not any(e['Xpath'] == '//@undefinedAttr' for e in result['2'])


def test_legacy_context_qualified_rule_skipped_for_non_matching_schema(result):
    # onlyForQualifiedSchema IS present in XML_CONTENT, but its objrule sits
    # in a contextrules qualified to a schema the checked object doesn't
    # declare, so it must not be reported even though the plain-flag-0 rule
    # for forbiddenElement (unqualified contextrules) does fire.
    assert not any(e['Xpath'] == '//onlyForQualifiedSchema' for e in result['0'])
    assert len(result['0']) == 1


def test_legacy_rules_pass_with_valid_values(valid_result):
    assert valid_result['0'] == []
    assert valid_result['1'] == []
    assert valid_result['2'] == []


def test_real_bundled_legacy_default_brex_end_to_end(tmp_path):
    # acd/brex/DMC-AE-A-04-10-0301-00A-022A-D_003-00.XML is the actual
    # S1000D <= 3.0 default BREX shipped with s1kd-brexcheck, unmodified: a
    # real-world file exercising this spelling, not just a hand-built fixture.
    real_brex_path = default_brex_path("DMC-AE-A-04-10-0301-00A-022A-D")

    bad_xml = tmp_path / "bad.xml"
    bad_xml.write_text('<dml><foo accpnltype="notallowed"/></dml>', encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(bad_xml))
    checker.override_brex_list([real_brex_path])
    bad_result = checker._check_rules()[real_brex_path]
    assert any(e['Xpath'] == '//@accpnltype' for e in bad_result['2'])

    good_xml = tmp_path / "good.xml"
    good_xml.write_text('<dml><foo accpnltype="accpnl02"/></dml>', encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(good_xml))
    checker.override_brex_list([real_brex_path])
    good_result = checker._check_rules()[real_brex_path]
    assert not any(e['Xpath'] == '//@accpnltype' for e in good_result['2'])
