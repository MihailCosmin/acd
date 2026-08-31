import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="2">//@constrainedAttr</objectPath>
        <objectUse>constrainedAttr can only have codes "aa" (lexical) or "bb" (restrictable).</objectUse>
        <objectValue valueForm="single" valueAllowed="aa" valueTailoring="lexical"/>
        <objectValue valueForm="single" valueAllowed="bb" valueTailoring="restrictable"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="2">//@untaggedAttr</objectPath>
        <objectUse>untaggedAttr can only have code "cc", no tailoring declared.</objectUse>
        <objectValue valueForm="single" valueAllowed="cc"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-3">
        <objectPath allowedObjectFlag="2">//@patternAttr</objectPath>
        <objectUse>patternAttr must match [0-9]{2} (restrictable).</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[0-9]{2}" valueTailoring="restrictable"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-4">
        <objectPath allowedObjectFlag="2">//@rangeAttr</objectPath>
        <objectUse>rangeAttr must be within accpnl51~accpnl99 (lexical).</objectUse>
        <objectValue valueForm="range" valueAllowed="accpnl51~accpnl99" valueTailoring="lexical"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-5">
        <objectPath allowedObjectFlag="1">//requiredAttr/@code</objectPath>
        <objectUse>requiredAttr must be present and its code must be "aa" (lexical).</objectUse>
        <objectValue valueForm="single" valueAllowed="aa" valueTailoring="lexical"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <someElement constrainedAttr="zz" untaggedAttr="zz" patternAttr="XX" rangeAttr="accpnl05"/>\n'
    '  <requiredAttr code="zz"/>\n'
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


def _violation_for(result, xpath):
    matches = [e for e in result['2'] if e['Xpath'] == xpath]
    assert len(matches) == 1, f"expected exactly one '2' violation for {xpath!r}, found {len(matches)}"
    return matches[0]


def test_single_form_violation_distinguishes_lexical_and_restrictable(result):
    entry = _violation_for(result, '//@constrainedAttr')
    tailoring_by_value = {t['valueAllowed']: t['valueTailoring'] for t in entry['ValueTailoring']}
    assert tailoring_by_value == {"aa": "lexical", "bb": "restrictable"}
    assert all(t['valueForm'] == 'single' for t in entry['ValueTailoring'])


def test_violation_without_declared_tailoring_reports_empty_list(result):
    entry = _violation_for(result, '//@untaggedAttr')
    assert entry['ValueTailoring'] == []


def test_pattern_form_violation_carries_tailoring(result):
    entry = _violation_for(result, '//@patternAttr')
    assert entry['ValueTailoring'] == [{
        'valueForm': 'pattern',
        'valueAllowed': '[0-9]{2}',
        'valueTailoring': 'restrictable',
    }]


def test_range_form_violation_carries_tailoring(result):
    entry = _violation_for(result, '//@rangeAttr')
    assert entry['ValueTailoring'] == [{
        'valueForm': 'range',
        'valueAllowed': 'accpnl51~accpnl99',
        'valueTailoring': 'lexical',
    }]


def test_flag_1_value_check_path_also_carries_tailoring(result):
    # A flag-1 rule whose node is present but whose value is invalid is
    # reported through the shared `_check_object_values` path (see
    # test_flag1_value_checking.py) and lands in bucket '2', same as a
    # flag-2 rule; it must carry ValueTailoring too. Ref §3.8.
    entry = _violation_for(result, '//requiredAttr/@code')
    assert entry['ValueTailoring'] == [{
        'valueForm': 'single',
        'valueAllowed': 'aa',
        'valueTailoring': 'lexical',
    }]


def test_legacy_spelling_has_no_tailoring_attribute_and_reports_empty_list(tmp_path):
    # @valueTailoring was introduced in a later S1000D issue than the
    # objpath/objval legacy spelling (CPF 2009-039S1); a legacy-spelled rule
    # simply has nothing to report.
    legacy_brex_content = """<brex>
  <contextrules>
    <structrules>
      <objrule>
        <objpath objappl="2">//@legacyAttr</objpath>
        <objuse>legacyAttr can only be "aa".</objuse>
        <objval valtype="single" val1="aa"/>
      </objrule>
    </structrules>
  </contextrules>
</brex>
"""
    brex_path = tmp_path / "legacy_brex.xml"
    brex_path.write_text(legacy_brex_content, encoding="utf-8")

    xml_content = "<dml>\n  <someElement legacyAttr=\"zz\"/>\n</dml>\n"
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(brex_path)])
    legacy_result = checker._check_rules()[str(brex_path)]

    entry = next(e for e in legacy_result['2'] if e['Xpath'] == '//@legacyAttr')
    assert entry['ValueTailoring'] == []
