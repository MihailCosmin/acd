"""End-to-end `objectValue/@valueForm` fixture set for brex_checker_rework.md
§4.5: `single`, `pattern` (including XSD character-class subtraction and
whole-value anchoring, §3.4) and `range` in both `~` range and `|` set form
(§3.5), each checked through `BrexChecker.validate()` rather than the
low-level helpers alone (those already have dedicated unit coverage in
test_range_set_values.py / test_xsd_regex_translation.py).
"""

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-SINGLE">
        <objectPath allowedObjectFlag="2">//singleAttr/@code</objectPath>
        <objectUse>singleAttr/@code must be exactly "aa".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-PATTERN-SUBTRACTION">
        <objectPath allowedObjectFlag="2">//patternSubtractionAttr/@code</objectPath>
        <objectUse>patternSubtractionAttr/@code must match 00[A-Z-[IO]]{1,3}, excluding I and O.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="00[A-Z-[IO]]{1,3}"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-PATTERN-WHOLE-VALUE">
        <objectPath allowedObjectFlag="2">//patternWholeValueAttr/@code</objectPath>
        <objectUse>patternWholeValueAttr/@code must be exactly two digits, not merely contain two digits.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[0-9]{2}"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-RANGE">
        <objectPath allowedObjectFlag="2">//rangeAttr/@code</objectPath>
        <objectUse>rangeAttr/@code must be within accpnl51~accpnl99.</objectUse>
        <objectValue valueForm="range" valueAllowed="accpnl51~accpnl99"/>
      </structureObjectRule>
      <structureObjectRule id="SOR-SET">
        <objectPath allowedObjectFlag="2">//setAttr/@code</objectPath>
        <objectUse>setAttr/@code must be one of A, B or C.</objectUse>
        <objectValue valueForm="range" valueAllowed="A|B|C"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


def make_dm(**attrs) -> str:
    elements = "\n".join(f'  <{name} code="{value}"/>' for name, value in attrs.items())
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        f"{elements}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def _validate(tmp_path, brex_path, **attrs):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_dm(**attrs), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker._check_rules()[brex_path]


# ---------------------------------------------------------------------------
# single
# ---------------------------------------------------------------------------

def test_single_matching_value_produces_no_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, singleAttr="aa")
    assert result['2'] == []


def test_single_non_matching_value_is_a_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, singleAttr="zz")
    assert len(result['2']) == 1
    entry = result['2'][0]
    assert entry['Xpath'] == '//singleAttr/@code'
    assert entry['Single Values'] == [['aa']]


# ---------------------------------------------------------------------------
# pattern -- XSD character-class subtraction
# ---------------------------------------------------------------------------

def test_pattern_subtraction_accepts_a_letter_outside_the_excluded_class(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, patternSubtractionAttr="00ABC")
    assert all(e['Xpath'] != '//patternSubtractionAttr/@code' for e in result['2'])


def test_pattern_subtraction_rejects_an_excluded_letter(tmp_path, brex_path):
    # "I" and "O" are excluded by [A-Z-[IO]] even though they are in A-Z.
    result = _validate(tmp_path, brex_path, patternSubtractionAttr="00IOI")
    matches = [e for e in result['2'] if e['Xpath'] == '//patternSubtractionAttr/@code']
    assert len(matches) == 1
    assert matches[0]['Pattern Values'] == [['00[A-Z--[IO]]{1,3}']]


# ---------------------------------------------------------------------------
# pattern -- whole-value anchoring
# ---------------------------------------------------------------------------

def test_pattern_whole_value_accepts_an_exact_match(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, patternWholeValueAttr="12")
    assert all(e['Xpath'] != '//patternWholeValueAttr/@code' for e in result['2'])


def test_pattern_whole_value_rejects_a_partial_match(tmp_path, brex_path):
    # A substring match ("12" inside "XX12XX") must not be accepted -- the
    # whole attribute value has to match the pattern (§3.4).
    result = _validate(tmp_path, brex_path, patternWholeValueAttr="XX12XX")
    matches = [e for e in result['2'] if e['Xpath'] == '//patternWholeValueAttr/@code']
    assert len(matches) == 1


# ---------------------------------------------------------------------------
# range (a~c form)
# ---------------------------------------------------------------------------

def test_range_value_inside_bounds_produces_no_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, rangeAttr="accpnl67")
    assert all(e['Xpath'] != '//rangeAttr/@code' for e in result['2'])


def test_range_value_outside_bounds_is_a_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, rangeAttr="accpnl05")
    matches = [e for e in result['2'] if e['Xpath'] == '//rangeAttr/@code']
    assert len(matches) == 1
    assert matches[0]['Range Values'] == [['accpnl51~accpnl99']]


# ---------------------------------------------------------------------------
# set (a|b|c form)
# ---------------------------------------------------------------------------

def test_set_member_produces_no_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, setAttr="B")
    assert all(e['Xpath'] != '//setAttr/@code' for e in result['2'])


def test_set_non_member_is_a_violation(tmp_path, brex_path):
    result = _validate(tmp_path, brex_path, setAttr="D")
    matches = [e for e in result['2'] if e['Xpath'] == '//setAttr/@code']
    assert len(matches) == 1
    assert matches[0]['Range Values'] == [['A|B|C']]


# ---------------------------------------------------------------------------
# All forms together, exactly as a real BREX mixes them within one rule set.
# ---------------------------------------------------------------------------

def test_all_forms_checked_together_report_only_the_actual_violations(tmp_path, brex_path):
    result = _validate(
        tmp_path, brex_path,
        singleAttr="aa",                     # valid
        patternSubtractionAttr="00IOI",      # invalid
        patternWholeValueAttr="12",          # valid
        rangeAttr="accpnl05",                # invalid
        setAttr="B",                          # valid
    )

    violated_xpaths = {e['Xpath'] for e in result['2']}
    assert violated_xpaths == {'//patternSubtractionAttr/@code', '//rangeAttr/@code'}
