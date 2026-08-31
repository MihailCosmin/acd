import pytest

from acd.brex_checker import BrexChecker, XPATH_VERSIONS

# Mirrors s1kd-brexcheck's brex_requires_xpath2 (s1kd-brexcheck.c:1296-1322):
# issues 2.0-3.0 are XPath-1.0-only, 4.0+ (and an undeclared schema) default
# to XPath 2.0. The decision is based on the BREX's own declared schema, not
# the checked object's.
S1000D_3_0_SCHEMA = "http://www.s1000d.org/S1000D_3-0/xml_schema_flat/brex.xsd"
S1000D_4_2_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

# matches() is an XPath 2.0-only function (fn:matches); a rule that uses it
# only compiles under XPath2Parser, giving an observable pass/fail difference
# between the two versions rather than just inspecting parser.version.
MATCHES_RULE = """
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//*[matches(@code, "^[0-9]+$")]</objectPath>
        <objectUse>code must not look numeric (uses XPath 2.0-only matches()).</objectUse>
      </structureObjectRule>"""


def _brex(schema: str = None, extra_rule: str = "") -> str:
    root_attrs = ""
    if schema:
        root_attrs = (
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
            f' xsi:noNamespaceSchemaLocation="{schema}"'
        )
    return f"""<brex{root_attrs}>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>{extra_rule}
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


@pytest.fixture
def write_brex(tmp_path):
    def _write(name: str, content: str) -> str:
        path = tmp_path / name
        path.write_text(content, encoding="utf-8")
        return str(path)
    return _write


def _rule_for(rules, needle):
    return next(r for r in rules if needle in r['xpath'])


def test_dynamic_xpath1_for_s1000d_3_0_brex(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_3_0_SCHEMA))
    rules = BrexChecker()._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "1.0"


def test_dynamic_xpath2_for_s1000d_4_2_brex(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_4_2_SCHEMA))
    rules = BrexChecker()._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "2.0"


def test_dynamic_xpath2_when_brex_declares_no_schema(write_brex):
    brex_path = write_brex("brex.xml", _brex(None))
    rules = BrexChecker()._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "2.0"


@pytest.mark.parametrize("issue", ["2-0", "2-1", "2-2", "2-3", "3-0"])
def test_dynamic_xpath1_for_every_legacy_issue(write_brex, issue):
    schema = f"http://www.s1000d.org/S1000D_{issue}/xml_schema_flat/brex.xsd"
    brex_path = write_brex("brex.xml", _brex(schema))
    rules = BrexChecker()._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "1.0"


def test_matches_function_fails_to_compile_under_dynamic_xpath1_for_3_0_brex(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_3_0_SCHEMA, MATCHES_RULE))
    rules = BrexChecker()._show_rules(brex_path)
    rule = _rule_for(rules, "matches(")
    assert rule['selector'] is None
    assert "matches" in rule['selector_error']


def test_matches_function_compiles_under_dynamic_xpath2_for_4_2_brex(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_4_2_SCHEMA, MATCHES_RULE))
    rules = BrexChecker()._show_rules(brex_path)
    rule = _rule_for(rules, "matches(")
    assert rule['selector'] is not None
    assert rule['selector'].parser.version == "2.0"


def test_explicit_override_forces_xpath2_for_3_0_brex(write_brex):
    # Without the override, this BREX's own matches() rule cannot even compile
    # (previous test) -- the override makes it usable regardless.
    brex_path = write_brex("brex.xml", _brex(S1000D_3_0_SCHEMA, MATCHES_RULE))
    checker = BrexChecker()
    checker.set_xpath_version("2.0")
    rule = _rule_for(checker._show_rules(brex_path), "matches(")
    assert rule['selector'] is not None
    assert rule['selector'].parser.version == "2.0"


def test_explicit_override_forces_xpath1_for_4_2_brex(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_4_2_SCHEMA))
    checker = BrexChecker()
    checker.set_xpath_version("1.0")
    rules = checker._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "1.0"


def test_set_xpath_version_none_restores_dynamic_selection(write_brex):
    brex_path = write_brex("brex.xml", _brex(S1000D_4_2_SCHEMA))
    checker = BrexChecker()
    checker.set_xpath_version("1.0")
    checker.set_xpath_version(None)
    rules = checker._show_rules(brex_path)
    assert rules[0]['selector'].parser.version == "2.0"


def test_set_xpath_version_rejects_invalid_value():
    with pytest.raises(ValueError):
        BrexChecker().set_xpath_version("3.0")


def test_xpath_versions_constant_matches_supported_values():
    assert XPATH_VERSIONS == ("1.0", "2.0")


def test_lint_brex_invalid_xpath_reflects_actual_checking_behaviour(write_brex):
    # lint_brex's InvalidXPath check must compile with the same XPath version
    # _show_rules would really use for this BREX, not always XPath 2.0 --
    # otherwise it could pass a rule as "compiles fine" that would actually
    # raise an xpathError during real checking (or vice versa).
    brex_30 = write_brex("brex30.xml", _brex(S1000D_3_0_SCHEMA, MATCHES_RULE))
    brex_42 = write_brex("brex42.xml", _brex(S1000D_4_2_SCHEMA, MATCHES_RULE))

    checker = BrexChecker()
    findings_30 = checker.lint_brex(brex_30)
    findings_42 = checker.lint_brex(brex_42)

    assert any(f['Category'] == 'InvalidXPath' for f in findings_30)
    assert not any(f['Category'] == 'InvalidXPath' for f in findings_42)


def test_end_to_end_xpath_error_recorded_for_3_0_brex_using_matches(write_brex, tmp_path):
    brex_path = write_brex("brex.xml", _brex(S1000D_3_0_SCHEMA, MATCHES_RULE))
    xml_path = tmp_path / "object.xml"
    xml_path.write_text("<dmodule><forbiddenElement/></dmodule>", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    result = checker._check_rules()[brex_path]

    assert any("matches" in e['Error'] for e in result['xpathError'])
    # The other, XPath-1-compatible rule in the same BREX is unaffected.
    assert len(result['0']) == 1
