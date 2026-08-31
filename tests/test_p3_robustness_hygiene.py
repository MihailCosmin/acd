import re
from unittest.mock import patch

import elementpath
import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbidden <emphasis>element</emphasis> must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    '  <forbiddenElement/>\n'
    '</dml>\n'
)


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def _make_checker(tmp_path, brex_file, xml_name="object.xml", xml_content=XML_CONTENT):
    xml_path = tmp_path / xml_name
    xml_path.write_text(xml_content, encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_file])
    return checker


# --- objectUse read as full text content, tolerating a missing one ---------

def test_object_use_reads_full_text_content_not_just_leading_text(tmp_path, brex_path):
    checker = _make_checker(tmp_path, brex_path)
    result = checker._check_rules()[brex_path]
    entry = next(e for e in result['0'] if e['Xpath'] == '//forbiddenElement')
    assert entry['Description'] == "forbidden element must not be present."


def test_missing_object_use_does_not_raise_index_error(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_file = tmp_path / "brex_no_use.xml"
    brex_file.write_text(brex_content, encoding="utf-8")
    checker = _make_checker(tmp_path, str(brex_file), xml_name="object_no_use.xml")
    result = checker._check_rules()[str(brex_file)]
    entry = next(e for e in result['0'] if e['Xpath'] == '//forbiddenElement')
    assert entry['Description'] is None


# --- BREX parsed once per run, rule list cached -----------------------------

def test_brex_parsed_once_per_run_not_per_document(tmp_path, brex_path):
    checker = BrexChecker()
    checker.override_brex_list([brex_path])

    with patch.object(checker, '_get_object_rule_nodes',
                       wraps=checker._get_object_rule_nodes) as spy:
        for name in ("object1.xml", "object2.xml", "object3.xml"):
            xml_path = tmp_path / name
            xml_path.write_text(XML_CONTENT, encoding="utf-8")
            checker.set_xml(str(xml_path))
            checker._check_rules()

    assert spy.call_count == 1


def test_get_content_rules_returns_cached_list_across_calls(tmp_path, brex_path):
    checker = _make_checker(tmp_path, brex_path)
    rules_1 = checker._get_content_rules(brex_path, schema=DMODULE_SCHEMA)
    rules_2 = checker._get_content_rules(brex_path, schema=DMODULE_SCHEMA)
    assert rules_1 is rules_2


# --- objectPath selector compiled once and reused ---------------------------

def test_rule_carries_a_compiled_reusable_selector(tmp_path, brex_path):
    checker = _make_checker(tmp_path, brex_path)
    rules = checker._get_content_rules(brex_path, schema=DMODULE_SCHEMA)
    rule = next(r for r in rules if r['xpath'] == '//forbiddenElement')
    assert isinstance(rule['selector'], elementpath.Selector)

    rules_again = checker._get_content_rules(brex_path, schema=DMODULE_SCHEMA)
    rule_again = next(r for r in rules_again if r['xpath'] == '//forbiddenElement')
    assert rule_again['selector'] is rule['selector']


def test_invalid_object_path_reports_xpath_error_without_raising(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">][invalid xpath(</objectPath>
        <objectUse>bad rule</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_file = tmp_path / "bad_brex.xml"
    brex_file.write_text(brex_content, encoding="utf-8")
    checker = _make_checker(tmp_path, str(brex_file), xml_name="object_bad.xml")
    result = checker._check_rules()[str(brex_file)]
    assert len(result['xpathError']) == 1
    assert result['xpathError'][0]['Xpath'] == '][invalid xpath('
    assert result['0'] == []


# --- regex_builder escapes attribute name/value -----------------------------

def test_regex_builder_escapes_attribute_value_metacharacters():
    checker = BrexChecker()
    pattern = checker.regex_builder('attr', 'a.b(c)', 'xpath')
    compiled = re.compile(pattern)
    assert compiled.search('attr = "a.b(c)"')
    assert compiled.search('attr = "aXbYcZ"') is None


def test_regex_builder_escapes_attribute_name_metacharacters():
    checker = BrexChecker()
    pattern = checker.regex_builder('a.b', None, 'xpath')
    compiled = re.compile(pattern)
    assert compiled.search('a.b = "x"')
    assert compiled.search('aXb = "x"') is None


# --- Saxon path removed -------------------------------------------------

def test_saxon_parameter_and_attribute_are_gone():
    with pytest.raises(TypeError):
        BrexChecker(saxon=True)
    checker = BrexChecker()
    assert not hasattr(checker, '_saxon')
