import pytest

from acd.brex_checker import BrexChecker, NS_DICT

# Two different, non-rdf/xsi namespaces are declared at two different points in
# the BREX: one on the root element (in scope for every objectPath), one on a
# single structureObjectRule (in scope only for that rule's objectPath). This
# reproduces §3.12: prefixes must be resolved from what is actually in scope at
# each objectPath node, not from a single hard-coded rdf+xsi dictionary.
BREX_CONTENT = """<brex xmlns:cust="http://example.com/custom-ns">
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//cust:forbidden</objectPath>
        <objectUse>cust:forbidden must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule xmlns:other="http://example.com/other-ns">
        <objectPath allowedObjectFlag="1">//other:required</objectPath>
        <objectUse>other:required must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//@xsi:type</objectPath>
        <objectUse>xsi:type must not be present (rdf/xsi still resolve with no local declaration).</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_CONTENT = """<dmodule xmlns:c="http://example.com/custom-ns"
                  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <c:forbidden/>
  <someElement xsi:type="whatever"/>
</dmodule>
"""


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


def test_root_scoped_prefix_is_resolved(result):
    # "cust" is declared only on the BREX root, several ancestors above
    # objectPath; it must still be picked up via the node's full nsmap.
    entry = _entry_for(result['0'], '//cust:forbidden')
    assert entry['Description'].startswith('cust:forbidden')


def test_rule_local_prefix_is_resolved(result):
    # "other" is declared only on this one structureObjectRule, so it must
    # not leak into or be missing from any other rule's namespace map.
    entry = _entry_for(result['1'], '//other:required')
    assert entry['Description'].startswith('other:required')


def test_rdf_xsi_still_resolve_without_local_declaration(result):
    # rdf/xsi are not declared anywhere in this BREX; NS_DICT is kept as a
    # base so these well-known prefixes remain usable for backward compat.
    entry = _entry_for(result['0'], '//@xsi:type')
    assert entry['Description'].startswith('xsi:type')


def test_namespaces_captured_per_rule(brex_path):
    checker = BrexChecker()
    rules = checker._show_rules(brex_path)
    by_xpath = {rule['xpath']: rule for rule in rules}

    cust_rule = by_xpath['//cust:forbidden']
    other_rule = by_xpath['//other:required']

    assert cust_rule['namespaces']['cust'] == 'http://example.com/custom-ns'
    assert 'other' not in cust_rule['namespaces']

    assert other_rule['namespaces']['other'] == 'http://example.com/other-ns'
    assert other_rule['namespaces']['cust'] == 'http://example.com/custom-ns'

    # NS_DICT stays available as a base on every rule.
    for uri_by_prefix in (cust_rule['namespaces'], other_rule['namespaces']):
        for prefix, uri in NS_DICT.items():
            assert uri_by_prefix[prefix] == uri
