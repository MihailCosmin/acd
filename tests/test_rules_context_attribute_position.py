"""End-to-end regression for brex_checker_rework.md §3.3 / §4.5: the checked
object's `xsi:noNamespaceSchemaLocation` must be read correctly regardless of
its position among the root element's attributes, so a `rulesContext`-
qualified rule group is selected the same way every time.

`test_xml_processing.py::test_get_schema_from_xml_is_position_independent`
already covers this at the `get_schema_from_xml` unit level; this asserts
the same thing end-to-end through `BrexChecker.validate()` -- that the
*same rule set* actually gets selected and checked, not just that the right
string is extracted.
"""

import pytest

from acd.brex_checker import BrexChecker

SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = f"""<brex>
  <contextRules rulesContext="{SCHEMA}">
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-QUALIFIED">
        <objectPath allowedObjectFlag="0">//qualifiedForbidden</objectPath>
        <objectUse>qualifiedForbidden must not be present (schema-qualified rule).</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-UNQUALIFIED">
        <objectPath allowedObjectFlag="0">//unqualifiedForbidden</objectPath>
        <objectUse>unqualifiedForbidden must not be present (unqualified rule).</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

XML_BODY = "\n  <qualifiedForbidden/>\n  <unqualifiedForbidden/>\n"

# Same three attribute-position variants as
# test_xml_processing.py::test_get_schema_from_xml_is_position_independent,
# but on a real <dmodule> root checked end-to-end.
ROOT_ATTRS = {
    "first": f'xsi:noNamespaceSchemaLocation="{SCHEMA}" xmlns:dc="http://www.purl.org/dc/elements/1.1/" id="1"',
    "middle": f'xmlns:dc="http://www.purl.org/dc/elements/1.1/" xsi:noNamespaceSchemaLocation="{SCHEMA}" id="1"',
    "last": f'xmlns:dc="http://www.purl.org/dc/elements/1.1/" id="1" xsi:noNamespaceSchemaLocation="{SCHEMA}"',
}


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


@pytest.mark.parametrize("position", ["first", "middle", "last"])
def test_qualified_and_unqualified_rules_both_fire_regardless_of_attribute_position(
        tmp_path, brex_path, position):
    xml_content = (
        f'<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" {ROOT_ATTRS[position]}>'
        f"{XML_BODY}</dmodule>\n"
    )
    xml_path = tmp_path / f"object_{position}.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    result = checker.validate()

    violated_xpaths = {v["Xpath"] for v in result[brex_path]["0"]}
    assert violated_xpaths == {"//qualifiedForbidden", "//unqualifiedForbidden"}
    assert result["Summary"] == "2 Errors"


def test_all_three_positions_select_the_identical_rule_set(tmp_path, brex_path):
    # The regression this guards against (§3.3): with the schema attribute
    # anywhere but last, the old order-dependent regex extraction failed to
    # find it at all, so `schema` resolved to None/garbage and the
    # rulesContext-qualified rule was silently skipped -- "first"/"middle"
    # would have reported only the unqualified rule, one violation short of
    # "last".
    results = {}
    for position in ("first", "middle", "last"):
        xml_content = (
            f'<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" {ROOT_ATTRS[position]}>'
            f"{XML_BODY}</dmodule>\n"
        )
        xml_path = tmp_path / f"object_{position}.xml"
        xml_path.write_text(xml_content, encoding="utf-8")

        checker = BrexChecker()
        checker.set_xml(str(xml_path))
        checker.override_brex_list([brex_path])
        result = checker.validate()
        results[position] = {v["Xpath"] for v in result[brex_path]["0"]}

    assert results["first"] == results["middle"] == results["last"] == {
        "//qualifiedForbidden", "//unqualifiedForbidden",
    }
