import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"


def make_xml(schema: str, extra: str = "") -> str:
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{schema}">\n'
        f"{extra}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path_multi(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenA</objectPath>
        <objectUse>forbiddenA must not be present</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenB</objectPath>
        <objectUse>forbiddenB must not be present</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredA</objectPath>
        <objectUse>requiredA must be present</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredB</objectPath>
        <objectUse>requiredB must be present</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//constrainedA</objectPath>
        <objectUse>constrainedA must equal one of the allowed values</objectUse>
        <objectValue valueForm="single" valueAllowed="allowedValue"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    path = tmp_path / "brex.xml"
    path.write_text(brex_content, encoding="utf-8")
    return str(path)


def test_append_summary_counts_every_violation_across_flags(tmp_path, brex_path_multi):
    xml_content = make_xml(
        DMODULE_SCHEMA,
        "<forbiddenA/>\n<forbiddenB/>\n<constrainedA>wrongValue</constrainedA>",
    )
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path_multi])

    result = checker.validate()

    assert len(result[brex_path_multi]['0']) == 2
    assert len(result[brex_path_multi]['1']) == 2
    assert len(result[brex_path_multi]['2']) == 1
    assert result["Summary"] == "5 Errors"


@pytest.fixture
def brex_path_with_malformed_rule(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenA</objectPath>
        <objectUse>forbiddenA must not be present</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//[[[malformed</objectPath>
        <objectUse>this rule cannot be evaluated</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    path = tmp_path / "brex.xml"
    path.write_text(brex_content, encoding="utf-8")
    return str(path)


def test_append_summary_excludes_xpath_errors_from_the_violation_count(tmp_path, brex_path_with_malformed_rule):
    xml_content = make_xml(DMODULE_SCHEMA, "<forbiddenA/>")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path_with_malformed_rule])

    result = checker.validate()

    assert len(result[brex_path_with_malformed_rule]['0']) == 1
    assert len(result[brex_path_with_malformed_rule]['xpathError']) == 1
    assert result["Summary"] == "1 Errors"
