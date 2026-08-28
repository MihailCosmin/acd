import pytest

from acd.brex_checker import BrexChecker

DDN_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/ddn.xsd"
DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = f"""<brex>
  <contextRules rulesContext="{DDN_SCHEMA}">
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present in a DDN</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenEverywhere</objectPath>
        <objectUse>forbiddenEverywhere must not be present in any object</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


def make_xml(schema: str, extra: str = "") -> str:
    return (
        '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{schema}">\n'
        f"{extra}\n"
        "</dml>\n"
    )


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def run_check(tmp_path, brex_path, schema, element):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_xml(schema, element), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker._check_rules()


def test_ddn_qualified_rule_fires_against_ddn_document(tmp_path, brex_path):
    result = run_check(tmp_path, brex_path, DDN_SCHEMA, "<forbiddenElement/>")

    assert len(result[brex_path]['0']) == 1
    assert result[brex_path]['0'][0]['Xpath'] == '//forbiddenElement'


def test_unqualified_rule_fires_against_ddn_document(tmp_path, brex_path):
    result = run_check(tmp_path, brex_path, DDN_SCHEMA, "<forbiddenEverywhere/>")

    assert len(result[brex_path]['0']) == 1
    assert result[brex_path]['0'][0]['Xpath'] == '//forbiddenEverywhere'


def test_ddn_qualified_rule_stays_scoped_to_ddn_schema(tmp_path, brex_path):
    result = run_check(tmp_path, brex_path, DMODULE_SCHEMA, "<forbiddenElement/>")

    assert result[brex_path]['0'] == []
