import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

FLAG_0_BREX = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

FLAG_1_BREX = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


def make_xml(extra: str) -> str:
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        f"{extra}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def flag_0_brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(FLAG_0_BREX, encoding="utf-8")
    return str(path)


@pytest.fixture
def flag_1_brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(FLAG_1_BREX, encoding="utf-8")
    return str(path)


def run_check(tmp_path, brex_path, extra, remove_deleted=False):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_xml(extra), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker._check_rules(remove_deleted=remove_deleted)


def test_deleted_element_still_flagged_by_default(tmp_path, flag_0_brex_path):
    result = run_check(tmp_path, flag_0_brex_path,
                        '<forbiddenElement changeType="delete"/>')

    assert len(result[flag_0_brex_path]['0']) == 1


def test_deleted_element_dropped_before_flag_0_check(tmp_path, flag_0_brex_path):
    result = run_check(tmp_path, flag_0_brex_path,
                        '<forbiddenElement changeType="delete"/>',
                        remove_deleted=True)

    assert result[flag_0_brex_path]['0'] == []


def test_non_deleted_element_still_flagged_with_remove_deleted_enabled(tmp_path, flag_0_brex_path):
    result = run_check(tmp_path, flag_0_brex_path,
                        '<forbiddenElement/>',
                        remove_deleted=True)

    assert len(result[flag_0_brex_path]['0']) == 1


def test_legacy_change_attribute_spelling_is_also_removed(tmp_path, flag_0_brex_path):
    result = run_check(tmp_path, flag_0_brex_path,
                        '<forbiddenElement change="delete"/>',
                        remove_deleted=True)

    assert result[flag_0_brex_path]['0'] == []


def test_deleted_parent_removes_its_children_too(tmp_path, flag_0_brex_path):
    result = run_check(tmp_path, flag_0_brex_path,
                        '<wrapper changeType="delete"><forbiddenElement/></wrapper>',
                        remove_deleted=True)

    assert result[flag_0_brex_path]['0'] == []


def test_deleted_required_element_becomes_a_violation(tmp_path, flag_1_brex_path):
    present = run_check(tmp_path, flag_1_brex_path,
                         '<requiredElement changeType="delete"/>',
                         remove_deleted=False)
    assert present[flag_1_brex_path]['1'] == []

    removed = run_check(tmp_path, flag_1_brex_path,
                         '<requiredElement changeType="delete"/>',
                         remove_deleted=True)
    assert len(removed[flag_1_brex_path]['1']) == 1


def test_remove_deleted_threaded_through_validate(tmp_path, flag_0_brex_path):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_xml('<forbiddenElement changeType="delete"/>'), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([flag_0_brex_path])

    default_result = checker.validate()
    assert default_result["Summary"] == "1 Errors"

    removed_result = checker.validate(remove_deleted=True)
    assert removed_result["Summary"] == "0 Errors"
