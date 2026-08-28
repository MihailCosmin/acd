import pytest

from acd.brex_checker import BrexChecker
from acd.brex_checker import BrexNotFound
from acd.brex_checker import NoBrexDefined
from acd.s1000d import find_document_by_reference

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"
BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

REAL_DM_CODE = (
    '<dmCode modelIdentCode="REALBX" systemDiffCode="A" systemCode="00" '
    'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
    'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
)

BREX_FNAME = "DMC-REALBX-A-00-00-00-00A-022A-D_001-00_EN-US.XML"

BREX_CONTENT = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{REAL_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbiddenReal</objectPath>
<objectUse>forbiddenReal must not be present</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""


def make_object_xml(extra: str = "") -> str:
    return f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content>
{extra}
</content>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{REAL_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""


def test_add_brex_search_path_finds_brex_not_in_primary_dir(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    other_dir = tmp_path / "other"
    other_dir.mkdir()
    brex_path = other_dir / BREX_FNAME
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_dir))
    checker.add_brex_search_path(str(other_dir))

    result = checker.validate()

    assert str(brex_path) in result
    assert len(result[str(brex_path)]['0']) == 1
    assert result[str(brex_path)]['0'][0]['Xpath'] == '//forbiddenReal'


def test_multiple_search_paths_tried_in_order_until_a_match(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    first_dir = tmp_path / "first"
    first_dir.mkdir()
    second_dir = tmp_path / "second"
    second_dir.mkdir()

    # The brex only lives in the second added search path; the first stays empty.
    brex_path = second_dir / BREX_FNAME
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_dir))
    checker.add_brex_search_path(str(first_dir))
    checker.add_brex_search_path(str(second_dir))

    result = checker.validate()

    assert str(brex_path) in result


def test_add_brex_search_path_rejects_a_file(tmp_path):
    not_a_dir = tmp_path / "not_a_dir.txt"
    not_a_dir.write_text("x", encoding="utf-8")

    checker = BrexChecker()
    with pytest.raises(BrexNotFound):
        checker.add_brex_search_path(str(not_a_dir))


def test_clear_brex_search_paths_removes_previously_added_paths(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    other_dir = tmp_path / "other"
    other_dir.mkdir()
    (other_dir / BREX_FNAME).write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_dir))
    checker.add_brex_search_path(str(other_dir))
    checker.clear_brex_search_paths()

    with pytest.raises(NoBrexDefined):
        checker.validate()


def test_recursive_search_enabled_by_default_finds_nested_brex(tmp_path):
    search_dir = tmp_path / "search"
    nested_dir = search_dir / "nested"
    nested_dir.mkdir(parents=True)
    brex_path = nested_dir / BREX_FNAME
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(search_dir))

    result = checker.validate()

    assert str(brex_path) in result


def test_recursive_search_disabled_ignores_nested_brex(tmp_path):
    search_dir = tmp_path / "search"
    nested_dir = search_dir / "nested"
    nested_dir.mkdir(parents=True)
    (nested_dir / BREX_FNAME).write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(search_dir))
    checker.set_brex_recursive_search(False)

    with pytest.raises(NoBrexDefined):
        checker.validate()


def test_recursive_search_flag_also_applies_to_added_search_paths(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    search_dir = tmp_path / "search"
    nested_dir = search_dir / "nested"
    nested_dir.mkdir(parents=True)
    brex_path = nested_dir / BREX_FNAME
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_dir))
    checker.add_brex_search_path(str(search_dir))
    checker.set_brex_recursive_search(False)

    with pytest.raises(NoBrexDefined):
        checker.validate()

    # _init_brex_list caches the (empty) resolution from the call above, so a
    # fresh checker is needed to re-resolve with recursive search back on.
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_dir))
    checker.add_brex_search_path(str(search_dir))
    checker.set_brex_recursive_search(True)

    result = checker.validate()
    assert str(brex_path) in result


def test_find_document_by_reference_recursive_by_default(tmp_path):
    nested = tmp_path / "a" / "b"
    nested.mkdir(parents=True)
    target = nested / BREX_FNAME
    target.write_text("x", encoding="utf-8")

    found = find_document_by_reference("DMC-REALBX-A-00-00-00-00A-022A-D", str(tmp_path))

    assert found == str(target)


def test_find_document_by_reference_non_recursive_only_looks_directly_inside(tmp_path):
    nested = tmp_path / "a"
    nested.mkdir()
    target = nested / BREX_FNAME
    target.write_text("x", encoding="utf-8")

    not_found = find_document_by_reference("DMC-REALBX-A-00-00-00-00A-022A-D", str(tmp_path), recursive=False)
    found_directly = find_document_by_reference("DMC-REALBX-A-00-00-00-00A-022A-D", str(nested), recursive=False)

    assert not_found is None
    assert found_directly == str(target)
