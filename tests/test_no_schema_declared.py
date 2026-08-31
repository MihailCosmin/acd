import pytest

from acd.brex_checker import BrexChecker
from acd.brex_checker import NoSchemaDeclared

BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

BREX_CONTENT = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
<objectUse>forbiddenElement must not be present.</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""

# No xsi:noNamespaceSchemaLocation at all -- matches a genuine S1000D <= 3.0 /
# DTD-based object, which never carries that attribute.
NO_SCHEMA_XML = "<dml><forbiddenElement/></dml>"

WITH_SCHEMA_XML = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}"><forbiddenElement/></dml>'
)


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


def _checker_for(tmp_path, brex_path, xml_content, filename="object.xml"):
    xml_path = tmp_path / filename
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def test_no_schema_does_not_raise_by_default(tmp_path, brex_path):
    # Default behaviour is unchanged: set_require_schema is off by default,
    # so a schema-less object (the common case for a genuine S1000D <= 3.0 /
    # DTD-based object) is still checked normally against unqualified rules.
    checker = _checker_for(tmp_path, brex_path, NO_SCHEMA_XML)

    result = checker._check_rules()[brex_path]

    assert len(result['0']) == 1
    assert result['0'][0]['Xpath'] == '//forbiddenElement'


def test_require_schema_raises_for_schema_less_object(tmp_path, brex_path):
    checker = _checker_for(tmp_path, brex_path, NO_SCHEMA_XML)
    checker.set_require_schema(True)

    with pytest.raises(NoSchemaDeclared):
        checker._check_rules()


def test_require_schema_does_not_raise_when_schema_is_declared(tmp_path, brex_path):
    checker = _checker_for(tmp_path, brex_path, WITH_SCHEMA_XML)
    checker.set_require_schema(True)

    result = checker._check_rules()[brex_path]

    assert len(result['0']) == 1


def test_require_schema_can_be_turned_back_off(tmp_path, brex_path):
    checker = _checker_for(tmp_path, brex_path, NO_SCHEMA_XML)
    checker.set_require_schema(True)
    checker.set_require_schema(False)

    result = checker._check_rules()[brex_path]

    assert len(result['0']) == 1


def test_require_schema_off_by_default_on_a_fresh_checker(tmp_path, brex_path):
    checker = _checker_for(tmp_path, brex_path, NO_SCHEMA_XML)
    assert checker._require_schema_declaration is False


def test_require_schema_via_validate(tmp_path, brex_path):
    checker = _checker_for(tmp_path, brex_path, NO_SCHEMA_XML)
    checker.set_require_schema(True)

    with pytest.raises(NoSchemaDeclared):
        checker.validate()
