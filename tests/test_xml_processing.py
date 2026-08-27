import pytest

from acd.xml_processing import get_schema_from_xml

_XSI_DECL = 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
_SCHEMA_ATTR = 'xsi:noNamespaceSchemaLocation="schema.xsd"'


@pytest.mark.parametrize("attrs", [
    f'{_XSI_DECL} {_SCHEMA_ATTR} id="1" other="2"',
    f'{_XSI_DECL} id="1" {_SCHEMA_ATTR} other="2"',
    f'{_XSI_DECL} id="1" other="2" {_SCHEMA_ATTR}',
], ids=["first", "middle", "last"])
def test_get_schema_from_xml_is_position_independent(attrs):
    xml_content = f'<root {attrs}/>'
    assert get_schema_from_xml(xml_content) == "schema.xsd"
