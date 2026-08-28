import pytest
from lxml import etree

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1" brSeverityLevel="brsl02">
        <brDecisionRef brDecisionIdentNumber="BR-GEN-00001"/>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2" brSeverityLevel="brsl01">
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-3">
        <objectPath allowedObjectFlag="0">][invalid xpath(</objectPath>
        <objectUse>Deliberately malformed rule.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""

SEVERITY_LEVELS_CONTENT = """<brSeverityLevels>
  <brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>
  <brSeverityLevel value="brsl02" fail="no">Warning</brSeverityLevel>
</brSeverityLevels>
"""

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <forbiddenElement/>\n"
    "</dml>\n"
)

@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def severity_levels_path(tmp_path):
    path = tmp_path / ".brseveritylevels"
    path.write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    return str(path)


def _checker(tmp_path, brex_path, xml_content=XML_CONTENT, xml_name="object.xml"):
    xml_path = tmp_path / xml_name
    xml_path.write_text(xml_content, encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def test_report_root_and_document_shape(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    xml_report = checker.to_xml_report(result)
    root = etree.fromstring(xml_report.encode("utf-8"))

    assert root.tag == "brexCheck"
    document = root.find("document")
    assert document.get("path") == checker._xml_path

    brex_node = document.find("brex")
    assert brex_node.get("path") == brex_path


def test_flag_0_error_has_br_decision_ref_object_path_use_and_object(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    errors = root.findall(".//brex/error")
    forbidden_error = next(e for e in errors if e.findtext("objectPath") == "//forbiddenElement")

    assert forbidden_error.get("brSeverityLevel") == "brsl02"
    assert forbidden_error.get("fail") == "no"

    br_decision_ref = forbidden_error.find("brDecisionRef")
    assert br_decision_ref.get("brDecisionIdentNumber") == "BR-GEN-00001"

    object_path = forbidden_error.find("objectPath")
    assert object_path.get("allowedObjectFlag") == "0"
    assert object_path.text == "//forbiddenElement"

    assert forbidden_error.findtext("objectUse") == "forbiddenElement must not be present."

    object_node = forbidden_error.find("object")
    assert object_node is not None
    assert object_node.get("line") == "2"
    assert object_node.get("xpath")
    assert object_node[0].tag == "forbiddenElement"


def test_flag_1_missing_element_has_no_object_child(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    errors = root.findall(".//brex/error")
    required_error = next(e for e in errors if e.findtext("objectPath") == "//requiredElement")

    assert required_error.get("brSeverityLevel") == "brsl01"
    assert required_error.get("fail") is None
    assert required_error.find("brDecisionRef") is None
    assert required_error.find("object") is None
    assert required_error.find("objectPath").get("allowedObjectFlag") == "1"


def test_error_without_severity_level_sets_fail_yes(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//noLevel</objectPath>
        <objectUse>noLevel must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    xml_content = (
        '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "  <noLevel/>\n"
        "</dml>\n"
    )
    checker = _checker(tmp_path, str(brex_path), xml_content=xml_content)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    error = root.find(".//brex/error")
    assert error.get("fail") == "yes"
    assert error.get("brSeverityLevel") is None


def test_xpath_error_reported(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    xpath_error = root.find(".//brex/xpathError")
    assert xpath_error is not None
    assert xpath_error.text == "][invalid xpath("
    assert xpath_error.get("error")


def test_sns_no_errors_node(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
      <snsSystem>
        <snsCode>21</snsCode>
        <snsTitle>Air conditioning</snsTitle>
      </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    xml_content = (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "  <identAndStatusSection>\n"
        "    <dmAddress>\n"
        "      <dmIdent>\n"
        '        <dmCode modelIdentCode="TEST" systemDiffCode="A" '
        'systemCode="21" subSystemCode="0" subSubSystemCode="0" assyCode="00" '
        'disassyCode="00" disassyCodeVariant="A" infoCode="000" '
        'infoCodeVariant="A" itemLocationCode="D"/>\n'
        "      </dmIdent>\n"
        "    </dmAddress>\n"
        "  </identAndStatusSection>\n"
        "</dmodule>\n"
    )
    checker = _checker(tmp_path, str(brex_path), xml_content=xml_content)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    sns_node = root.find(".//document/sns")
    assert sns_node is not None
    assert sns_node.find("noErrors") is not None
    assert sns_node.find("error") is None


def test_sns_error_reported(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
      <snsSystem>
        <snsCode>21</snsCode>
        <snsTitle>Air conditioning</snsTitle>
      </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    xml_content = (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "  <identAndStatusSection>\n"
        "    <dmAddress>\n"
        "      <dmIdent>\n"
        '        <dmCode modelIdentCode="TEST" systemDiffCode="A" '
        'systemCode="99" subSystemCode="00" subSubSystemCode="00" assyCode="00" '
        'disassyCode="00" disassyCodeVariant="A" infoCode="000" '
        'infoCodeVariant="A" itemLocationCode="D"/>\n'
        "      </dmIdent>\n"
        "    </dmAddress>\n"
        "  </identAndStatusSection>\n"
        "</dmodule>\n"
    )
    checker = _checker(tmp_path, str(brex_path), xml_content=xml_content)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    sns_error = root.find(".//document/sns/error")
    assert sns_error is not None
    assert sns_error.findtext("code") == "systemCode"
    assert sns_error.findtext("invalidValue") == "99"


def test_notation_error_reported(tmp_path):
    brex_content = """<brex>
  <notationRuleList>
    <notationRule>
      <notationName allowedNotationFlag="1">cgm</notationName>
      <objectUse>Only CGM graphics are permitted.</objectUse>
    </notationRule>
  </notationRuleList>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    xml_content = (
        "<!DOCTYPE dmodule [\n"
        '<!ENTITY graphic1 SYSTEM "graphic1.tif" NDATA tiff>\n'
        "]>\n"
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "</dmodule>\n"
    )
    checker = _checker(tmp_path, str(brex_path), xml_content=xml_content)
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    notation_error = root.find(".//document/notations/error")
    assert notation_error is not None
    assert notation_error.findtext("invalidNotation") == "tiff"
    assert notation_error.findtext("objectUse") == "Only CGM graphics are permitted."


def test_directory_mode_conversion(tmp_path, brex_path):
    xml_dir = tmp_path / "objects"
    xml_dir.mkdir()
    (xml_dir / "a.xml").write_text(XML_CONTENT, encoding="utf-8")
    (xml_dir / "b.xml").write_text(
        (
            '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
            f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}"/>\n'
        ),
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_xml_dir(str(xml_dir))
    checker.override_brex_list([brex_path])
    result = checker.validate()

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    documents = root.findall("document")
    assert {d.get("path") for d in documents} == {"a.xml", "b.xml"}


def test_skipped_result_produces_empty_document(tmp_path, brex_path):
    empty_path = tmp_path / "empty.xml"
    empty_path.write_text("", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(empty_path))
    checker.override_brex_list([brex_path])
    checker.set_ignore_empty(True)
    result = checker.validate()

    assert result["Skipped"] is True

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    document = root.find("document")
    assert document is not None
    assert document.find("brex") is None
