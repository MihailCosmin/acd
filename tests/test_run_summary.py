import pytest
from lxml import etree

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1" brSeverityLevel="brsl02">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2" brSeverityLevel="brsl01">
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
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

PASSING_XML = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <requiredElement/>\n"
    "</dml>\n"
)

FAILING_XML = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "</dml>\n"
)

WARNING_ONLY_XML = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
    "  <forbiddenElement/>\n"
    "  <requiredElement/>\n"
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


def _checker(tmp_path, brex_path, xml_content, xml_name="object.xml"):
    xml_path = tmp_path / xml_name
    xml_path.write_text(xml_content, encoding="utf-8")
    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def test_single_passing_document(tmp_path, brex_path):
    checker = _checker(tmp_path, brex_path, PASSING_XML)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals == {
        "DocumentsChecked": 1,
        "DocumentsPassed": 1,
        "DocumentsFailed": 0,
        "DocumentsSkipped": 0,
        "Errors": 0,
        "Warnings": 0,
        "ViolationsBySeverity": {},
    }


def test_single_failing_document_tallies_severity(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path, FAILING_XML)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals["DocumentsChecked"] == 1
    assert totals["DocumentsFailed"] == 1
    assert totals["DocumentsPassed"] == 0
    assert totals["Errors"] == 1
    assert totals["ViolationsBySeverity"] == {"brsl01": 1}


def test_warning_only_document_still_passes(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path, WARNING_ONLY_XML)
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals["DocumentsPassed"] == 1
    assert totals["DocumentsFailed"] == 0
    assert totals["Errors"] == 0
    assert totals["Warnings"] == 1
    assert totals["ViolationsBySeverity"] == {"brsl02": 1}


def test_sns_and_notation_violations_counted_as_errors_with_no_severity(tmp_path):
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
    checker = _checker(tmp_path, str(brex_path), xml_content)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals["DocumentsFailed"] == 1
    assert totals["Errors"] == 1
    assert totals["ViolationsBySeverity"] == {None: 1}


def test_directory_mode_aggregates_across_documents(tmp_path, brex_path, severity_levels_path):
    xml_dir = tmp_path / "objects"
    xml_dir.mkdir()
    (xml_dir / "pass.xml").write_text(PASSING_XML, encoding="utf-8")
    (xml_dir / "fail.xml").write_text(FAILING_XML, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(xml_dir))
    checker.override_brex_list([brex_path])
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals["DocumentsChecked"] == 2
    assert totals["DocumentsPassed"] == 1
    assert totals["DocumentsFailed"] == 1
    assert totals["Errors"] == 1
    assert totals["ViolationsBySeverity"] == {"brsl01": 1}


def test_skipped_document_counted_separately(tmp_path, brex_path):
    empty_path = tmp_path / "empty.xml"
    empty_path.write_text("", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(empty_path))
    checker.override_brex_list([brex_path])
    checker.set_ignore_empty(True)
    result = checker.validate()

    totals = checker.run_summary(result)
    assert totals == {
        "DocumentsChecked": 0,
        "DocumentsPassed": 0,
        "DocumentsFailed": 0,
        "DocumentsSkipped": 1,
        "Errors": 0,
        "Warnings": 0,
        "ViolationsBySeverity": {},
    }


def test_xml_report_summary_node_matches_run_summary(tmp_path, brex_path, severity_levels_path):
    xml_dir = tmp_path / "objects"
    xml_dir.mkdir()
    (xml_dir / "pass.xml").write_text(PASSING_XML, encoding="utf-8")
    (xml_dir / "fail.xml").write_text(FAILING_XML, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(xml_dir))
    checker.override_brex_list([brex_path])
    checker.set_severity_levels_path(severity_levels_path)
    result = checker.validate()

    totals = checker.run_summary(result)
    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))

    assert root[0].tag == "summary"
    summary_node = root.find("summary")
    assert summary_node.get("documentsChecked") == str(totals["DocumentsChecked"])
    assert summary_node.get("documentsPassed") == str(totals["DocumentsPassed"])
    assert summary_node.get("documentsFailed") == str(totals["DocumentsFailed"])
    assert summary_node.get("errors") == str(totals["Errors"])

    severity_node = summary_node.find("severity")
    assert severity_node.get("value") == "brsl01"
    assert severity_node.get("count") == "1"
