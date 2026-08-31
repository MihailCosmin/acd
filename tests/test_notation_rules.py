import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"


def make_dm_with_entities(entity_decls: str) -> str:
    return (
        "<!DOCTYPE dmodule [\n"
        f"{entity_decls}"
        "]>\n"
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path_with_notations(tmp_path):
    brex_content = """<brex>
  <notationRuleList>
    <notationRule>
      <notationName allowedNotationFlag="1">cgm</notationName>
      <objectUse>Only CGM graphics are permitted.</objectUse>
    </notationRule>
    <notationRule>
      <notationName allowedNotationFlag="0">tiff</notationName>
      <objectUse>TIFF graphics are not permitted.</objectUse>
    </notationRule>
  </notationRuleList>
</brex>
"""
    path = tmp_path / "brex.xml"
    path.write_text(brex_content, encoding="utf-8")
    return str(path)


def _validate(tmp_path, brex_path, entity_decls):
    xml_content = make_dm_with_entities(entity_decls)
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    # Notation checking is opt-in (`-n` in s1kd-brexcheck); see
    # `validate(check_notations=...)`.
    return checker.validate(check_notations=True)


def test_allowed_notation_produces_no_violation(tmp_path, brex_path_with_notations):
    result = _validate(
        tmp_path, brex_path_with_notations,
        '<!ENTITY graphic1 SYSTEM "graphic1.cgm" NDATA cgm>\n',
    )

    assert result["notations"] == []
    assert result["Summary"] == "0 Errors"


def test_explicitly_disallowed_notation_is_reported(tmp_path, brex_path_with_notations):
    result = _validate(
        tmp_path, brex_path_with_notations,
        '<!ENTITY graphic1 SYSTEM "graphic1.tif" NDATA tiff>\n',
    )

    assert len(result["notations"]) == 1
    assert result["notations"][0]["Entity"] == "graphic1"
    assert result["notations"][0]["Notation"] == "tiff"
    assert result["notations"][0]["Description"] == "Only CGM graphics are permitted."
    assert result["Summary"] == "1 Errors"


def test_unlisted_notation_is_reported_against_first_rule(tmp_path, brex_path_with_notations):
    # "png" is named by no notationRule at all; the C original's fallback
    # XPath always resolves to the first notationRule in document order.
    result = _validate(
        tmp_path, brex_path_with_notations,
        '<!ENTITY graphic1 SYSTEM "graphic1.png" NDATA png>\n',
    )

    assert len(result["notations"]) == 1
    assert result["notations"][0]["Notation"] == "png"
    assert result["notations"][0]["Description"] == "Only CGM graphics are permitted."


def test_internal_and_external_parsed_entities_are_not_checked(tmp_path, brex_path_with_notations):
    result = _validate(
        tmp_path, brex_path_with_notations,
        '<!ENTITY internal "some text">\n'
        '<!ENTITY externalParsed SYSTEM "external.xml">\n',
    )

    assert result["notations"] == []
    assert result["Summary"] == "0 Errors"


def test_multiple_entities_report_multiple_violations(tmp_path, brex_path_with_notations):
    result = _validate(
        tmp_path, brex_path_with_notations,
        '<!ENTITY graphic1 SYSTEM "graphic1.tif" NDATA tiff>\n'
        '<!ENTITY graphic2 SYSTEM "graphic2.cgm" NDATA cgm>\n'
        '<!ENTITY graphic3 SYSTEM "graphic3.png" NDATA png>\n',
    )

    assert len(result["notations"]) == 2
    reported_entities = {entry["Entity"] for entry in result["notations"]}
    assert reported_entities == {"graphic1", "graphic3"}
    assert result["Summary"] == "2 Errors"


def test_no_internal_dtd_subset_is_not_checked(tmp_path, brex_path_with_notations):
    xml_content = (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "</dmodule>\n"
    )
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path_with_notations])

    result = checker.validate(check_notations=True)

    assert result["notations"] == []
    assert result["Summary"] == "0 Errors"


def test_no_notation_rules_defined_reports_with_generic_description(tmp_path):
    brex_content = "<brex></brex>\n"
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    result = _validate(
        tmp_path, str(brex_path),
        '<!ENTITY graphic1 SYSTEM "graphic1.cgm" NDATA cgm>\n',
    )

    assert len(result["notations"]) == 1
    assert result["notations"][0]["Notation"] == "cgm"
    assert result["notations"][0]["Description"] == "Notation 'cgm' is not allowed."
