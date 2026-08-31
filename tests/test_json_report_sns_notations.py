"""`to_json_report` must carry everything its own `summary` counts.

`violations()` only represents content-rule violations (an SNS or notation
violation has no `objectPath`/`objectUse`/allowed-values shape to fit
`BrexViolation`), so an object whose only findings are SNS/notation ones used
to produce a report whose `summary` said "5 Errors" next to `"violations":
[]`, with nothing naming the five. `to_xml_report` never had that problem --
it emits `sns`/`notations` nodes per `document` -- so these tests pin the JSON
report to the same content.
"""

import json

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"


@pytest.fixture
def brex_path(tmp_path):
    """A BREX carrying SNS rules, notation rules and a nonContextRule, but no
    content rule at all -- so every finding lands outside `violations()`.
    """
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
      <snsSystem id="SNSR-1">
        <snsCode>21</snsCode>
        <snsTitle>Air conditioning</snsTitle>
        <snsSubSystem>
          <snsCode>1</snsCode>
          <snsTitle>Compression</snsTitle>
        </snsSubSystem>
      </snsSystem>
    </snsDescr>
  </snsRules>
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
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BR-NC-1"/>
      <simplePara>Illustrations must be approved by the publications authority.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    path = tmp_path / "brex.xml"
    path.write_text(brex_content, encoding="utf-8")
    return str(path)


def _write_dm(path, system_code, notation):
    path.write_text(
        "<!DOCTYPE dmodule [\n"
        f'<!ENTITY graphic1 SYSTEM "graphic1.bin" NDATA {notation}>\n'
        "]>\n"
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "  <identAndStatusSection>\n"
        "    <dmAddress>\n"
        "      <dmIdent>\n"
        '        <dmCode modelIdentCode="TEST" systemDiffCode="A" '
        f'systemCode="{system_code}" subSystemCode="1" '
        'subSubSystemCode="0" assyCode="00" '
        'disassyCode="00" disassyCodeVariant="A" infoCode="000" '
        'infoCodeVariant="A" itemLocationCode="D"/>\n'
        "      </dmIdent>\n"
        "    </dmAddress>\n"
        "  </identAndStatusSection>\n"
        "</dmodule>\n",
        encoding="utf-8",
    )
    return str(path)


def _report(checker, result):
    return json.loads(checker.to_json_report(result))


def test_sns_and_notation_violations_back_the_summary_counts(tmp_path, brex_path):
    xml_path = _write_dm(tmp_path / "object.xml", "99", "tiff")

    checker = BrexChecker()
    checker.set_xml(xml_path)
    checker.override_brex_list([brex_path])
    result = checker.validate(check_sns=True, check_notations=True)

    report = _report(checker, result)

    # Two findings, neither of them a content-rule violation: without the
    # sns/notations sections the report would count them and list nothing.
    assert report["summary"]["Errors"] == 2
    assert report["violations"] == []

    assert report["sns"] == [{
        "document": xml_path,
        "code": "systemCode",
        "invalidValue": "99",
        "objectUse": "systemCode is not valid according to the SNS rules.",
    }]
    assert report["notations"] == [{
        "document": xml_path,
        "entity": "graphic1",
        "invalidNotation": "tiff",
        "objectUse": "Only CGM graphics are permitted.",
    }]


def test_non_context_rules_are_reported_like_the_xml_report_node(tmp_path, brex_path):
    xml_path = _write_dm(tmp_path / "object.xml", "21", "cgm")

    checker = BrexChecker()
    checker.set_xml(xml_path)
    checker.override_brex_list([brex_path])
    result = checker.validate(check_sns=True, check_notations=True)

    report = _report(checker, result)

    # Informational, not violations -- present even on a clean object.
    assert report["summary"]["Errors"] == 0
    assert report["sns"] == []
    assert report["notations"] == []
    assert report["nonContextRules"] == [{
        "document": xml_path,
        "brex": brex_path,
        "brDecisionIdentNumber": "BR-NC-1",
        "text": "Illustrations must be approved by the publications authority.",
    }]
    assert report["brexFallback"] == []


def test_directory_mode_tags_each_entry_with_its_own_document(tmp_path, brex_path):
    objects_dir = tmp_path / "objects"
    objects_dir.mkdir()
    _write_dm(objects_dir / "bad.xml", "99", "tiff")
    _write_dm(objects_dir / "good.xml", "21", "cgm")

    checker = BrexChecker()
    checker.set_xml_dir(str(objects_dir))
    checker.override_brex_list([brex_path])
    result = checker.validate(check_sns=True, check_notations=True)

    report = _report(checker, result)

    assert report["summary"]["Errors"] == 2
    # Documents are named the same way `to_xml_report` names them in
    # `document/@path` -- the result's own keys, i.e. bare filenames here.
    # A checked-and-passed object contributes no entry, the JSON equivalent
    # of the XML report's `<noErrors/>` child.
    assert [entry["document"] for entry in report["sns"]] == ["bad.xml"]
    assert [entry["document"] for entry in report["notations"]] == ["bad.xml"]
    assert {entry["document"] for entry in report["nonContextRules"]} == {"bad.xml", "good.xml"}
