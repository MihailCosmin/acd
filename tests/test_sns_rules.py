import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"


def make_dm(system_code, sub_system_code, sub_sub_system_code, assy_code="00") -> str:
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "  <identAndStatusSection>\n"
        "    <dmAddress>\n"
        "      <dmIdent>\n"
        '        <dmCode modelIdentCode="TEST" systemDiffCode="A" '
        f'systemCode="{system_code}" subSystemCode="{sub_system_code}" '
        f'subSubSystemCode="{sub_sub_system_code}" assyCode="{assy_code}" '
        'disassyCode="00" disassyCodeVariant="A" infoCode="000" '
        'infoCodeVariant="A" itemLocationCode="D"/>\n'
        "      </dmIdent>\n"
        "    </dmAddress>\n"
        "  </identAndStatusSection>\n"
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path_with_sns(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem id="SNSR-1">
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
            <snsSubSystem>
                <snsCode>0</snsCode>
                <snsTitle>General</snsTitle>
            </snsSubSystem>
            <snsSubSystem>
                <snsCode>1</snsCode>
                <snsTitle>Compression</snsTitle>
                <snsSubSubSystem>
                    <snsCode>0</snsCode>
                    <snsTitle>General</snsTitle>
                </snsSubSubSystem>
                <snsSubSubSystem>
                    <snsCode>1</snsCode>
                    <snsTitle>Compressor</snsTitle>
                </snsSubSubSystem>
            </snsSubSystem>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    path = tmp_path / "brex.xml"
    path.write_text(brex_content, encoding="utf-8")
    return str(path)


def _validate(tmp_path, brex_path, system_code, sub_system_code, sub_sub_system_code, assy_code="00", sns_mode="normal"):
    xml_content = make_dm(system_code, sub_system_code, sub_sub_system_code, assy_code)
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])

    return checker.validate(sns_mode=sns_mode)


def test_sns_valid_code_produces_no_violation(tmp_path, brex_path_with_sns):
    result = _validate(tmp_path, brex_path_with_sns, "21", "1", "1")

    assert result["sns"] == []
    assert result["Summary"] == "0 Errors"


def test_sns_invalid_system_code_is_reported(tmp_path, brex_path_with_sns):
    result = _validate(tmp_path, brex_path_with_sns, "99", "1", "1")

    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "systemCode"
    assert result["sns"][0]["invalidValue"] == "99"
    assert result["Summary"] == "1 Errors"


def test_sns_invalid_sub_sub_system_code_is_reported_and_stops_at_first_failure(tmp_path, brex_path_with_sns):
    # subSystemCode "1" is valid, but subSubSystemCode "9" is not defined under it.
    result = _validate(tmp_path, brex_path_with_sns, "21", "1", "9")

    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "subSubSystemCode"
    assert result["sns"][0]["invalidValue"] == "21-19"
    assert result["Summary"] == "1 Errors"


def test_sns_placeholder_codes_are_skipped_when_no_rules_defined_at_that_level(tmp_path, brex_path_with_sns):
    # subSubSystemCode "0" is a placeholder; snsSubSystem "0" has no snsSubSubSystem
    # children defined at all, so the level is skipped rather than flagged.
    result = _validate(tmp_path, brex_path_with_sns, "21", "0", "0")

    assert result["sns"] == []
    assert result["Summary"] == "0 Errors"


def test_sns_strict_mode_rejects_placeholder_not_defined_at_that_level(tmp_path, brex_path_with_sns):
    # In normal mode this exact input (21, 0, 0) produces no violation because
    # snsSubSystem "0" defines no snsSubSubSystem children at all, so the
    # placeholder subSubSystemCode "0" is skipped (see the "placeholder codes
    # are skipped" test above). Strict mode has no shorthand: "0" must itself
    # match a defined snsCode, which it does not.
    result = _validate(tmp_path, brex_path_with_sns, "21", "0", "0", sns_mode="strict")

    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "subSubSystemCode"
    assert result["sns"][0]["invalidValue"] == "21-00"
    assert result["Summary"] == "1 Errors"


def test_sns_strict_mode_rejects_undefined_assy_level(tmp_path, brex_path_with_sns):
    # (21, 1, 1) is fully valid in normal mode (see test_sns_valid_code_produces_no_violation)
    # because the fixture defines no snsAssy rules anywhere, so the placeholder
    # assyCode "00" is skipped. Strict mode checks it anyway.
    result = _validate(tmp_path, brex_path_with_sns, "21", "1", "1", sns_mode="strict")

    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "assyCode"
    assert result["sns"][0]["invalidValue"] == "21-11-00"
    assert result["Summary"] == "1 Errors"


def test_sns_unstrict_mode_accepts_any_code_when_level_omitted(tmp_path, brex_path_with_sns):
    # subSubSystemCode "5" is not a placeholder, so normal mode checks it and
    # fails, since snsSubSystem "0" defines no snsSubSubSystem children.
    normal_result = _validate(tmp_path, brex_path_with_sns, "21", "0", "5", sns_mode="normal")
    assert len(normal_result["sns"]) == 1
    assert normal_result["sns"][0]["code"] == "subSubSystemCode"

    # Unstrict mode only checks a level the BREX actually defines rules for;
    # since snsSubSystem "0" defines none, any subSubSystemCode is accepted.
    unstrict_result = _validate(tmp_path, brex_path_with_sns, "21", "0", "5", sns_mode="unstrict")
    assert unstrict_result["sns"] == []
    assert unstrict_result["Summary"] == "0 Errors"


def test_sns_invalid_mode_raises(tmp_path, brex_path_with_sns):
    with pytest.raises(ValueError):
        _validate(tmp_path, brex_path_with_sns, "21", "1", "1", sns_mode="bogus")


def test_sns_not_checked_for_non_dmodule_objects(tmp_path, brex_path_with_sns):
    xml_content = (
        '<pm xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        "</pm>\n"
    )
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path_with_sns])

    result = checker.validate()

    assert "sns" not in result
    assert result["Summary"] == "0 Errors"
