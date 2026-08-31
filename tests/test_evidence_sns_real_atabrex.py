"""SNS fixture for brex_checker_rework.md §4.5, exercising normal/strict/
unstrict modes against the real
`DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML` (the ATA CMP BREX from
the plan's evidence base, 749 real `snsCode` entries -- see §1).

Skipped when the evidence corpus is not present on disk (it lives outside
this repository); override its location with `ACD_BREX_EVIDENCE_DIR` if it
is available somewhere other than the default path used throughout the plan
document. The specific systemCode/subSystemCode combinations exercised here
were read directly out of the real file (systemCode "00" -> subSystemCode
"1" defines no snsSubSubSystem children at all; subSystemCode "4" does; no
BREX in the evidence base defines any snsAssy), not fabricated, so this
doubles as a check that real-world SNS tables with partially-populated
levels are handled correctly.
"""

import os
from os.path import isfile, join

import pytest

from acd.brex_checker import BrexChecker

EVIDENCE_DIR = os.environ.get(
    "ACD_BREX_EVIDENCE_DIR",
    r"C:\Users\munte\Develop\TD\SITEC\Seventh Delivery\CMP 21-77-05",
)
ATABREX_01A = join(EVIDENCE_DIR, "DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML")

pytestmark = [
    pytest.mark.evidence,
    pytest.mark.skipif(
        not isfile(ATABREX_01A),
        reason="real CMP 21-77-05 evidence folder not available on this machine "
               "(set ACD_BREX_EVIDENCE_DIR to point at it)",
    ),
]

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


def _validate(tmp_path, system_code, sub_system_code, sub_sub_system_code, assy_code="00", sns_mode="normal"):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_dm(system_code, sub_system_code, sub_sub_system_code, assy_code), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([ATABREX_01A])

    return checker.validate(sns_mode=sns_mode)


def test_normal_mode_accepts_a_fully_placeholder_code(tmp_path):
    # subSystem "1" has no snsSubSubSystem children and no BREX in the
    # evidence base defines snsAssy at all -- both placeholder levels are
    # skipped under the normal-mode shorthand.
    result = _validate(tmp_path, "00", "1", "0", assy_code="00", sns_mode="normal")
    assert result["sns"] == []


def test_normal_mode_rejects_an_undefined_system_code(tmp_path):
    result = _validate(tmp_path, "99", "1", "0", sns_mode="normal")
    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "systemCode"
    assert result["sns"][0]["invalidValue"] == "99"


def test_normal_mode_rejects_a_non_placeholder_subsubsystem_where_none_are_defined(tmp_path):
    # subSubSystemCode "5" is not a placeholder, so normal mode checks it
    # even though subSystem "1" defines no snsSubSubSystem children at all.
    result = _validate(tmp_path, "00", "1", "5", sns_mode="normal")
    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "subSubSystemCode"
    assert result["sns"][0]["invalidValue"] == "00-15"


def test_strict_mode_rejects_the_placeholder_shorthand_normal_mode_accepts(tmp_path):
    # Identical input to test_normal_mode_accepts_a_fully_placeholder_code:
    # strict mode has no shorthand, so subSubSystemCode "0" must itself
    # match a defined snsCode under subSystem "1" -- which, per the real
    # file, defines none.
    result = _validate(tmp_path, "00", "1", "0", assy_code="00", sns_mode="strict")
    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "subSubSystemCode"
    assert result["sns"][0]["invalidValue"] == "00-10"


def test_strict_mode_fails_on_assycode_even_along_a_fully_defined_path(tmp_path):
    # subSystem "4" DOES define snsSubSubSystem "0", so strict mode passes
    # both of those levels -- but no BREX in the evidence base defines any
    # snsAssy at all, so strict mode (which always checks every level,
    # including the "00" placeholder) can never pass the assyCode level
    # against this real file.
    result = _validate(tmp_path, "00", "4", "0", assy_code="00", sns_mode="strict")
    assert len(result["sns"]) == 1
    assert result["sns"][0]["code"] == "assyCode"
    assert result["sns"][0]["invalidValue"] == "00-40-00"


def test_unstrict_mode_accepts_any_code_where_the_level_defines_no_rules(tmp_path):
    # Same (systemCode, subSystemCode, subSubSystemCode) as
    # test_normal_mode_rejects_a_non_placeholder_subsubsystem_where_none_are_defined,
    # which normal mode rejects -- unstrict mode skips a level entirely (any
    # code accepted) once the current scope defines no rules for it at all.
    result = _validate(tmp_path, "00", "1", "5", sns_mode="unstrict")
    assert result["sns"] == []
