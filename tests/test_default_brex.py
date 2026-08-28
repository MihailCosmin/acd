from os.path import isfile

import pytest

from acd.brex_checker import BrexChecker
from acd.brex_checker import NoBrexDefined
from acd.default_brex import default_brex_dmc
from acd.default_brex import default_brex_path
from acd.default_brex import find_default_brex_fallback

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

# All seven built-in default BREX, keyed by the logical DMC used throughout
# default_brex.py, alongside the schema substring that selects each one via
# default_brex_dmc (None for "DMC-S1000D-A-...", which is fallback-only).
ALL_BUILTIN_BREX = {
    "DMC-S1000D-H-04-10-0301-00A-022A-D": "S1000D_6",
    "DMC-S1000D-G-04-10-0301-00A-022A-D": "S1000D_5-0",
    "DMC-S1000D-F-04-10-0301-00A-022A-D": "S1000D_4-2",
    "DMC-S1000D-E-04-10-0301-00A-022A-D": "S1000D_4-1",
    "DMC-S1000D-D-04-10-0301-00A-022A-D": "S1000D_4-0",
    "DMC-S1000D-A-04-10-0301-00A-022A-D": None,
    "DMC-AE-A-04-10-0301-00A-022A-D": None,
}


@pytest.mark.parametrize("logical_dmc", list(ALL_BUILTIN_BREX))
def test_default_brex_path_resolves_every_builtin_brex_to_a_real_file(logical_dmc):
    path = default_brex_path(logical_dmc)
    assert path is not None
    assert isfile(path)


def test_default_brex_path_returns_none_for_an_unknown_dmc():
    assert default_brex_path("DMC-NOT-A-BUILTIN-00-00-00-00A-022A-D") is None


@pytest.mark.parametrize(
    "schema,expected_dmc",
    [
        (f"http://www.s1000d.org/S1000D_6/xml_schema_flat/{s}.xsd", "DMC-S1000D-H-04-10-0301-00A-022A-D")
        for s in ["dmodule"]
    ] + [
        ("http://www.s1000d.org/S1000D_5-0/xml_schema_flat/dmodule.xsd", "DMC-S1000D-G-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd", "DMC-S1000D-F-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_4-1-A/xml_schema_flat/dmodule.xsd", "DMC-S1000D-E-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_4-0-1/xml_schema_flat/dmodule.xsd", "DMC-S1000D-D-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_4-0-2/xml_schema_flat/dmodule.xsd", "DMC-S1000D-D-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_3-0/xml_schema_flat/dmodule.xsd", "DMC-AE-A-04-10-0301-00A-022A-D"),
        ("http://www.s1000d.org/S1000D_2-3/xml_schema_flat/dmodule.xsd", "DMC-AE-A-04-10-0301-00A-022A-D"),
        (None, "DMC-S1000D-H-04-10-0301-00A-022A-D"),
        ("", "DMC-S1000D-H-04-10-0301-00A-022A-D"),
    ],
)
def test_default_brex_dmc_selects_by_schema_including_sub_issues(schema, expected_dmc):
    # S1000D_4-0-1/4-0-2 and S1000D_4-1-A are real sub-issue schema locations
    # used by the bundled S1000D-A and S1000D-E BREX themselves; the C
    # original matches them by substring, not exact equality, so a data
    # module declaring one of these sub-issues must still resolve.
    assert default_brex_dmc(schema) == expected_dmc


def test_default_brex_dmc_never_selects_the_fallback_only_issue_a():
    # "DMC-S1000D-A-..." is a valid built-in BREX (reachable via
    # find_default_brex_fallback) but default_brex_dmc must never return it,
    # matching search_brex_fname_from_default_brex in the C original.
    for schema in [s for s in ALL_BUILTIN_BREX.values() if s] + [None, "", "unrecognised schema"]:
        assert default_brex_dmc(schema) != "DMC-S1000D-A-04-10-0301-00A-022A-D"


def _dm_code(model_ident_code: str, system_diff_code: str, issue_info: str = "") -> str:
    return (
        f'<dmCode modelIdentCode="{model_ident_code}" systemDiffCode="{system_diff_code}" systemCode="04" '
        'subSystemCode="1" subSubSystemCode="0" assyCode="0301" disassyCode="00" '
        'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
        f'{issue_info}'
    )


AE_A_REF = {
    "modelIdentCode": "AE", "systemDiffCode": "A", "systemCode": "04",
    "subSystemCode": "1", "subSubSystemCode": "0", "assyCode": "0301",
    "disassyCode": "00", "disassyCodeVariant": "A", "infoCode": "022",
    "infoCodeVariant": "A", "itemLocationCode": "D",
}


def test_find_default_brex_fallback_matches_unversioned_reference():
    ref = dict(AE_A_REF)
    assert find_default_brex_fallback(ref) == "DMC-AE-A-04-10-0301-00A-022A-D"


def test_find_default_brex_fallback_matches_exact_bundled_issue():
    ref = dict(AE_A_REF, issueNumber="003", inWork="00")
    assert find_default_brex_fallback(ref) == "DMC-AE-A-04-10-0301-00A-022A-D"


def test_find_default_brex_fallback_rejects_a_different_issue():
    ref = dict(AE_A_REF, issueNumber="999", inWork="00")
    assert find_default_brex_fallback(ref) is None


def test_find_default_brex_fallback_rejects_a_non_builtin_reference():
    ref = dict(AE_A_REF, modelIdentCode="NOTBUILTIN")
    assert find_default_brex_fallback(ref) is None


def make_object_xml(schema: str, extra: str = "") -> str:
    return f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{schema}">
<content>{extra}</content>
</dmodule>
"""


def test_use_default_brex_checks_against_the_schema_selected_builtin_brex(tmp_path):
    # A legacy S1000D 3.0 schema selects the bundled DMC-AE-A-... BREX (the
    # smallest built-in file), with no brexDmRef needed at all.
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(
        make_object_xml("http://www.s1000d.org/S1000D_3-0/xml_schema_flat/dmodule.xsd"),
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.use_default_brex()

    result = checker.validate()

    expected_path = default_brex_path("DMC-AE-A-04-10-0301-00A-022A-D")
    assert expected_path in result
    assert isinstance(result["Summary"], str)
    assert result["brexFallback"] == []


def test_use_default_brex_ignores_a_brexdmref_to_an_unrelated_brex(tmp_path):
    # -B/--default-brex overrides brexDmRef resolution entirely: even though
    # the object references some other (non-existent) BREX, the schema-
    # selected built-in is used and no NoBrexDefined/BrexNotFound is raised.
    xml_path = tmp_path / "object.xml"
    xml_content = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content/>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{_dm_code("DOESNOTEXIST", "A")}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(tmp_path))  # empty dir: the referenced BREX cannot be found here
    checker.use_default_brex()

    result = checker.validate()

    expected_path = default_brex_path("DMC-S1000D-F-04-10-0301-00A-022A-D")
    assert expected_path in result


def test_referenced_brex_not_on_disk_falls_back_to_builtin_and_is_flagged(tmp_path):
    xml_path = tmp_path / "object.xml"
    xml_content = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content/>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{_dm_code("AE", "A")}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    xml_path.write_text(xml_content, encoding="utf-8")

    empty_brex_dir = tmp_path / "brex"
    empty_brex_dir.mkdir()

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_brex_dir))

    result = checker.validate()

    expected_path = default_brex_path("DMC-AE-A-04-10-0301-00A-022A-D")
    assert expected_path in result
    assert isinstance(result["Summary"], str)

    assert len(result["brexFallback"]) == 1
    fallback = result["brexFallback"][0]
    assert fallback["UsedBuiltinBrex"] == "DMC-AE-A-04-10-0301-00A-022A-D"
    assert fallback["BuiltinBrexPath"] == expected_path
    assert isfile(fallback["BuiltinBrexPath"])


def test_referenced_brex_not_found_and_not_a_builtin_still_raises(tmp_path):
    xml_path = tmp_path / "object.xml"
    xml_content = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content/>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{_dm_code("NOTBUILTIN", "A")}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    xml_path.write_text(xml_content, encoding="utf-8")

    empty_brex_dir = tmp_path / "brex"
    empty_brex_dir.mkdir()

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(empty_brex_dir))

    with pytest.raises(NoBrexDefined):
        checker.validate()


def test_use_default_brex_false_restores_normal_brexdmref_resolution(tmp_path):
    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    brex_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA.replace('dmodule', 'brex')}">
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbidden</objectPath>
        <objectUse>forbidden must not be present</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = brex_dir / "DMC-CUSTOM-A-04-10-0301-00A-022A-D_001-00_EN-US.XML"
    brex_path.write_text(brex_content, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_content = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content><forbidden/></content>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{_dm_code("CUSTOM", "A")}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(brex_dir))
    checker.use_default_brex(True)
    checker.use_default_brex(False)

    result = checker.validate()

    assert str(brex_path) in result
    assert len(result[str(brex_path)]['0']) == 1
    assert result["brexFallback"] == []
