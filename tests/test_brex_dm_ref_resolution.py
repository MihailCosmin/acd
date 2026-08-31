import threading

import pytest

from acd.brex_checker import BrexChecker
from acd.s1000d import get_brex_ref
from acd.s1000d import ref_dict_to_str

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"
BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

# An unrelated dmRef that happens to carry infoCode 022 (the BREX infoCode) and
# appears earlier in the document than the real brexDmRef. Under the old "first
# dmRef with infoCode 022" lookup this decoy would hijack BREX resolution.
DECOY_DM_CODE = (
    '<dmCode modelIdentCode="DECOYBX" systemDiffCode="Z" systemCode="99" '
    'subSystemCode="9" subSubSystemCode="9" assyCode="99" disassyCode="99" '
    'disassyCodeVariant="Z" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
)
REAL_DM_CODE = (
    '<dmCode modelIdentCode="REALBX" systemDiffCode="A" systemCode="00" '
    'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
    'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
)


def make_object_xml(extra: str = "") -> str:
    return f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content>
<refs>
<dmRef><dmRefIdent>{DECOY_DM_CODE}</dmRefIdent></dmRef>
</refs>
{extra}
</content>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{REAL_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""


def test_get_brex_ref_prefers_brexdmref_over_an_earlier_infocode_022_dmref(tmp_path):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml(), encoding="utf-8")

    ref = get_brex_ref(str(xml_path))

    assert ref["modelIdentCode"] == "REALBX"
    assert ref_dict_to_str(ref) == "DMC-REALBX-A-00-00-00-00A-022A-D"


def test_get_brex_ref_returns_none_without_a_dedicated_brex_element(tmp_path):
    xml_content = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<refs>
<dmRef><dmRefIdent>{DECOY_DM_CODE}</dmRefIdent></dmRef>
</refs>
</dmodule>
"""
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    assert get_brex_ref(str(xml_path)) is None


def test_get_brex_ref_supports_legacy_brexref_avee(tmp_path):
    xml_content = """<!-- s1000d S1000D_2-3/ -->
<dmodule>
<refs>
<refdm><avee>
<modelic>DECOY</modelic><sdc>Z</sdc><chapnum>99</chapnum><section>9</section>
<subsect>9</subsect><subject>99</subject><discode>99</discode><discodev>Z</discodev>
<incode>022</incode><incodev>A</incodev><itemloc>D</itemloc>
</avee></refdm>
</refs>
<idstatus>
<brexref><refdm><avee>
<modelic>REALBREX</modelic><sdc>A</sdc><chapnum>00</chapnum><section>0</section>
<subsect>0</subsect><subject>00</subject><discode>00</discode><discodev>A</discodev>
<incode>022</incode><incodev>A</incodev><itemloc>D</itemloc>
</avee></refdm></brexref>
</idstatus>
</dmodule>
"""
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    ref = get_brex_ref(str(xml_path))

    assert ref["modelIdentCode"] == "REALBREX"
    assert ref_dict_to_str(ref) == "DMC-REALBREX-A-00-00-00-00A-022A-D"


@pytest.fixture
def brex_dir_with_self_referencing_brex(tmp_path):
    brex_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
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
    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    brex_path = brex_dir / "DMC-REALBX-A-00-00-00-00A-022A-D_001-00_EN-US.XML"
    brex_path.write_text(brex_content, encoding="utf-8")
    return str(brex_dir), str(brex_path)


def test_init_brex_list_resolves_the_real_brex_ignoring_the_decoy_dmref(tmp_path, brex_dir_with_self_referencing_brex):
    # The decoy DMC referenced by the earlier dmRef does not exist on disk: were
    # resolution to still follow that dmRef (the old behaviour), it would raise
    # NoBrexDefined instead of finding and checking the real BREX below.
    brex_dir, brex_path = brex_dir_with_self_referencing_brex

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_object_xml("<forbiddenReal/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(brex_dir)

    result = checker.validate()

    assert brex_path in result
    assert len(result[brex_path]['0']) == 1
    assert result[brex_path]['0'][0]['Xpath'] == '//forbiddenReal'


def test_init_brex_list_resolves_self_referencing_brex_checked_directly(brex_dir_with_self_referencing_brex):
    # Ref §3.11: once BREX data modules are checked like any other object
    # (no longer excluded from directory mode's file filter), a
    # self-referencing master/default BREX can itself be the primary
    # checked object. Its own brexDmRef resolves to itself on the very
    # first lookup in the walk, which must resolve to "check it against
    # itself" (mirrors s1kd's `strcmp(brex_fnames[0], dmod_fnames[i]) == 0`
    # case in `main()`) rather than an empty, unresolved chain -- which
    # would raise NoBrexDefined for every self-referencing BREX in a CSDB.
    brex_dir, brex_path = brex_dir_with_self_referencing_brex

    checker = BrexChecker()
    checker.set_xml(brex_path)
    checker.set_brex_path(brex_dir)
    checker._init_brex_list()

    assert checker._brex_list[0] == [brex_path]


def test_validate_checks_a_self_referencing_brex_against_itself_without_raising(brex_dir_with_self_referencing_brex):
    brex_dir, brex_path = brex_dir_with_self_referencing_brex

    checker = BrexChecker()
    checker.set_xml(brex_path)
    checker.set_brex_path(brex_dir)
    result = checker.validate()

    # //forbiddenReal is the BREX's own rule; the BREX document itself (a
    # <brex> root, not <dmodule>) contains no such element, so it passes.
    assert result[brex_path]['0'] == []
    assert result["Summary"] == "0 Errors"


def _dm_code(model_ident_code: str) -> str:
    return (
        f'<dmCode modelIdentCode="{model_ident_code}" systemDiffCode="A" systemCode="00" '
        'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
        'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
    )


def _filename_for(model_ident_code: str) -> str:
    return f"DMC-{model_ident_code}-A-00-00-00-00A-022A-D_001-00_EN-US.XML"


def test_init_brex_list_cycle_guard_prevents_infinite_loop(tmp_path):
    # BREX A references BREX B, and BREX B references BREX A back: neither is
    # self-referencing, so the old "stop when a brex references itself" check
    # never fires and the walk would loop between the two files forever.
    cycle_a_code = _dm_code("CYCLEA1")
    cycle_b_code = _dm_code("CYCLEB1")

    object_xml = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content/>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{cycle_a_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    brex_a_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{cycle_b_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules/>
</brex>
"""
    brex_b_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{cycle_a_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules/>
</brex>
"""

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(object_xml, encoding="utf-8")

    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    brex_a_path = brex_dir / _filename_for("CYCLEA1")
    brex_b_path = brex_dir / _filename_for("CYCLEB1")
    brex_a_path.write_text(brex_a_content, encoding="utf-8")
    brex_b_path.write_text(brex_b_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(brex_dir))

    thread = threading.Thread(target=checker._init_brex_list, daemon=True)
    thread.start()
    thread.join(timeout=5)

    assert not thread.is_alive(), "_init_brex_list did not terminate: cycle guard is missing"
    assert str(brex_a_path) in checker._brex_list[0]
    assert str(brex_b_path) in checker._brex_list[0]


def test_init_brex_list_handles_a_layered_brex_with_no_further_ref(tmp_path):
    # BREX A is not self-referencing (its own DMC isn't in its filename) but
    # also carries no brexDmRef/brexref of its own, so get_brex_ref(brex_a)
    # returns None. Resolving that with ref_dict_to_str used to raise a
    # TypeError instead of just ending the walk.
    dead_end_code = _dm_code("DEADEND")

    object_xml = f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content/>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{dead_end_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
"""
    brex_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<contextRules/>
</brex>
"""

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(object_xml, encoding="utf-8")

    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    brex_path = brex_dir / _filename_for("DEADEND")
    brex_path.write_text(brex_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.set_brex_path(str(brex_dir))

    checker._init_brex_list()

    assert checker._brex_list[0] == [str(brex_path)]
