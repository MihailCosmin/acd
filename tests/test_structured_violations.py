import json

from acd.brex_checker import BrexChecker
from acd.brex_checker import BrexViolation

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"
BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

BASE_DM_CODE = (
    '<dmCode modelIdentCode="BASE0001" systemDiffCode="A" systemCode="00" '
    'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
    'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
)
PROJECT_DM_CODE = (
    '<dmCode modelIdentCode="PROJECT1" systemDiffCode="A" systemCode="00" '
    'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
    'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
)


def _write_layered_brex(tmp_path):
    """A two-layer BREX chain (project -> self-terminating master), same
    shape as test_brex_layer_conflicts.py's fixture. Both layers declare the
    identical `//dupElement` forbidden rule (inherited/re-stated unchanged,
    a common real-world pattern -- see `lint_brex_layers`), so checking an
    object against both layers should report it once, not twice. Each layer
    also carries a rule the other does not, to confirm those are unaffected.
    """
    project_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{BASE_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule id="dup-rule">
<objectPath allowedObjectFlag="0">//dupElement</objectPath>
<objectUse>dupElement must not be present (declared identically in both layers).</objectUse>
<brDecisionRef brDecisionIdentNumber="BR-DUP"/>
</structureObjectRule>
<structureObjectRule id="project-only-rule">
<objectPath allowedObjectFlag="0">//projectOnlyElement</objectPath>
<objectUse>projectOnlyElement must not be present (project layer only).</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""
    base_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{BASE_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule id="dup-rule">
<objectPath allowedObjectFlag="0">//dupElement</objectPath>
<objectUse>dupElement must not be present (declared identically in both layers).</objectUse>
<brDecisionRef brDecisionIdentNumber="BR-DUP"/>
</structureObjectRule>
<structureObjectRule id="base-only-rule">
<objectPath allowedObjectFlag="0">//baseOnlyElement</objectPath>
<objectUse>baseOnlyElement must not be present (master layer only).</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""
    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    project_path = brex_dir / "DMC-PROJECT1-A-00-00-00-00A-022A-D_001-00_EN-US.XML"
    base_path = brex_dir / "DMC-BASE0001-A-00-00-00-00A-022A-D_001-00_EN-US.XML"
    project_path.write_text(project_content, encoding="utf-8")
    base_path.write_text(base_content, encoding="utf-8")
    return str(project_path), str(base_path)


def _write_object(tmp_path, extra=""):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(f"""<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">
<content>
<dupElement/>
{extra}
</content>
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{PROJECT_DM_CODE}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
</dmodule>
""", encoding="utf-8")
    return str(xml_path)


def _checked_result(tmp_path):
    project_path, base_path = _write_layered_brex(tmp_path)
    xml_path = _write_object(tmp_path, extra="<projectOnlyElement/>")

    checker = BrexChecker()
    checker.set_xml(xml_path)
    checker.set_brex_path(str(tmp_path / "brex"))
    result = checker.validate()
    return checker, result, project_path, base_path


def test_raw_result_keeps_one_entry_per_layer_but_marks_the_later_one_duplicate(tmp_path):
    checker, result, project_path, base_path = _checked_result(tmp_path)

    # Both layers still carry their own record -- `_check_rules`'s per-BREX
    # keys are unchanged and nothing is removed from either list.
    assert len(result[project_path]['0']) == 2  # dupElement + projectOnlyElement
    assert len(result[base_path]['0']) == 1  # dupElement only (baseOnlyElement absent from the object)

    project_dup = next(v for v in result[project_path]['0'] if v['Xpath'] == '//dupElement')
    base_dup = next(v for v in result[base_path]['0'] if v['Xpath'] == '//dupElement')

    # Project is layer 0 (nearest/most specific), walked first -- its copy is
    # canonical; the master layer's identical copy is the duplicate.
    assert project_dup['Duplicate'] is False
    assert base_dup['Duplicate'] is True

    project_only = next(v for v in result[project_path]['0'] if v['Xpath'] == '//projectOnlyElement')
    assert project_only['Duplicate'] is False


def test_summary_counts_the_duplicate_only_once(tmp_path):
    checker, result, project_path, base_path = _checked_result(tmp_path)

    # 2 real violations: dupElement (reported once, not twice) and
    # projectOnlyElement. baseOnlyElement never fires (absent from the
    # object).
    assert result["Summary"] == "2 Errors"


def test_run_summary_excludes_duplicates_too(tmp_path):
    checker, result, project_path, base_path = _checked_result(tmp_path)

    totals = checker.run_summary(result)
    assert totals["Errors"] == 2


def test_violations_returns_dataclass_instances_with_expected_fields(tmp_path):
    checker, result, project_path, base_path = _checked_result(tmp_path)

    records = checker.violations(result)
    assert all(isinstance(r, BrexViolation) for r in records)

    dup_records = [r for r in records if r.object_path == '//dupElement']
    assert len(dup_records) == 2  # both layers still represented in the full list
    assert sorted(r.duplicate for r in dup_records) == [False, True]

    canonical = next(r for r in dup_records if not r.duplicate)
    assert canonical.brex == project_path
    assert canonical.flag == '0'
    assert canonical.rule_id == 'dup-rule'
    assert canonical.br_decision_ident_number == 'BR-DUP'
    assert canonical.object_use == 'dupElement must not be present (declared identically in both layers).'
    assert canonical.node_xpath is not None
    assert canonical.line != 'x'
    assert canonical.fail is True


def test_to_json_report_excludes_duplicates_and_matches_run_summary(tmp_path):
    checker, result, project_path, base_path = _checked_result(tmp_path)

    report = json.loads(checker.to_json_report(result))

    assert report["summary"]["Errors"] == 2
    assert len(report["violations"]) == 2
    object_paths = {v["object_path"] for v in report["violations"]}
    assert object_paths == {'//dupElement', '//projectOnlyElement'}
    assert all(not v["duplicate"] for v in report["violations"])


def test_to_xml_report_emits_one_error_for_the_deduplicated_violation(tmp_path):
    from lxml import etree

    checker, result, project_path, base_path = _checked_result(tmp_path)

    xml_report = checker.to_xml_report(result)
    root = etree.fromstring(xml_report.encode("utf-8"))

    dup_errors = root.xpath('//error[objectPath="//dupElement"]')
    assert len(dup_errors) == 1

    all_errors = root.xpath('//error')
    assert len(all_errors) == 2  # dupElement (once) + projectOnlyElement


def test_single_brex_duplicate_rules_are_also_deduplicated(tmp_path):
    # The dedup identity key deliberately does not include the BREX path:
    # two literal duplicate rule definitions within one file are the same
    # real-world defect too.
    brex_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbidden</objectPath>
<objectUse>forbidden must not be present (first copy).</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbidden</objectPath>
<objectUse>forbidden must not be present (second, accidental duplicate copy).</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text("<dml><forbidden/></dml>", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(brex_path)])
    result = checker._check_rules()

    entries = result[str(brex_path)]['0']
    assert len(entries) == 2
    assert sorted(e['Duplicate'] for e in entries) == [False, True]
