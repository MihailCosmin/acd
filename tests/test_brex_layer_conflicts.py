from acd.brex_checker import BrexChecker

BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"


def _dm_code(model_ident_code: str) -> str:
    return (
        f'<dmCode modelIdentCode="{model_ident_code}" systemDiffCode="A" systemCode="00" '
        'subSystemCode="0" subSubSystemCode="0" assyCode="00" disassyCode="00" '
        'disassyCodeVariant="A" infoCode="022" infoCodeVariant="A" itemLocationCode="D"/>'
    )


def _filename_for(model_ident_code: str) -> str:
    return f"DMC-{model_ident_code}-A-00-00-00-00A-022A-D_001-00_EN-US.XML"


def _write_layered_brex(tmp_path):
    """A two-layer BREX chain: a project-specific BREX (layer 0, "lower" per
    S1000D valueTailoring terminology) that references a self-terminating
    master BREX (layer 1, "higher"). The two layers agree on `//safeElement`,
    disagree on `//foo`'s allowedObjectFlag, and disagree on the restrictable
    value set for `//code/@value` (the project layer adds "zz", which the
    master layer never declared).
    """
    base_code = _dm_code("BASE0001")

    project_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{base_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//foo</objectPath>
<objectUse>foo must not be present (project rule).</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="2">//code/@value</objectPath>
<objectUse>code/@value constrained (project rule).</objectUse>
<objectValue valueForm="single" valueAllowed="aa" valueTailoring="restrictable"/>
<objectValue valueForm="single" valueAllowed="bb" valueTailoring="restrictable"/>
<objectValue valueForm="single" valueAllowed="zz" valueTailoring="restrictable"/>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//safeElement</objectPath>
<objectUse>safeElement must not be present (agrees with master).</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//narrowedElement/@value</objectPath>
<objectUse>narrowedElement/@value forbidden outright by the project (a legitimate narrowing of the master's "2" rule below, not a conflict).</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""

    base_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<dmStatus>
<brexDmRef><dmRef><dmRefIdent>{base_code}</dmRefIdent></dmRef></brexDmRef>
</dmStatus>
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="1">//foo</objectPath>
<objectUse>foo must be present (master rule) -- conflicts with the project layer.</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="2">//code/@value</objectPath>
<objectUse>code/@value constrained (master rule).</objectUse>
<objectValue valueForm="single" valueAllowed="aa" valueTailoring="restrictable"/>
<objectValue valueForm="single" valueAllowed="bb" valueTailoring="restrictable"/>
<objectValue valueForm="single" valueAllowed="cc" valueTailoring="restrictable"/>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="2">//narrowedElement/@value</objectPath>
<objectUse>narrowedElement/@value constrained, if present (master rule).</objectUse>
<objectValue valueForm="single" valueAllowed="aa"/>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//safeElement</objectPath>
<objectUse>safeElement must not be present (agrees with project).</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""

    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    project_path = brex_dir / _filename_for("PROJECT1")
    base_path = brex_dir / _filename_for("BASE0001")
    project_path.write_text(project_content, encoding="utf-8")
    base_path.write_text(base_content, encoding="utf-8")
    return str(project_path), str(base_path)


def _categories(findings):
    return [f["Category"] for f in findings]


def test_conflicting_allowed_object_flag_detected_across_layers(tmp_path):
    project_path, base_path = _write_layered_brex(tmp_path)

    checker = BrexChecker()
    findings = checker.lint_brex_layers(project_path)

    conflicts = [f for f in findings if f["Category"] == "ConflictingAllowedObjectFlag"]
    assert len(conflicts) == 1
    assert conflicts[0]["Xpath"] == "//foo"
    flags_by_brex = {layer["Brex"]: layer["ObjectFlag"] for layer in conflicts[0]["Layers"]}
    assert flags_by_brex[project_path] == "0"
    assert flags_by_brex[base_path] == "1"


def test_restrictable_value_set_widening_detected_across_layers(tmp_path):
    project_path, base_path = _write_layered_brex(tmp_path)

    checker = BrexChecker()
    findings = checker.lint_brex_layers(project_path)

    widened = [f for f in findings if f["Category"] == "RestrictableValueSetWidened"]
    assert len(widened) == 1
    assert widened[0]["Xpath"] == "//code/@value"
    assert widened[0]["LowerBrex"] == project_path
    assert widened[0]["HigherBrex"] == base_path
    assert widened[0]["ExtraValues"] == ["zz"]


def test_forbidden_in_one_layer_and_value_constrained_in_another_not_flagged(tmp_path):
    # A project layer forbidding an attribute outright ("0") while a more
    # general layer only value-constrains it ("2") is not a contradiction:
    # an object that omits the attribute entirely satisfies both rules at
    # once (the "2" rule has nothing to check when the attribute is absent).
    # Verified against the real ATABREX 01A/00A/S1000D-default chain, where
    # this exact pattern occurs twice and is not an authoring mistake.
    project_path, _base_path = _write_layered_brex(tmp_path)

    checker = BrexChecker()
    findings = checker.lint_brex_layers(project_path)

    for finding in findings:
        assert finding.get("Xpath") != "//narrowedElement/@value"


def test_agreeing_rule_across_layers_not_flagged(tmp_path):
    project_path, _base_path = _write_layered_brex(tmp_path)

    checker = BrexChecker()
    findings = checker.lint_brex_layers(project_path)

    for finding in findings:
        assert finding.get("Xpath") != "//safeElement"


def test_single_layer_brex_produces_no_findings(tmp_path):
    brex_content = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//foo</objectPath>
<objectUse>foo must not be present.</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")

    checker = BrexChecker()
    assert checker.lint_brex_layers(str(brex_path)) == []
