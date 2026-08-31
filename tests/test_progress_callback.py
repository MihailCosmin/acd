import sys

from acd.brex_checker import BrexChecker

BREX_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/brex.xsd"

BREX_CONTENT = f"""<brex xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="{BREX_SCHEMA}">
<contextRules>
<structureObjectRuleGroup>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbiddenA</objectPath>
<objectUse>forbiddenA must not be present.</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbiddenB</objectPath>
<objectUse>forbiddenB must not be present.</objectUse>
</structureObjectRule>
<structureObjectRule>
<objectPath allowedObjectFlag="0">//forbiddenC</objectPath>
<objectUse>forbiddenC must not be present.</objectUse>
</structureObjectRule>
</structureObjectRuleGroup>
</contextRules>
</brex>
"""

XML_CONTENT = "<dml><forbiddenA/></dml>"


def test_brex_checker_does_not_import_tqdm():
    # Ref P3: "Add a progress callback rather than the hard tqdm dependency
    # inside the library." The module must not depend on tqdm at all any more.
    import acd.brex_checker as brex_checker_module
    assert not hasattr(brex_checker_module, "tqdm")


def _checker(tmp_path, brex_path):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def test_check_rules_reports_progress_per_content_rule(tmp_path):
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")
    checker = _checker(tmp_path, str(brex_path))

    calls = []
    checker._check_rules(progress_callback=lambda current, total, stage: calls.append((current, total, stage)))

    assert calls == [(1, 3, "rules"), (2, 3, "rules"), (3, 3, "rules")]


def test_check_rules_with_no_callback_does_not_raise(tmp_path):
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")
    checker = _checker(tmp_path, str(brex_path))

    result = checker._check_rules()

    assert len(result[str(brex_path)]['0']) == 1


def test_validate_reports_progress_per_file_in_directory_mode(tmp_path):
    brex_dir = tmp_path / "brex"
    brex_dir.mkdir()
    brex_path = brex_dir / "brex.xml"
    brex_path.write_text(BREX_CONTENT, encoding="utf-8")

    obj_dir = tmp_path / "objects"
    obj_dir.mkdir()
    (obj_dir / "a.xml").write_text(XML_CONTENT, encoding="utf-8")
    (obj_dir / "b.xml").write_text("<dml><forbiddenB/></dml>", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(obj_dir))
    checker.override_brex_list([str(brex_path)])

    file_calls = []
    rule_calls = []

    def on_progress(current, total, stage):
        if stage == "files":
            file_calls.append((current, total))
        else:
            rule_calls.append((current, total))

    results = checker.validate(progress_callback=on_progress)

    assert len(results) == 2
    assert file_calls == [(1, 2), (2, 2)]
    # One "rules" progression (length 3, matching the 3 rules in BREX_CONTENT)
    # per checked file.
    assert len(rule_calls) == 6
