import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

SEVERITY_LEVELS_CONTENT = """<brSeverityLevels>
  <brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>
  <brSeverityLevel value="brsl02" fail="no">Warning</brSeverityLevel>
</brSeverityLevels>
"""


def make_xml(extra: str = "") -> str:
    return (
        '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        f"{extra}\n"
        "</dml>\n"
    )


def make_brex(default_severity: str = "") -> str:
    default_attr = f' defaultBrSeverityLevel="{default_severity}"' if default_severity else ""
    return f"""<brex{default_attr}>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule brSeverityLevel="brsl02">
        <objectPath allowedObjectFlag="0">//warnOnly</objectPath>
        <objectUse>warnOnly should not be present, but only warns.</objectUse>
      </structureObjectRule>
      <structureObjectRule brSeverityLevel="brsl01">
        <objectPath allowedObjectFlag="0">//failHard</objectPath>
        <objectUse>failHard must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//noLevel</objectPath>
        <objectUse>noLevel must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule brSeverityLevel="brslUnknown">
        <objectPath allowedObjectFlag="0">//unknownLevel</objectPath>
        <objectUse>unknownLevel must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


@pytest.fixture
def severity_levels_path(tmp_path):
    path = tmp_path / ".brseveritylevels"
    path.write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(make_brex(), encoding="utf-8")
    return str(path)


@pytest.fixture
def brex_path_with_default(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(make_brex(default_severity="brsl02"), encoding="utf-8")
    return str(path)


def _entry_for(entries, xpath):
    matches = [e for e in entries if e['Xpath'] == xpath]
    assert len(matches) == 1, f"expected exactly one entry for {xpath!r}, found {len(matches)}"
    return matches[0]


def _checker(tmp_path, brex_path):
    return _checker_at(tmp_path, brex_path)


def _checker_at(xml_dir, brex_path):
    xml_content = make_xml("<warnOnly/>\n<failHard/>\n<noLevel/>\n<unknownLevel/>")
    xml_path = xml_dir / "object.xml"
    xml_path.write_text(xml_content, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker


def test_fail_no_severity_level_reports_as_warning_when_levels_file_set(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//warnOnly')
    assert entry['BrSeverityLevel'] == "brsl02"
    assert entry['Fail'] is False


def test_fail_yes_severity_level_still_reports_as_error(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//failHard')
    assert entry['BrSeverityLevel'] == "brsl01"
    assert entry['Fail'] is True


def test_missing_severity_level_defaults_to_failure(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//noLevel')
    assert entry['BrSeverityLevel'] is None
    assert entry['Fail'] is True


def test_unresolved_severity_level_defaults_to_failure(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//unknownLevel')
    assert entry['BrSeverityLevel'] == "brslUnknown"
    assert entry['Fail'] is True


def test_without_severity_levels_file_every_violation_fails(tmp_path, brex_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_search(False)

    result = checker._check_rules()[brex_path]

    for xpath in ('//warnOnly', '//failHard', '//noLevel', '//unknownLevel'):
        entry = _entry_for(result['0'], xpath)
        assert entry['Fail'] is True


def test_default_br_severity_level_falls_back_when_rule_has_none(tmp_path, brex_path_with_default, severity_levels_path):
    checker = _checker(tmp_path, brex_path_with_default)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker._check_rules()[brex_path_with_default]

    entry = _entry_for(result['0'], '//noLevel')
    assert entry['BrSeverityLevel'] == "brsl02"
    assert entry['Fail'] is False


def test_summary_separates_warnings_from_errors(tmp_path, brex_path, severity_levels_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(severity_levels_path)

    result = checker.validate()

    assert result["Summary"] == "3 Errors, 1 Warnings"


def test_summary_counts_everything_as_errors_without_severity_levels_file(tmp_path, brex_path):
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_search(False)

    result = checker.validate()

    assert result["Summary"] == "4 Errors"


def test_auto_discovers_severity_levels_in_same_directory(tmp_path, brex_path):
    (tmp_path / ".brseveritylevels").write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    checker = _checker(tmp_path, brex_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//warnOnly')
    assert entry['Fail'] is False


def test_auto_discovers_severity_levels_in_ancestor_directory(tmp_path, brex_path):
    (tmp_path / ".brseveritylevels").write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    nested_dir = tmp_path / "csdb" / "dmodule"
    nested_dir.mkdir(parents=True)
    checker = _checker_at(nested_dir, brex_path)

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//warnOnly')
    assert entry['Fail'] is False


def test_explicit_severity_levels_path_overrides_auto_discovered_file(tmp_path, brex_path):
    (tmp_path / ".brseveritylevels").write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    override_content = SEVERITY_LEVELS_CONTENT.replace(
        'value="brsl02" fail="no"', 'value="brsl02" fail="yes"'
    )
    override_path = tmp_path / "override.brseveritylevels"
    override_path.write_text(override_content, encoding="utf-8")

    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_path(str(override_path))

    result = checker._check_rules()[brex_path]

    entry = _entry_for(result['0'], '//warnOnly')
    assert entry['BrSeverityLevel'] == "brsl02"
    assert entry['Fail'] is True


def test_disabling_search_ignores_discoverable_file(tmp_path, brex_path):
    (tmp_path / ".brseveritylevels").write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    checker = _checker(tmp_path, brex_path)
    checker.set_severity_levels_search(False)

    result = checker._check_rules()[brex_path]

    for xpath in ('//warnOnly', '//failHard', '//noLevel', '//unknownLevel'):
        entry = _entry_for(result['0'], xpath)
        assert entry['Fail'] is True
