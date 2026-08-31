from acd.brex_checker import BrexChecker


def _lint(tmp_path, brex_content):
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    return checker, checker.lint_brex(str(brex_path))


def _categories(findings):
    return [f["Category"] for f in findings]


def test_well_formed_brex_has_no_findings(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must match the allowed pattern or range.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[A-Z]{2}[0-9]{2}"/>
        <objectValue valueForm="range" valueAllowed="aa01~aa09"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert findings == []


def test_invalid_xpath_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//foo[@bar=</objectPath>
        <objectUse>malformed xpath.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidXPath" in _categories(findings)
    finding = next(f for f in findings if f["Category"] == "InvalidXPath")
    assert finding["Xpath"] == "//foo[@bar="


def test_missing_object_use_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "MissingObjectUse" in _categories(findings)


def test_flag_2_without_object_value_reported_as_no_op(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must match a constraint, but none is declared.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "EmptyValueFlag2" in _categories(findings)


def test_flag_1_without_object_value_not_flagged(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present, no value constraint.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "EmptyValueFlag2" not in _categories(findings)


def test_invalid_pattern_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must match the pattern.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[a-z"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidPattern" in _categories(findings)


def test_pattern_with_no_value_allowed_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must match the pattern.</objectUse>
        <objectValue valueForm="pattern"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidPattern" in _categories(findings)


def test_invalid_range_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must fall in the range.</objectUse>
        <objectValue valueForm="range" valueAllowed="aa01~"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidRange" in _categories(findings)


def test_invalid_set_member_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must be in the set.</objectUse>
        <objectValue valueForm="range" valueAllowed="A||B"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidRange" in _categories(findings)


def test_valid_set_and_range_not_flagged(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must be in the set or range.</objectUse>
        <objectValue valueForm="range" valueAllowed="A|B|C"/>
        <objectValue valueForm="range" valueAllowed="aa01~aa09"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidRange" not in _categories(findings)


def test_unknown_severity_level_reported(tmp_path):
    severity_path = tmp_path / ".brseveritylevels"
    severity_path.write_text(
        '<brSeverityLevels>\n'
        '  <brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>\n'
        '</brSeverityLevels>\n',
        encoding="utf-8",
    )
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule brSeverityLevel="brslUnknown">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    checker.set_severity_levels_path(str(severity_path))
    findings = checker.lint_brex(str(brex_path))
    assert "UnknownSeverityLevel" in _categories(findings)
    finding = next(f for f in findings if f["Category"] == "UnknownSeverityLevel")
    assert finding["BrSeverityLevel"] == "brslUnknown"


def test_known_severity_level_not_flagged(tmp_path):
    severity_path = tmp_path / ".brseveritylevels"
    severity_path.write_text(
        '<brSeverityLevels>\n'
        '  <brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>\n'
        '</brSeverityLevels>\n',
        encoding="utf-8",
    )
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule brSeverityLevel="brsl01">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    checker.set_severity_levels_path(str(severity_path))
    findings = checker.lint_brex(str(brex_path))
    assert "UnknownSeverityLevel" not in _categories(findings)


def test_no_severity_file_configured_skips_severity_check(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule brSeverityLevel="brslUnknown">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "UnknownSeverityLevel" not in _categories(findings)


def test_duplicate_id_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    duplicate_findings = [f for f in findings if f["Category"] == "DuplicateId"]
    assert len(duplicate_findings) == 1
    assert duplicate_findings[0]["Id"] == "SOR-1"
    assert len(duplicate_findings[0]["Lines"]) == 2


def test_distinct_ids_not_flagged(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="SOR-1">
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="SOR-2">
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "DuplicateId" not in _categories(findings)


def test_duplicate_br_decision_ident_number_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
        <brDecisionRef brDecisionIdentNumber="SOR-230"/>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
        <brDecisionRef brDecisionIdentNumber="SOR-230"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    duplicate_findings = [f for f in findings if f["Category"] == "DuplicateBrDecisionIdentNumber"]
    assert len(duplicate_findings) == 1
    assert duplicate_findings[0]["BrDecisionIdentNumber"] == "SOR-230"


def test_unreachable_rules_context_reported_when_schema_absent_from_csdb(tmp_path):
    brex_content = """<brex>
  <contextRules rulesContext="http://example.com/other-schema.xsd">
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    findings = checker.lint_brex(str(brex_path), csdb_schemas={"http://example.com/used-schema.xsd"})
    assert "UnreachableRulesContext" in _categories(findings)
    finding = next(f for f in findings if f["Category"] == "UnreachableRulesContext")
    assert finding["RulesContext"] == "http://example.com/other-schema.xsd"


def test_unreachable_rules_context_not_flagged_when_schema_used_in_csdb(tmp_path):
    brex_content = """<brex>
  <contextRules rulesContext="http://example.com/used-schema.xsd">
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    findings = checker.lint_brex(str(brex_path), csdb_schemas={"http://example.com/used-schema.xsd"})
    assert "UnreachableRulesContext" not in _categories(findings)


def test_unreachable_rules_context_skipped_without_csdb_schemas(tmp_path):
    brex_content = """<brex>
  <contextRules rulesContext="http://example.com/other-schema.xsd">
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "UnreachableRulesContext" not in _categories(findings)


def test_unreachable_rules_context_skipped_when_csdb_schemas_empty(tmp_path):
    brex_content = """<brex>
  <contextRules rulesContext="http://example.com/other-schema.xsd">
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    findings = checker.lint_brex(str(brex_path), csdb_schemas=set())
    assert "UnreachableRulesContext" not in _categories(findings)


def test_unqualified_rules_context_never_flagged_as_unreachable(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    checker = BrexChecker()
    findings = checker.lint_brex(str(brex_path), csdb_schemas={"http://example.com/anything.xsd"})
    assert "UnreachableRulesContext" not in _categories(findings)


def test_duplicate_sns_code_reported(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
        </snsSystem>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Duplicate air conditioning</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    duplicate_findings = [f for f in findings if f["Category"] == "DuplicateSnsCode"]
    assert len(duplicate_findings) == 1
    assert duplicate_findings[0]["Level"] == "snsSystem"
    assert duplicate_findings[0]["SnsCode"] == "21"
    assert len(duplicate_findings[0]["Lines"]) == 2


def test_distinct_sns_codes_not_flagged_as_duplicate(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
        </snsSystem>
        <snsSystem>
            <snsCode>22</snsCode>
            <snsTitle>Auto flight</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "DuplicateSnsCode" not in _categories(findings)


def test_missing_sns_title_reported(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "MissingSnsTitle" in _categories(findings)
    finding = next(f for f in findings if f["Category"] == "MissingSnsTitle")
    assert finding["Level"] == "snsSystem"
    assert finding["SnsCode"] == "21"


def test_sns_title_present_not_flagged(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "MissingSnsTitle" not in _categories(findings)


def test_sns_code_outside_declared_pattern_reported(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//dmCode/@systemCode</objectPath>
        <objectUse>systemCode must be two digits.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[0-9]{2}"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
        </snsSystem>
        <snsSystem>
            <snsCode>ZZ</snsCode>
            <snsTitle>Not a valid systemCode</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    outside_findings = [f for f in findings if f["Category"] == "SnsCodeOutsidePattern"]
    assert len(outside_findings) == 1
    assert outside_findings[0]["Level"] == "snsSystem"
    assert outside_findings[0]["SnsCode"] == "ZZ"


def test_sns_code_matching_declared_pattern_not_flagged(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//dmCode/@systemCode</objectPath>
        <objectUse>systemCode must be two digits.</objectUse>
        <objectValue valueForm="pattern" valueAllowed="[0-9]{2}"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>21</snsCode>
            <snsTitle>Air conditioning</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "SnsCodeOutsidePattern" not in _categories(findings)


def test_sns_code_not_checked_against_pattern_when_none_declared(tmp_path):
    brex_content = """<brex>
  <snsRules>
    <snsDescr>
        <snsSystem>
            <snsCode>ZZ</snsCode>
            <snsTitle>No pattern rule declared for systemCode.</snsTitle>
        </snsSystem>
    </snsDescr>
  </snsRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "SnsCodeOutsidePattern" not in _categories(findings)


def test_no_sns_rules_produces_no_sns_findings(tmp_path):
    brex_content = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    sns_categories = {"DuplicateSnsCode", "MissingSnsTitle", "SnsCodeOutsidePattern"}
    assert sns_categories.isdisjoint(_categories(findings))


def test_legacy_brex_spelling_is_linted(tmp_path):
    brex_content = """<brex>
  <contextrules>
    <structrules>
      <objrule>
        <objpath>//forbiddenElement</objpath>
        <objuse>forbiddenElement must not be present.</objuse>
      </objrule>
      <objrule>
        <objpath>//code/@value</objpath>
        <objuse>code/@value must match a constraint.</objuse>
        <objval valtype="pattern" val1="[a-z"/>
      </objrule>
    </structrules>
  </contextrules>
</brex>
"""
    _, findings = _lint(tmp_path, brex_content)
    assert "InvalidPattern" in _categories(findings)
