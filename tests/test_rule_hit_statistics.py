from acd.brex_checker import BrexChecker

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must never be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//presentElement</objectPath>
        <objectUse>presentElement must not be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="1">//requiredElement</objectPath>
        <objectUse>requiredElement must be present.</objectUse>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//code/@value</objectPath>
        <objectUse>code/@value must be "aa".</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
      </structureObjectRule>
      <structureObjectRule>
        <objectPath allowedObjectFlag="2">//neverUsed/@value</objectPath>
        <objectUse>declared as value-constrained, but no objectValue given.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


def _write(tmp_path, name, content):
    path = tmp_path / name
    path.write_text(content, encoding="utf-8")
    return str(path)


def _doc(has_present, has_required, code_value):
    body = ""
    if has_present:
        body += "<presentElement/>"
    if has_required:
        body += "<requiredElement/>"
    body += f'<code value="{code_value}"/>'
    return f"<container>{body}</container>\n"


def _stats_by_xpath_and_flag(checker):
    return {(s["Xpath"], s["ObjectFlag"]): s for s in checker.rule_hit_statistics()}


def test_rule_hit_statistics_accumulate_across_documents(tmp_path):
    brex_path = _write(tmp_path, "brex.xml", BREX_CONTENT)
    doc1 = _write(tmp_path, "doc1.xml", _doc(has_present=True, has_required=True, code_value="aa"))
    doc2 = _write(tmp_path, "doc2.xml", _doc(has_present=False, has_required=False, code_value="bb"))

    checker = BrexChecker()
    checker.override_brex_list([brex_path])

    checker.set_xml(doc1)
    checker.validate()
    checker.set_xml(doc2)
    checker.validate()

    stats = _stats_by_xpath_and_flag(checker)

    forbidden = stats[("//forbiddenElement", "0")]
    assert forbidden["Evaluated"] == 2
    assert forbidden["Matched"] == 0
    assert forbidden["Violated"] == 0

    present = stats[("//presentElement", "0")]
    assert present["Evaluated"] == 2
    assert present["Matched"] == 1
    assert present["Violated"] == 1

    required = stats[("//requiredElement", "1")]
    assert required["Evaluated"] == 2
    assert required["Matched"] == 1
    assert required["Violated"] == 1

    code = stats[("//code/@value", "2")]
    assert code["Evaluated"] == 2
    assert code["Matched"] == 2
    assert code["Violated"] == 1

    never_used = stats[("//neverUsed/@value", "2")]
    assert never_used["Evaluated"] == 2
    assert never_used["Matched"] == 0
    assert never_used["Violated"] == 0


def test_rule_hit_statistics_carry_identifying_fields(tmp_path):
    brex_path = _write(tmp_path, "brex.xml", BREX_CONTENT)
    doc_path = _write(tmp_path, "doc.xml", _doc(True, True, "aa"))

    checker = BrexChecker()
    checker.override_brex_list([brex_path])
    checker.set_xml(doc_path)
    checker.validate()

    stats = _stats_by_xpath_and_flag(checker)
    required = stats[("//requiredElement", "1")]
    assert required["Brex"] == brex_path
    assert required["ObjectUse"] == "requiredElement must be present."
    assert required["ContextRules"] == ""


def test_reset_rule_statistics_clears_accumulated_stats(tmp_path):
    brex_path = _write(tmp_path, "brex.xml", BREX_CONTENT)
    doc_path = _write(tmp_path, "doc.xml", _doc(True, True, "aa"))

    checker = BrexChecker()
    checker.override_brex_list([brex_path])
    checker.set_xml(doc_path)
    checker.validate()

    assert checker.rule_hit_statistics() != []
    checker.reset_rule_statistics()
    assert checker.rule_hit_statistics() == []


def test_rule_hit_statistics_empty_before_any_check(tmp_path):
    checker = BrexChecker()
    assert checker.rule_hit_statistics() == []
