import pytest
from lxml import etree

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

XML_CONTENT = (
    '<dml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}"/>\n'
)


def _validate(tmp_path, brex_content):
    brex_path = tmp_path / "brex.xml"
    brex_path.write_text(brex_content, encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(brex_path)])
    return checker, checker.validate()


def test_non_context_rule_with_br_decision_ref_is_surfaced(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BREX-S1-00245"/>
      <simplePara>Deletion of data modules is treated as a special case of update.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    _, result = _validate(tmp_path, brex_content)

    assert len(result["nonContextRules"]) == 1
    entry = result["nonContextRules"][0]
    assert entry["Text"] == "Deletion of data modules is treated as a special case of update."
    assert entry["BrDecisionIdentNumber"] == "BREX-S1-00245"


def test_non_context_rule_without_br_decision_ref_reports_none(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <simplePara>No decision reference on this one.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    _, result = _validate(tmp_path, brex_content)

    assert len(result["nonContextRules"]) == 1
    assert result["nonContextRules"][0]["BrDecisionIdentNumber"] is None


def test_non_context_rule_without_simple_para_falls_back_to_full_text(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>Bare text with no simplePara wrapper.</nonContextRule>
  </nonContextRules>
</brex>
"""
    _, result = _validate(tmp_path, brex_content)

    assert result["nonContextRules"][0]["Text"] == "Bare text with no simplePara wrapper."


def test_multiple_non_context_rules_all_surfaced_in_document_order(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BR-1"/>
      <simplePara>First rule.</simplePara>
    </nonContextRule>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BR-2"/>
      <simplePara>Second rule.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    _, result = _validate(tmp_path, brex_content)

    assert [e["Text"] for e in result["nonContextRules"]] == ["First rule.", "Second rule."]
    assert [e["BrDecisionIdentNumber"] for e in result["nonContextRules"]] == ["BR-1", "BR-2"]


def test_no_non_context_rules_defined_reports_empty_list(tmp_path):
    _, result = _validate(tmp_path, "<brex></brex>\n")

    assert result["nonContextRules"] == []


def test_non_context_rules_do_not_affect_error_count(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BREX-S1-00245"/>
      <simplePara>Purely informational, not a violation.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    _, result = _validate(tmp_path, brex_content)

    assert result["Summary"] == "0 Errors"


def test_non_context_rules_collected_across_layered_brex(tmp_path):
    lower_brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BR-LOWER"/>
      <simplePara>Rule from the lower-layer BREX.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    upper_brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BR-UPPER"/>
      <simplePara>Rule from the upper-layer BREX.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    lower_path = tmp_path / "lower_brex.xml"
    lower_path.write_text(lower_brex_content, encoding="utf-8")
    upper_path = tmp_path / "upper_brex.xml"
    upper_path.write_text(upper_brex_content, encoding="utf-8")

    xml_path = tmp_path / "object.xml"
    xml_path.write_text(XML_CONTENT, encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([str(upper_path), str(lower_path)])
    result = checker.validate()

    br_numbers = {e["BrDecisionIdentNumber"] for e in result["nonContextRules"]}
    assert br_numbers == {"BR-LOWER", "BR-UPPER"}


def test_xml_report_includes_non_context_rule_node(tmp_path):
    brex_content = """<brex>
  <nonContextRules>
    <nonContextRule>
      <brDecisionRef brDecisionIdentNumber="BREX-S1-00245"/>
      <simplePara>Deletion of data modules is treated as a special case of update.</simplePara>
    </nonContextRule>
  </nonContextRules>
</brex>
"""
    checker, result = _validate(tmp_path, brex_content)

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    rule_node = root.find(".//document/nonContextRules/nonContextRule")
    assert rule_node is not None
    assert rule_node.find("brDecisionRef").get("brDecisionIdentNumber") == "BREX-S1-00245"
    assert rule_node.findtext("text") == "Deletion of data modules is treated as a special case of update."


def test_xml_report_omits_non_context_rules_node_when_none_defined(tmp_path):
    checker, result = _validate(tmp_path, "<brex></brex>\n")

    root = etree.fromstring(checker.to_xml_report(result).encode("utf-8"))
    document = root.find(".//document")
    assert document.find("nonContextRules") is None
