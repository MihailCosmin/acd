"""`XmlChecker` -- the encoding/syntax/structure/DTD/schema checker that runs
before a BREX check can mean anything.

Covers each of the five layers in turn, the order they short-circuit in (a
file that does not parse is not schema-checked), schema resolution and
caching, and the three report formats.

Every test builds its own objects and its own schema on disk: nothing here
touches the network, so the suite behaves the same on a build agent as on a
workstation.
"""

from html.parser import HTMLParser

import pytest

from acd.xml_checker import CHECK_RULES
from acd.xml_checker import XmlChecker

SCHEMA = """<?xml version="1.0" encoding="UTF-8"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <xsd:element name="dmodule">
    <xsd:complexType>
      <xsd:sequence>
        <xsd:element name="identAndStatusSection" minOccurs="1">
          <xsd:complexType>
            <xsd:sequence>
              <xsd:element name="dmAddress">
                <xsd:complexType>
                  <xsd:sequence>
                    <xsd:element name="dmIdent">
                      <xsd:complexType>
                        <xsd:sequence>
                          <xsd:element name="dmCode">
                            <xsd:complexType>
                              <xsd:attribute name="modelIdentCode" type="xsd:string"/>
                              <xsd:attribute name="systemDiffCode" type="xsd:string"/>
                              <xsd:attribute name="systemCode" type="xsd:string"/>
                              <xsd:attribute name="subSystemCode" type="xsd:string"/>
                              <xsd:attribute name="subSubSystemCode" type="xsd:string"/>
                              <xsd:attribute name="assyCode" type="xsd:string"/>
                              <xsd:attribute name="disassyCode" type="xsd:string"/>
                              <xsd:attribute name="disassyCodeVariant" type="xsd:string"/>
                              <xsd:attribute name="infoCode" type="xsd:string"/>
                              <xsd:attribute name="infoCodeVariant" type="xsd:string"/>
                              <xsd:attribute name="itemLocationCode" type="xsd:string"/>
                            </xsd:complexType>
                          </xsd:element>
                        </xsd:sequence>
                      </xsd:complexType>
                    </xsd:element>
                  </xsd:sequence>
                </xsd:complexType>
              </xsd:element>
            </xsd:sequence>
          </xsd:complexType>
        </xsd:element>
        <xsd:element name="content" minOccurs="1"/>
      </xsd:sequence>
    </xsd:complexType>
  </xsd:element>
</xsd:schema>
"""

# A filename whose code matches DM_IDENT below, so a test that is not about
# the ident check does not trip over it.
DM_NAME = "DMC-H160-B-67-34-0200-00A-040A-D_001-01_SX-US.xml"

DM_IDENT = (
    '<dmCode modelIdentCode="H160" systemDiffCode="B" systemCode="67" '
    'subSystemCode="3" subSubSystemCode="4" assyCode="0200" disassyCode="00" '
    'disassyCodeVariant="A" infoCode="040" infoCodeVariant="A" '
    'itemLocationCode="D"/>'
)


def make_dm(schema: str = None, ident: str = DM_IDENT, body: str = "<content/>",
            doctype: str = "", declaration: str = '<?xml version="1.0" encoding="UTF-8"?>') -> str:
    """A minimal but structurally real data module."""
    schema_attribute = (
        f' xsi:noNamespaceSchemaLocation="{schema}"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
        if schema else ""
    )
    return (
        f"{declaration}\n{doctype}"
        f"<dmodule{schema_attribute}>\n"
        f"  <identAndStatusSection><dmAddress><dmIdent>{ident}"
        "</dmIdent></dmAddress></identAndStatusSection>\n"
        f"  {body}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def schema_path(tmp_path):
    """A local schema on disk, so no test needs the network."""
    path = tmp_path / "schemas" / "dmodule.xsd"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(SCHEMA, encoding="utf-8")
    return path


def _check(path, **kwargs):
    """Check one object with the network off, returning `(checker, result)`."""
    checker = XmlChecker()
    checker.set_allow_network(False)
    for search_path in kwargs.pop("search_paths", []):
        checker.add_schema_search_path(str(search_path))
    checker.set_xml(str(path))
    return checker, checker.validate(**kwargs)


def _rules(checker, result):
    """The rule code of every finding, in report order."""
    return [row["Rule"] for row in checker.findings(result)]


# ---------------------------------------------------------------------------
# Layer 1 -- encoding
# ---------------------------------------------------------------------------

def test_empty_file_is_reported_and_stops_every_later_layer(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(b"   \n\t ")

    checker, result = _check(path)

    # Only ENC-EMPTY: parsing whitespace would produce a syntax error that
    # says nothing the empty-file finding has not already said.
    assert _rules(checker, result) == ["ENC-EMPTY"]
    assert result["Summary"]["Status"] == "Failed"


def test_utf8_bom_is_a_warning_not_an_error(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(b"\xef\xbb\xbf" + make_dm().encode("utf-8"))

    checker, result = _check(path)
    findings = [row for row in checker.findings(result) if row["Rule"] == "ENC-BOM"]

    # A BOM parses, so it must not fail the document -- but it breaks tools
    # that concatenate files, so it is still worth saying.
    assert len(findings) == 1
    assert findings[0]["Status"] == "Warning"
    assert result["Summary"]["Errors"] == 0


def test_content_that_does_not_match_its_declared_encoding_is_reported(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(
        b'<?xml version="1.0" encoding="UTF-8"?>\n<dmodule>\xff\xfe</dmodule>\n'
    )

    checker, result = _check(path)

    assert "ENC-DECL" in _rules(checker, result)


def test_illegal_control_character_is_reported_with_its_line(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(
        b'<?xml version="1.0" encoding="UTF-8"?>\n<dmodule>\n<content>\x03</content>\n</dmodule>\n'
    )

    checker, result = _check(path)
    finding = next(row for row in checker.findings(result) if row["Rule"] == "ENC-CTRL")

    assert finding["Line"] == 3
    assert "0x03" in finding["Detail"]


# ---------------------------------------------------------------------------
# Layer 2 -- syntax
# ---------------------------------------------------------------------------

def test_every_syntax_error_is_reported_not_only_the_first(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(
        '<?xml version="1.0"?>\n<dmodule>\n  <a></b>\n  <c></d>\n</dmodule>\n',
        encoding="utf-8",
    )

    checker, result = _check(path)
    syntax = [row for row in checker.findings(result) if row["Check"] == "syntax"]

    # The whole point of reporting the parser's error log rather than
    # error_log[0]: one run shows every problem, not one per re-run.
    assert len(syntax) > 1
    assert all(row["Line"] for row in syntax)


def test_a_file_that_does_not_parse_is_not_schema_checked(tmp_path, schema_path):
    path = tmp_path / DM_NAME
    path.write_text('<dmodule><unclosed></dmodule>', encoding="utf-8")

    checker, result = _check(path, search_paths=[schema_path.parent])
    checks = {row["Check"] for row in checker.findings(result)}

    # Schema errors from an unparseable document are artefacts of the syntax
    # error, so the layer is skipped rather than guessed at.
    assert checks == {"syntax"}


def test_undeclared_entity_is_a_syntax_error(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(body="<content>&nosuch;</content>"), encoding="utf-8")

    checker, result = _check(path)

    assert "XML-SYNTAX" in _rules(checker, result)


# ---------------------------------------------------------------------------
# Layer 3 -- structure
# ---------------------------------------------------------------------------

def test_root_element_is_checked_against_the_filename_type(tmp_path):
    path = tmp_path / "PMC-H160-D9893-06771-01_000-01_SX-US.xml"
    path.write_text(make_dm(), encoding="utf-8")

    checker, result = _check(path)
    finding = next(row for row in checker.findings(result) if row["Rule"] == "STR-ROOT")

    assert finding["Status"] == "Warning"
    assert "<pm>" in finding["Detail"] and "<dmodule>" in finding["Detail"]


def test_object_declaring_neither_schema_nor_doctype_is_reported(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(), encoding="utf-8")

    checker, result = _check(path)

    assert "STR-NOSCHEMA" in _rules(checker, result)


def test_filename_code_is_compared_with_the_code_inside_the_file(tmp_path):
    # Filename says itemLocationCode D; the ident says B.
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(ident=DM_IDENT.replace('itemLocationCode="D"', 'itemLocationCode="B"')),
        encoding="utf-8",
    )

    checker, result = _check(path)
    finding = next(row for row in checker.findings(result) if row["Rule"] == "STR-IDENT")

    assert finding["Status"] == "Error"
    assert "itemLocationCode" in finding["Detail"]
    assert "'D'" in finding["Detail"] and "'B'" in finding["Detail"]


def test_a_matching_filename_and_ident_raise_nothing(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(), encoding="utf-8")

    checker, result = _check(path)

    assert "STR-IDENT" not in _rules(checker, result)


def test_the_ident_is_read_from_the_ident_section_not_from_a_reference(tmp_path):
    # A brexDmRef carries a dmCode too, and it is a *different* object's code.
    # Reading the wrong one would report every object in the CSDB as mismatched.
    body = (
        "<content><refs><dmRef><dmRefIdent>"
        '<dmCode modelIdentCode="OTHER" systemDiffCode="Z" systemCode="99" '
        'subSystemCode="9" subSubSystemCode="9" assyCode="9999" disassyCode="99" '
        'disassyCodeVariant="Z" infoCode="999" infoCodeVariant="Z" '
        'itemLocationCode="Z"/>'
        "</dmRefIdent></dmRef></refs></content>"
    )
    path = tmp_path / DM_NAME
    path.write_text(make_dm(body=body), encoding="utf-8")

    checker, result = _check(path)

    assert "STR-IDENT" not in _rules(checker, result)


# ---------------------------------------------------------------------------
# Layer 4 -- DTD
# ---------------------------------------------------------------------------

def test_entity_only_internal_subset_is_not_treated_as_a_content_model(tmp_path):
    # The S1000D 4.x pattern: a DOCTYPE that exists purely to declare ICN
    # graphic entities. Validating against it would report every element in
    # the document as undeclared -- thousands of findings, all artefacts.
    doctype = (
        "<!DOCTYPE dmodule [\n"
        '<!ENTITY ICN-X SYSTEM "ICN-X.cgm" NDATA cgm>\n'
        "]>\n"
    )
    path = tmp_path / DM_NAME
    path.write_text(make_dm(doctype=doctype), encoding="utf-8")

    checker, result = _check(path)

    assert not [row for row in checker.findings(result) if row["Check"] == "dtd"]


def test_an_internal_subset_that_declares_elements_is_validated(tmp_path):
    doctype = (
        "<!DOCTYPE dmodule [\n"
        "<!ELEMENT dmodule (identAndStatusSection, content)>\n"
        "<!ELEMENT identAndStatusSection (dmAddress)>\n"
        "<!ELEMENT dmAddress (dmIdent)>\n"
        "<!ELEMENT dmIdent (dmCode)>\n"
        "<!ELEMENT dmCode EMPTY>\n"
        "<!ELEMENT content (illegal)*>\n"
        "]>\n"
    )
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(doctype=doctype, body="<content><notdeclared/></content>"),
        encoding="utf-8",
    )

    checker, result = _check(path)

    assert "DTD-VALID" in _rules(checker, result)


# ---------------------------------------------------------------------------
# Layer 5 -- schema
# ---------------------------------------------------------------------------

def test_a_schema_valid_object_produces_no_schema_findings(tmp_path, schema_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(schema="dmodule.xsd"), encoding="utf-8")

    checker, result = _check(path, search_paths=[schema_path.parent])

    assert not [row for row in checker.findings(result) if row["Check"] == "schema"]
    assert result["Schema"]["source"] == "local"


def test_schema_violations_are_reported_with_line_and_element(tmp_path, schema_path):
    # `content` is mandatory in the schema; omit it.
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(schema="dmodule.xsd", body=""),
        encoding="utf-8",
    )

    checker, result = _check(path, search_paths=[schema_path.parent])
    findings = [row for row in checker.findings(result) if row["Rule"] == "XSD-VALID"]

    assert findings
    assert findings[0]["Line"]
    assert findings[0]["Element"]


def test_an_unresolvable_schema_is_a_warning_that_names_itself(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(schema="http://example.invalid/nope.xsd"), encoding="utf-8"
    )

    checker, result = _check(path)
    finding = next(row for row in checker.findings(result) if row["Rule"] == "XSD-UNRESOLVED")

    # A run that could not check anything must not look like a clean one.
    assert finding["Status"] == "Warning"
    assert "network access is off" in finding["Detail"]
    assert result["Schema"]["source"] == "unresolved"


def test_a_schema_that_will_not_compile_is_reported_as_such(tmp_path):
    broken = tmp_path / "schemas" / "dmodule.xsd"
    broken.parent.mkdir(parents=True, exist_ok=True)
    broken.write_text("<xsd:schema/>", encoding="utf-8")
    path = tmp_path / DM_NAME
    path.write_text(make_dm(schema="dmodule.xsd"), encoding="utf-8")

    checker, result = _check(path, search_paths=[broken.parent])

    assert "XSD-BROKEN" in _rules(checker, result)


def test_a_schema_is_compiled_once_per_run_however_many_objects_use_it(tmp_path, schema_path):
    for index in range(4):
        name = DM_NAME.replace("040A", f"04{index}A")
        (tmp_path / name).write_text(
            make_dm(schema="dmodule.xsd",
                    ident=DM_IDENT.replace('infoCode="040"', f'infoCode="04{index}"')),
            encoding="utf-8",
        )

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.add_schema_search_path(str(schema_path.parent))
    checker.set_xml_dir(str(tmp_path))
    result = checker.validate()

    assert len(result) == 4
    # One compiled schema serves every object -- the whole reason a folder of
    # 90 data modules does not take 90 schema compilations.
    assert len(checker._schema_cache) == 1


# ---------------------------------------------------------------------------
# Directory mode and summaries
# ---------------------------------------------------------------------------

def test_directory_mode_returns_one_result_per_file(tmp_path, schema_path):
    (tmp_path / DM_NAME).write_text(make_dm(schema="dmodule.xsd"), encoding="utf-8")
    (tmp_path / DM_NAME.replace("040A", "041A")).write_text(
        make_dm(schema="dmodule.xsd", body=""), encoding="utf-8"
    )

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.add_schema_search_path(str(schema_path.parent))
    checker.set_xml_dir(str(tmp_path))
    result = checker.validate()
    totals = checker.run_summary(result)

    assert totals["DocumentsChecked"] == 2
    assert totals["DocumentsPassed"] == 1
    assert totals["DocumentsFailed"] == 1


def test_progress_callback_is_called_once_per_file(tmp_path):
    for index in range(3):
        (tmp_path / DM_NAME.replace("040A", f"04{index}A")).write_text(
            make_dm(), encoding="utf-8"
        )
    calls = []

    checker = XmlChecker()
    checker.set_xml_dir(str(tmp_path))
    checker.validate(progress_callback=lambda current, total, stage: calls.append(
        (current, total, stage)
    ))

    assert [_[0] for _ in calls] == [1, 2, 3]
    assert {_[2] for _ in calls} == {"files"}


def test_summary_counts_findings_by_layer_and_by_rule(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(b"\xef\xbb\xbf" + make_dm(
        ident=DM_IDENT.replace('itemLocationCode="D"', 'itemLocationCode="B"')
    ).encode("utf-8"))

    checker, result = _check(path)
    totals = checker.run_summary(result)

    assert totals["FindingsByRule"]["ENC-BOM"] == 1
    assert totals["FindingsByRule"]["STR-IDENT"] == 1
    assert totals["FindingsByCheck"]["encoding"] == 1


def test_checks_can_be_switched_off_individually(tmp_path):
    path = tmp_path / DM_NAME
    path.write_bytes(b"\xef\xbb\xbf" + make_dm().encode("utf-8"))

    checker, result = _check(path, check_encoding=False, check_structure=False)

    assert not checker.findings(result)


def test_validate_without_an_input_says_so(tmp_path):
    with pytest.raises(ValueError, match="set_xml"):
        XmlChecker().validate()


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------

class _Collector(HTMLParser):
    """Structural view of the HTML report: ids, and anything with a src/href."""

    def __init__(self):
        super().__init__()
        self.ids = []
        self.external_refs = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if "id" in attributes:
            self.ids.append(attributes["id"])
        for name in ("src", "href"):
            if name in attributes:
                self.external_refs.append(attributes[name])


def _parsed(html):
    collector = _Collector()
    collector.feed(html)
    return collector


def test_html_report_is_a_complete_self_contained_document(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(), encoding="utf-8")
    checker, result = _check(path)

    html = checker.to_html_report(result)
    parsed = _parsed(html)

    assert html.startswith("<!doctype html>")
    assert html.rstrip().endswith("</html>")
    # No stylesheet link, script src, image or webfont: the report has to open
    # identically off a network share or an e-mail attachment.
    assert parsed.external_refs == []
    assert "theme-toggle" in parsed.ids


def test_html_report_carries_a_row_per_finding_and_is_filterable(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(ident=DM_IDENT.replace('itemLocationCode="D"', 'itemLocationCode="B"')),
        encoding="utf-8",
    )
    checker, result = _check(path)

    html = checker.to_html_report(result)
    ids = set(_parsed(html).ids)

    assert {"findings", "finding-filter", "errors-only"} <= ids
    assert "STR-IDENT" in html


def test_html_report_shows_an_empty_state_for_a_clean_run(tmp_path, schema_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(schema="dmodule.xsd"), encoding="utf-8")
    checker, result = _check(path, search_paths=[schema_path.parent])

    html = checker.to_html_report(result)

    assert "No syntax or schema problems found." in html
    assert 'id="findings"' not in html


def test_html_report_escapes_markup_in_messages(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text("<dmodule><a></b></dmodule>", encoding="utf-8")
    checker, result = _check(path)

    html = checker.to_html_report(result)

    # Parser messages quote the offending markup; it must not become markup.
    assert "<b>" not in html.split("<script>")[0].replace("<b>irth", "")


def test_html_report_is_written_to_disk_when_a_path_is_given(tmp_path):
    path = tmp_path / DM_NAME
    path.write_text(make_dm(), encoding="utf-8")
    checker, result = _check(path)
    out = tmp_path / "report.html"

    returned = checker.to_html_report(result, str(out))

    assert out.exists()
    assert out.read_text(encoding="utf-8") == returned


def test_excel_report_has_a_summary_and_a_findings_sheet(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(ident=DM_IDENT.replace('itemLocationCode="D"', 'itemLocationCode="B"')),
        encoding="utf-8",
    )
    checker, result = _check(path)
    out = tmp_path / "report.xlsx"

    checker.to_excel_report(result, str(out))
    workbook = openpyxl.load_workbook(out)

    assert workbook.sheetnames[:2] == ["Summary", "Findings"]
    findings = workbook["Findings"]
    headers = [cell.value for cell in findings[1]]
    assert headers[:4] == ["Document", "Check", "Rule", "Status"]
    rules = [findings.cell(row=row, column=3).value
             for row in range(2, findings.max_row + 1)]
    assert "STR-IDENT" in rules


def test_excel_report_styles_the_header_the_way_the_brex_report_does(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / DM_NAME
    path.write_text(make_dm(), encoding="utf-8")
    checker, result = _check(path)
    out = tmp_path / "report.xlsx"

    checker.to_excel_report(result, str(out))
    header = openpyxl.load_workbook(out)["Findings"]["A1"]

    # Same shared `acd.report` layer as BrexChecker, so the two workbooks are
    # visibly one product.
    assert header.font.bold is True
    assert header.fill.fgColor.rgb.endswith("1F4E79")


def test_schemas_sheet_shows_where_each_schema_came_from(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / DM_NAME
    path.write_text(make_dm(schema="http://example.invalid/nope.xsd"), encoding="utf-8")
    checker, result = _check(path)
    out = tmp_path / "report.xlsx"

    checker.to_excel_report(result, str(out))
    sheet = openpyxl.load_workbook(out)["Schemas"]
    row = [sheet.cell(row=2, column=column).value for column in range(1, 5)]

    # The antidote to a run that looks clean because nothing was checked.
    assert row[0] == "http://example.invalid/nope.xsd"
    assert row[2] == "unresolved"


def test_json_report_carries_the_summary_and_every_finding(tmp_path):
    from json import loads

    path = tmp_path / DM_NAME
    path.write_text(
        make_dm(ident=DM_IDENT.replace('itemLocationCode="D"', 'itemLocationCode="B"')),
        encoding="utf-8",
    )
    checker, result = _check(path)

    payload = loads(checker.to_json_report(result))

    assert payload["summary"]["Errors"] == 1
    assert payload["documents"][0]["findings"][0]["Rule"] in CHECK_RULES
