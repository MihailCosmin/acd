"""`BrexChecker.to_html_report` -- the formatted, self-contained HTML report
(category D6), the browser-readable counterpart of `to_excel_report`.

Covers both result shapes (single object and batch/folder), the content the
report is built from, HTML escaping, the dark/light theming contract and the
fact that the document pulls in nothing from the network.
"""

from html.parser import HTMLParser

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

SEVERITY_LEVELS_CONTENT = """<brSeverityLevels>
  <brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>
  <brSeverityLevel value="brsl02" fail="no">Warning</brSeverityLevel>
</brSeverityLevels>
"""

BREX_CONTENT = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="R-FORBIDDEN" brSeverityLevel="brsl01">
        <objectPath allowedObjectFlag="0">//forbidden</objectPath>
        <objectUse>forbidden &lt;element&gt; must not be present &amp; is an error.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="R-DISCOURAGED" brSeverityLevel="brsl02">
        <objectPath allowedObjectFlag="0">//discouraged</objectPath>
        <objectUse>discouraged should not be present, but only warns.</objectUse>
      </structureObjectRule>
      <structureObjectRule id="R-CODE">
        <objectPath allowedObjectFlag="2">//coded/@code</objectPath>
        <objectUse>coded/@code must be aa or two digits.</objectUse>
        <objectValue valueForm="single" valueAllowed="aa"/>
        <objectValue valueForm="pattern" valueAllowed="[0-9]{2}"/>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


class _Collector(HTMLParser):
    """Minimal structural view of the report: tags seen, ids, classes, and the
    `data-status` of every violation row."""

    def __init__(self):
        super().__init__()
        self.tags = []
        self.ids = []
        self.classes = []
        self.row_statuses = []
        self.external_refs = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        self.tags.append(tag)
        if attributes.get("id"):
            self.ids.append(attributes["id"])
        if attributes.get("class"):
            self.classes.extend(attributes["class"].split())
        if tag == "tr" and attributes.get("data-status"):
            self.row_statuses.append(attributes["data-status"])
        for attribute in ("src", "href"):
            if attributes.get(attribute):
                self.external_refs.append(attributes[attribute])


def make_dm(body: str = "") -> str:
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        f"{body}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def severity_levels_path(tmp_path):
    path = tmp_path / ".brseveritylevels"
    path.write_text(SEVERITY_LEVELS_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def object_dir(tmp_path):
    objects = tmp_path / "objects"
    objects.mkdir()
    (objects / "clean.xml").write_text(make_dm('  <coded code="aa"/>'), encoding="utf-8")
    (objects / "bad.xml").write_text(
        make_dm('  <forbidden/>\n  <discouraged/>\n  <coded code="zz"/>'), encoding="utf-8"
    )
    return str(objects)


def _batch_html(object_dir, brex_path, severity_levels_path, **kwargs):
    checker = BrexChecker()
    checker.set_xml_dir(object_dir)
    checker.override_brex_list([brex_path])
    checker.set_severity_levels_path(severity_levels_path)
    return checker.to_html_report(checker.validate(), **kwargs)


def _single_html(tmp_path, brex_path, body, **kwargs):
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(make_dm(body), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([brex_path])
    return checker.to_html_report(checker.validate(), **kwargs)


def _parsed(html):
    collector = _Collector()
    collector.feed(html)
    return collector


# ---------------------------------------------------------------------------
# Structure and content
# ---------------------------------------------------------------------------

def test_report_is_a_complete_html_document(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, "  <forbidden/>")

    assert html.startswith("<!doctype html>")
    assert html.rstrip().endswith("</html>")
    assert "<title>BREX check report</title>" in html
    tags = _parsed(html).tags
    assert {"html", "head", "body", "style", "script", "table"} <= set(tags)


def test_batch_report_carries_a_row_per_violation_and_per_document(object_dir, brex_path,
                                                                   severity_levels_path):
    html = _batch_html(object_dir, brex_path, severity_levels_path)
    parsed = _parsed(html)

    # One row per violation, in the same flag ('0'/'1'/'2') order every other
    # report uses: the two flag-0 rules first (an error and, per the
    # .brseveritylevels fail="no" entry, a warning), then the value one.
    assert parsed.row_statuses == ["error", "warning", "error"]
    assert "bad.xml" in html and "clean.xml" in html
    # The per-document pass/fail table is what makes a folder run readable.
    assert "Passed" in html and "Failed" in html


def test_summary_cards_carry_the_run_totals(object_dir, brex_path, severity_levels_path):
    html = _batch_html(object_dir, brex_path, severity_levels_path)

    assert '<span class="card-value">2</span><span class="card-label">Documents checked</span>' in html
    assert '<span class="card-value">1</span><span class="card-label">Failed</span>' in html
    assert '<span class="card-value">1</span><span class="card-label">Warnings</span>' in html
    # Severity chips mirror run_summary()['ViolationsBySeverity'].
    assert '<span class="chip-key">brsl01</span>' in html


def test_value_violation_shows_what_was_found_and_what_is_allowed(object_dir, brex_path,
                                                                  severity_levels_path):
    html = _batch_html(object_dir, brex_path, severity_levels_path)

    assert "Element/Attribute (zz) did not match the object values." in html
    assert "Allowed (single):" in html
    assert "<code>[0-9]{2}</code>" in html


def test_single_object_report_omits_the_documents_table(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, "  <forbidden/>")

    assert "<span class=\"section-title\">Documents</span>" not in html
    assert "<span class=\"section-title\">Violations</span>" in html


def test_clean_run_reports_an_empty_state_instead_of_a_table(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, '  <coded code="42"/>')

    assert "No BREX violations found." in html
    assert 'id="violations"' not in html


def test_markup_in_rule_text_and_node_snippets_is_escaped(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, "  <forbidden/>")

    # The objectUse itself contains < > &, and the node snippet is XML.
    assert "forbidden &lt;element&gt; must not be present &amp; is an error." in html
    assert "&lt;forbidden" in html
    assert "<forbidden" not in html.split("<body>")[1]


def test_report_is_written_to_disk_when_a_path_is_given(tmp_path, brex_path):
    destination = tmp_path / "report.html"
    html = _single_html(tmp_path, brex_path, "  <forbidden/>", path=str(destination))

    assert destination.read_text(encoding="utf-8") == html


def test_title_is_configurable(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, "  <forbidden/>", title="CMP 21-77-05 BREX run")

    assert "<title>CMP 21-77-05 BREX run</title>" in html
    assert "<h1>CMP 21-77-05 BREX run</h1>" in html


def test_xpath_errors_are_reported_as_their_own_section(tmp_path):
    brex = tmp_path / "broken_brex.xml"
    brex.write_text(
        """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule id="R-BROKEN">
        <objectPath allowedObjectFlag="0">//[[[broken</objectPath>
        <objectUse>A rule whose objectPath does not compile.</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
""",
        encoding="utf-8",
    )
    html = _single_html(tmp_path, str(brex), "  <anything/>")

    assert "XPath errors" in html
    assert "//[[[broken" in html


# ---------------------------------------------------------------------------
# Theming and self-containment
# ---------------------------------------------------------------------------

def test_dark_and_light_are_both_defined_and_toggleable(tmp_path, brex_path):
    html = _single_html(tmp_path, brex_path, "  <forbidden/>")

    # Full light palette on bare :root, dark redefined under the media query
    # (guarded so an explicit light choice wins) and again under the toggle's
    # attribute, so the toggle overrides the reader's OS setting both ways.
    assert "@media (prefers-color-scheme: dark)" in html
    assert ':root:not([data-theme="light"])' in html
    assert ':root[data-theme="dark"]' in html
    assert "theme-toggle" in _parsed(html).ids
    # Storage access is guarded -- a file:// report with site data blocked
    # must still render.
    assert "try {" in html and "localStorage" in html


def test_filter_controls_are_present_when_there_are_violations(object_dir, brex_path,
                                                               severity_levels_path):
    ids = _parsed(_batch_html(object_dir, brex_path, severity_levels_path)).ids

    assert {"violation-filter", "errors-only", "violations"} <= set(ids)


def test_report_makes_no_external_requests(object_dir, brex_path, severity_levels_path):
    html = _batch_html(object_dir, brex_path, severity_levels_path)

    # Nothing with a src/href at all -- no stylesheet link, script src, image
    # or webfont -- so the report renders identically off a network share or
    # an e-mail attachment. (A schema URL inside an escaped node snippet is
    # text, not a request, which is why this checks attributes, not the
    # document text.)
    assert _parsed(html).external_refs == []
    assert "<link" not in html
    assert "<script src" not in html
