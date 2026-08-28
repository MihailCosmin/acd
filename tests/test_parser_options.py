import pytest

from lxml import etree

from acd.brex_checker import BrexChecker, BrexNotFound

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

FLAG_0_BREX = """<brex>
  <contextRules>
    <structureObjectRuleGroup>
      <structureObjectRule>
        <objectPath allowedObjectFlag="0">//forbiddenElement</objectPath>
        <objectUse>forbiddenElement must not be present</objectUse>
      </structureObjectRule>
    </structureObjectRuleGroup>
  </contextRules>
</brex>
"""


@pytest.fixture
def flag_0_brex_path(tmp_path):
    path = tmp_path / "brex.xml"
    path.write_text(FLAG_0_BREX, encoding="utf-8")
    return str(path)


# --- XInclude -----------------------------------------------------------

def test_xinclude_disabled_by_default_leaves_include_unresolved(tmp_path, flag_0_brex_path):
    (tmp_path / "included.xml").write_text("<forbiddenElement/>", encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xi="http://www.w3.org/2001/XInclude" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        '<xi:include href="included.xml"/>\n'
        "</dmodule>\n",
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    checker.override_brex_list([flag_0_brex_path])

    result = checker._check_rules()
    assert result[flag_0_brex_path]['0'] == []


def test_xinclude_enabled_pulls_in_referenced_content(tmp_path, flag_0_brex_path):
    (tmp_path / "included.xml").write_text("<forbiddenElement/>", encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xi="http://www.w3.org/2001/XInclude" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        '<xi:include href="included.xml"/>\n'
        "</dmodule>\n",
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_xinclude(True)
    checker.set_xml(str(xml_path))
    checker.override_brex_list([flag_0_brex_path])

    result = checker._check_rules()
    assert len(result[flag_0_brex_path]['0']) == 1


# --- Entity resolution ----------------------------------------------------

def test_external_dtd_entity_unresolved_by_default(tmp_path):
    (tmp_path / "ext.dtd").write_text('<!ENTITY ext "External DTD Content">', encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(
        '<!DOCTYPE dmodule SYSTEM "ext.dtd">\n<dmodule><descr>&ext;</descr></dmodule>\n',
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_xml(str(xml_path))
    with pytest.raises(etree.XMLSyntaxError):
        checker._parse_xml_file(str(xml_path))


def test_external_dtd_entity_resolved_when_enabled(tmp_path):
    (tmp_path / "ext.dtd").write_text('<!ENTITY ext "External DTD Content">', encoding="utf-8")
    xml_path = tmp_path / "object.xml"
    xml_path.write_text(
        '<!DOCTYPE dmodule SYSTEM "ext.dtd">\n<dmodule><descr>&ext;</descr></dmodule>\n',
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_entity_resolution(load_external_dtd=True)
    checker.set_xml(str(xml_path))
    tree = checker._parse_xml_file(str(xml_path))
    assert tree.getroot().find("descr").text == "External DTD Content"


def test_set_entity_resolution_disallows_network_unless_dtd_loading_enabled():
    checker = BrexChecker()
    checker.set_entity_resolution(load_external_dtd=False, allow_network=True)
    assert checker._allow_network is False

    checker.set_entity_resolution(load_external_dtd=True, allow_network=True)
    assert checker._allow_network is True


# --- XML catalogs -----------------------------------------------------------

def test_set_xml_catalog_registers_file_in_env_var(tmp_path, monkeypatch):
    monkeypatch.delenv("XML_CATALOG_FILES", raising=False)
    catalog_path = tmp_path / "catalog.xml"
    catalog_path.write_text(
        '<catalog xmlns="urn:oasis:names:tc:entity:xmlns:xml:catalog"/>', encoding="utf-8"
    )

    checker = BrexChecker()
    checker.set_xml_catalog(str(catalog_path))

    import os
    assert str(catalog_path) in os.environ["XML_CATALOG_FILES"].split()


def test_set_xml_catalog_appends_without_duplicating(tmp_path, monkeypatch):
    monkeypatch.delenv("XML_CATALOG_FILES", raising=False)
    first = tmp_path / "first.xml"
    second = tmp_path / "second.xml"
    for path in (first, second):
        path.write_text('<catalog xmlns="urn:oasis:names:tc:entity:xmlns:xml:catalog"/>', encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_catalog(str(first))
    checker.set_xml_catalog(str(second))
    checker.set_xml_catalog(str(first))

    import os
    entries = os.environ["XML_CATALOG_FILES"].split()
    assert entries.count(str(first)) == 1
    assert entries.count(str(second)) == 1


def test_set_xml_catalog_rejects_missing_file(tmp_path):
    checker = BrexChecker()
    with pytest.raises(BrexNotFound):
        checker.set_xml_catalog(str(tmp_path / "does_not_exist.xml"))


# --- ignore-empty -----------------------------------------------------------

def test_ignore_empty_disabled_raises_on_empty_file(tmp_path):
    empty_path = tmp_path / "empty.xml"
    empty_path.write_text("", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml(str(empty_path))
    with pytest.raises(Exception):
        checker.validate()


def test_ignore_empty_skips_empty_file_in_single_mode(tmp_path):
    empty_path = tmp_path / "empty.xml"
    empty_path.write_text("", encoding="utf-8")

    checker = BrexChecker()
    checker.set_ignore_empty(True)
    checker.set_xml(str(empty_path))
    result = checker.validate()
    assert result == {"Skipped": True, "Summary": "Skipped (empty or non-XML file)"}


def test_ignore_empty_skips_non_xml_file_in_single_mode(tmp_path):
    bad_path = tmp_path / "not_xml.xml"
    bad_path.write_text("this is not xml at all", encoding="utf-8")

    checker = BrexChecker()
    checker.set_ignore_empty(True)
    checker.set_xml(str(bad_path))
    result = checker.validate()
    assert result["Skipped"] is True


def test_ignore_empty_skips_bad_files_in_directory_mode(tmp_path, flag_0_brex_path):
    object_dir = tmp_path / "objects"
    object_dir.mkdir()
    (object_dir / "empty.xml").write_text("", encoding="utf-8")
    (object_dir / "clean.xml").write_text(
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}"/>\n',
        encoding="utf-8",
    )

    checker = BrexChecker()
    checker.set_ignore_empty(True)
    checker.set_xml_dir(str(object_dir))
    checker.override_brex_list([flag_0_brex_path])

    results = checker.validate()
    assert "empty.xml" not in results
    assert "clean.xml" in results


def test_ignore_empty_without_flag_raises_in_directory_mode(tmp_path, flag_0_brex_path):
    object_dir = tmp_path / "objects"
    object_dir.mkdir()
    (object_dir / "empty.xml").write_text("", encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(object_dir))
    checker.override_brex_list([flag_0_brex_path])

    with pytest.raises(Exception):
        checker.validate()
