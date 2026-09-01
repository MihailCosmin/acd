"""Resolution fallback and the last-line-of-defence prompt.

Both checkers resolve a reference the same way -- the checked object's own
folder, then registered search paths, then the `res/` tree shipped beside the
application, and only then a prompt. These tests pin that order, the
discovery of `res/` without anything being passed in, and above all the
guarantee that a prompt can never fire where there is nobody to answer it.
"""

from os.path import basename
from os.path import join

import pytest

from acd import prompt as prompt_module
from acd import resources
from acd.brex_checker import BrexChecker
from acd.brex_checker import NoBrexDefined
from acd.xml_checker import XmlChecker

SCHEMA_URL = "http://www.s1000d.org/S1000D_4-1/xml_schema_flat/descript.xsd"

SCHEMA = """<?xml version="1.0" encoding="UTF-8"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <xsd:element name="dmodule">
    <xsd:complexType>
      <xsd:sequence>
        <xsd:element name="content"/>
      </xsd:sequence>
    </xsd:complexType>
  </xsd:element>
</xsd:schema>
"""

DM = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    f'<dmodule xsi:noNamespaceSchemaLocation="{SCHEMA_URL}"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
    "  <content/>\n"
    "</dmodule>\n"
)

DM_NAME = "DMC-H160-B-67-34-0200-00A-040A-D_001-01_SX-US.xml"


@pytest.fixture
def no_res_tree(monkeypatch, tmp_path):
    """Neither bundled resources nor an application tree.

    Both have to be neutralised. Discovery walks out from the running script
    and the working directory, so it would otherwise find a real application's
    `res/`; and the package now ships its own `acd/xsd`, which resolves most
    S1000D schemas outright. A test about an *unresolvable* reference has to
    take away both.
    """
    absent = str(tmp_path / "no-such-resources")
    monkeypatch.setattr(resources, "BUNDLED_XSD_DIR", absent)
    monkeypatch.setattr(resources, "BUNDLED_BREX_DIR", absent)
    resources._ROOT_CACHE["root"] = None  # pylint: disable=protected-access
    yield
    resources._ROOT_CACHE.clear()  # pylint: disable=protected-access


@pytest.fixture
def no_bundled(monkeypatch, tmp_path):
    """Only the application's `res/` tree, so a test can pin what it supplies
    without the bundled copies answering first."""
    absent = str(tmp_path / "no-such-bundle")
    monkeypatch.setattr(resources, "BUNDLED_XSD_DIR", absent)
    monkeypatch.setattr(resources, "BUNDLED_BREX_DIR", absent)


@pytest.fixture
def res_tree(tmp_path, monkeypatch):
    """A `res/` tree in the shape an application ships, pointed at by
    `ACD_RES_DIR` so discovery does not depend on where the suite is run."""
    root = tmp_path / "app" / "res"
    (root / "xsd" / "4.1").mkdir(parents=True)
    (root / "xsd" / "4.1" / "descript.xsd").write_text(SCHEMA, encoding="utf-8")
    (root / "brex" / "brex_4_2").mkdir(parents=True)
    monkeypatch.setenv(resources.RES_DIR_ENV, str(root))
    resources.find_resource_root(refresh=True)
    yield root
    resources._ROOT_CACHE.clear()  # pylint: disable=protected-access


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------

def test_the_package_ships_its_own_schemas_entities_and_brex():
    # The point of bundling: a checkout or an install resolves an S1000D
    # schema with no configuration, no application tree and no network.
    issues = {basename(_) for _ in resources.schema_dirs()}

    assert {"4.1", "4.2"} <= issues
    assert resources.catalog_path() == resources.BUNDLED_CATALOG
    assert resources.find_schema(SCHEMA_URL).startswith(resources.BUNDLED_XSD_DIR)


def test_resource_root_is_found_from_the_environment(res_tree):
    assert resources.find_resource_root(refresh=True) == str(res_tree)


def test_an_application_tree_is_searched_after_the_bundled_one(res_tree):
    directories = resources.schema_dirs()
    from_app = str(res_tree / "xsd" / "4.1")
    from_package = [_ for _ in directories if _.startswith(resources.BUNDLED_XSD_DIR)]

    # Bundled issues come first, then whatever the application adds, so an
    # application can supply an issue the package does not carry without
    # shadowing the ones it does.
    assert from_package
    assert from_app in directories
    assert directories.index(from_app) > directories.index(from_package[-1])


def test_the_bundled_copy_wins_over_an_application_tree(res_tree):
    found = resources.find_schema(SCHEMA_URL)

    assert found.startswith(resources.BUNDLED_XSD_DIR)
    assert not found.startswith(str(res_tree))


def test_an_application_tree_supplies_what_the_bundle_lacks(res_tree, no_bundled):
    assert resources.find_schema(SCHEMA_URL) == str(
        res_tree / "xsd" / "4.1" / "descript.xsd"
    )


def test_a_directory_called_res_without_schemas_or_brex_is_not_a_resource_root(
        tmp_path, monkeypatch):
    decoy = tmp_path / "res"
    decoy.mkdir()
    monkeypatch.setenv(resources.RES_DIR_ENV, str(decoy))

    assert resources.find_resource_root(refresh=True) != str(decoy)
    resources._ROOT_CACHE.clear()  # pylint: disable=protected-access


def test_schema_is_found_by_the_issue_its_url_names(res_tree, no_bundled):
    assert resources.find_schema(SCHEMA_URL) == str(
        res_tree / "xsd" / "4.1" / "descript.xsd"
    )


def test_an_unknown_issue_falls_back_to_any_directory_holding_the_file(
        res_tree, no_bundled):
    # A CSDB is not always consistent about which issue's URL it quotes, and
    # validating against a near neighbour beats not validating at all -- the
    # report's `source` column is what keeps that honest.
    other = SCHEMA_URL.replace("S1000D_4-1", "S1000D_5-0")

    assert resources.find_schema(other) == str(
        res_tree / "xsd" / "4.1" / "descript.xsd"
    )


def test_a_schema_no_version_directory_has_is_not_invented(res_tree, no_bundled):
    assert resources.find_schema(
        "http://www.s1000d.org/S1000D_4-1/xml_schema_flat/nosuch.xsd"
    ) is None


# ---------------------------------------------------------------------------
# Resolution order
# ---------------------------------------------------------------------------

def test_schema_resolves_from_the_res_tree_with_nothing_configured(
        tmp_path, res_tree, no_bundled):
    path = tmp_path / DM_NAME
    path.write_text(DM, encoding="utf-8")

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.set_xml(str(path))
    result = checker.validate()

    # No add_schema_search_path call, no cache, no network: the whole point.
    assert result["Schema"]["source"] == "bundled"
    assert not checker.unresolved_schemas(result)


def test_a_registered_search_path_wins_over_the_res_tree(tmp_path, res_tree, no_bundled):
    local = tmp_path / "local"
    local.mkdir()
    (local / "descript.xsd").write_text(SCHEMA, encoding="utf-8")
    path = tmp_path / DM_NAME
    path.write_text(DM, encoding="utf-8")

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.add_schema_search_path(str(local))
    checker.set_xml(str(path))
    result = checker.validate()

    assert result["Schema"]["source"] == "local"
    assert result["Schema"]["resolved"] == str(local / "descript.xsd")


def test_brex_directories_are_bundled_first_then_the_application_tree(res_tree):
    found = resources.brex_dirs(include_default=False)
    from_app = str(res_tree / "brex" / "brex_4_2")

    assert found[0].startswith(resources.BUNDLED_BREX_DIR)
    assert from_app in found
    assert found.index(from_app) > 0


def test_the_flat_bundled_brex_directory_is_only_added_on_request():
    # BrexChecker passes include_default=False: it reaches the flat master set
    # through the default-BREX fallback instead, which matches on the logical
    # DMC and records the substitution for the report.
    without = resources.brex_dirs(include_default=False)
    with_default = resources.brex_dirs(include_default=True)

    assert resources.BUNDLED_BREX_DIR not in without
    assert with_default[-1] == resources.BUNDLED_BREX_DIR


# ---------------------------------------------------------------------------
# Entity catalog
# ---------------------------------------------------------------------------

def test_the_bundled_catalog_covers_the_issues_the_schemas_do():
    from lxml import etree

    catalog = etree.parse(resources.catalog_path())
    starts = catalog.getroot().xpath(
        "//*[local-name()='rewriteSystem']/@systemIdStartString"
    )

    # Every entry maps an s1000d.org entity URL to a local directory; without
    # them a data module carrying &ndash; needs network access to parse.
    assert starts
    assert all(_.startswith("http://www.s1000d.org/") for _ in starts)


def test_register_catalog_adds_the_bundled_catalog_to_the_environment(monkeypatch):
    monkeypatch.delenv("XML_CATALOG_FILES", raising=False)

    registered = resources.register_catalog()

    assert registered == resources.BUNDLED_CATALOG
    assert registered in os_environ()["XML_CATALOG_FILES"]


def test_register_catalog_is_idempotent(monkeypatch):
    monkeypatch.delenv("XML_CATALOG_FILES", raising=False)

    resources.register_catalog()
    resources.register_catalog()

    # libxml2 reads the variable once per process; a duplicated entry is
    # harmless but a growing one is a leak across a long-running application.
    assert os_environ()["XML_CATALOG_FILES"].split().count(
        resources.BUNDLED_CATALOG) == 1


def test_constructing_a_checker_registers_the_catalog(monkeypatch):
    monkeypatch.delenv("XML_CATALOG_FILES", raising=False)

    XmlChecker()

    assert resources.BUNDLED_CATALOG in os_environ()["XML_CATALOG_FILES"]


def os_environ():
    """`os.environ`, imported here so the tests above read as prose."""
    from os import environ
    return environ


# ---------------------------------------------------------------------------
# The prompt guard -- the property that matters most
# ---------------------------------------------------------------------------

def test_prompting_is_off_under_pytest():
    # This is what stops a modal dialog hanging the suite until it is killed.
    assert prompt_module.prompting_enabled() is False


def test_ask_for_folder_returns_none_when_prompting_is_disabled():
    assert prompt_module.ask_for_folder("schemas", ["a.xsd", "b.xsd"]) is None


def test_ask_for_folder_returns_none_with_nothing_missing():
    assert prompt_module.ask_for_folder("schemas", []) is None


def test_no_prompt_environment_variable_disables_prompting(monkeypatch):
    monkeypatch.setenv(prompt_module.NO_PROMPT_ENV, "1")

    assert prompt_module.prompting_enabled() is False


def test_set_prompting_toggles_the_module_switch(monkeypatch):
    monkeypatch.setattr(prompt_module, "_ENABLED", True)
    prompt_module.set_prompting(False)
    try:
        assert prompt_module.prompting_enabled() is False
    finally:
        prompt_module.set_prompting(True)


def test_an_unresolvable_schema_is_reported_not_prompted_for_in_batch(tmp_path, no_res_tree):
    path = tmp_path / DM_NAME
    path.write_text(DM, encoding="utf-8")

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.set_xml(str(path))
    result = checker.validate()

    # Prompting is unavailable here, so the run degrades to reporting rather
    # than blocking -- exactly how a service or CI job must behave.
    assert checker.unresolved_schemas(result) == [SCHEMA_URL]
    assert result["Schema"]["source"] == "unresolved"


# ---------------------------------------------------------------------------
# Prompt-and-retry, with the dialog stubbed
# ---------------------------------------------------------------------------

def test_a_chosen_folder_resolves_the_schema_and_the_object_is_rechecked(
        tmp_path, monkeypatch, no_res_tree):
    chosen = tmp_path / "chosen"
    chosen.mkdir()
    (chosen / "descript.xsd").write_text(SCHEMA, encoding="utf-8")
    path = tmp_path / DM_NAME
    path.write_text(DM, encoding="utf-8")

    asked = []

    def fake_ask(what, missing, **_kwargs):
        asked.append((what, list(missing)))
        return str(chosen)

    monkeypatch.setattr("acd.xml_checker.ask_for_folder", fake_ask)

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.set_xml(str(path))
    result = checker.validate()

    assert asked == [("schemas", [SCHEMA_URL])]
    assert result["Schema"]["source"] == "local"
    assert not checker.unresolved_schemas(result)


def test_one_prompt_covers_a_whole_folder_run(tmp_path, monkeypatch, no_res_tree):
    objects = tmp_path / "objects"
    objects.mkdir()
    for index in range(5):
        (objects / DM_NAME.replace("040A", f"04{index}A")).write_text(
            DM, encoding="utf-8"
        )
    chosen = tmp_path / "chosen"
    chosen.mkdir()
    (chosen / "descript.xsd").write_text(SCHEMA, encoding="utf-8")

    calls = []

    def fake_ask(what, missing, **_kwargs):
        calls.append(list(missing))
        return str(chosen)

    monkeypatch.setattr("acd.xml_checker.ask_for_folder", fake_ask)

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.set_xml_dir(str(objects))
    result = checker.validate()

    # Five objects, one missing schema between them, one dialog.
    assert calls == [[SCHEMA_URL]]
    assert not checker.unresolved_schemas(result)
    assert checker.run_summary(result)["DocumentsChecked"] == 5


def test_a_declined_prompt_leaves_the_findings_alone(tmp_path, monkeypatch, no_res_tree):
    path = tmp_path / DM_NAME
    path.write_text(DM, encoding="utf-8")
    monkeypatch.setattr("acd.xml_checker.ask_for_folder", lambda *a, **k: None)

    checker = XmlChecker()
    checker.set_allow_network(False)
    checker.set_schema_cache(None)
    checker.set_xml(str(path))
    result = checker.validate()

    assert checker.unresolved_schemas(result) == [SCHEMA_URL]


def test_brex_prompt_names_the_missing_reference_and_is_asked_once(
        tmp_path, monkeypatch, no_res_tree):
    dm = (
        '<?xml version="1.0"?>\n'
        '<dmodule xsi:noNamespaceSchemaLocation='
        '"http://www.s1000d.org/S1000D_4-2/xml_schema_flat/descript.xsd"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<identAndStatusSection><dmAddress><dmIdent>"
        '<dmCode modelIdentCode="H160" systemDiffCode="B" systemCode="67"'
        ' subSystemCode="3" subSubSystemCode="4" assyCode="0200" disassyCode="00"'
        ' disassyCodeVariant="A" infoCode="040" infoCodeVariant="A"'
        ' itemLocationCode="D"/>'
        "</dmIdent></dmAddress><dmStatus><brexDmRef><dmRef><dmRefIdent>"
        '<dmCode modelIdentCode="NOSUCH" systemDiffCode="Z" systemCode="99"'
        ' subSystemCode="9" subSubSystemCode="9" assyCode="9999" disassyCode="99"'
        ' disassyCodeVariant="Z" infoCode="022" infoCodeVariant="A"'
        ' itemLocationCode="D"/>'
        "</dmRefIdent></dmRef></brexDmRef></dmStatus>"
        "</identAndStatusSection><content/></dmodule>"
    )
    path = tmp_path / DM_NAME
    path.write_text(dm, encoding="utf-8")

    calls = []

    def fake_ask(what, missing, **_kwargs):
        calls.append((what, list(missing)))
        return None  # decline, so the usual exception still surfaces

    monkeypatch.setattr("acd.brex_checker.ask_for_folder", fake_ask)

    checker = BrexChecker()
    checker.set_xml(str(path))
    with pytest.raises(NoBrexDefined):
        checker.validate()

    assert len(calls) == 1
    what, missing = calls[0]
    assert what == "BREX data modules"
    assert missing and "NOSUCH" in missing[0]
