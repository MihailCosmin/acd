import json
from os import mkdir
from os.path import basename
from os.path import join

import pytest

from acd.brex_checker import BrexChecker

DMODULE_SCHEMA = "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/dmodule.xsd"

BREX_CONTENT = """<brex>
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


def make_xml(extra: str = "") -> str:
    return (
        '<dmodule xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xsi:noNamespaceSchemaLocation="{DMODULE_SCHEMA}">\n'
        f"{extra}\n"
        "</dmodule>\n"
    )


@pytest.fixture
def brex_path(tmp_path):
    brex_dir = tmp_path / "brex_dir"
    brex_dir.mkdir()
    path = brex_dir / "brex.xml"
    path.write_text(BREX_CONTENT, encoding="utf-8")
    return str(path)


@pytest.fixture
def object_dir(tmp_path):
    objects = tmp_path / "objects"
    objects.mkdir()
    (objects / "clean.xml").write_text(make_xml(), encoding="utf-8")
    (objects / "violating.xml").write_text(make_xml("<forbiddenElement/>"), encoding="utf-8")
    return str(objects)


def test_directory_mode_returns_mapping_of_filename_to_result(object_dir, brex_path):
    checker = BrexChecker()
    checker.set_xml_dir(object_dir)
    checker.override_brex_list([brex_path])

    results = checker.validate()

    assert set(results.keys()) == {"clean.xml", "violating.xml"}
    assert results["clean.xml"][brex_path]['0'] == []
    assert len(results["violating.xml"][brex_path]['0']) == 1
    assert results["clean.xml"]["Summary"] == "0 Errors"
    assert results["violating.xml"]["Summary"] == "1 Errors"


def test_directory_mode_empty_directory_does_not_raise(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()

    checker = BrexChecker()
    checker.set_xml_dir(str(empty_dir))

    assert checker.validate() == {}


def test_directory_mode_debug_output_is_valid_json(object_dir, brex_path, tmp_path, monkeypatch):
    monkeypatch.setattr("acd.brex_checker.expanduser", lambda _: str(tmp_path))

    checker = BrexChecker()
    checker.set_xml_dir(object_dir)
    checker.override_brex_list([brex_path])

    checker.validate(debug=True)

    report_path = tmp_path / f"Errors_{basename(object_dir)}.json"
    with open(report_path, encoding="utf-8") as f:
        data = json.load(f)

    assert set(data.keys()) == {"clean.xml", "violating.xml"}


def test_directory_mode_preserves_explicit_override_brex_list_across_files(object_dir, brex_path):
    checker = BrexChecker()
    checker.set_xml_dir(object_dir)
    checker.override_brex_list([brex_path])

    results = checker.validate()

    for filename in ("clean.xml", "violating.xml"):
        assert brex_path in results[filename]


def test_directory_mode_uses_real_extension_test_not_substring(tmp_path, brex_path):
    # Ref §3.11: `".xml" in name` matched "foo.xml.bak"; a real extension
    # test must not.
    objects = tmp_path / "objects_with_decoy"
    objects.mkdir()
    (objects / "clean.xml").write_text(make_xml(), encoding="utf-8")
    (objects / "backup.xml.bak").write_text(make_xml(), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(objects))
    checker.override_brex_list([brex_path])

    results = checker.validate()

    assert set(results.keys()) == {"clean.xml"}


def test_directory_mode_matches_uppercase_xml_extension(tmp_path, brex_path):
    # Real S1000D deliveries name data modules with an uppercase .XML
    # extension; the extension test must be case-insensitive.
    objects = tmp_path / "objects_uppercase"
    objects.mkdir()
    (objects / "CLEAN.XML").write_text(make_xml(), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(objects))
    checker.override_brex_list([brex_path])

    results = checker.validate()

    assert set(results.keys()) == {"CLEAN.XML"}


def test_directory_mode_ignores_a_directory_named_like_an_xml_file(object_dir, brex_path):
    mkdir(join(object_dir, "weird.xml"))

    checker = BrexChecker()
    checker.set_xml_dir(object_dir)
    checker.override_brex_list([brex_path])

    results = checker.validate()

    assert "weird.xml" not in results
    assert set(results.keys()) == {"clean.xml", "violating.xml"}


def test_directory_mode_checks_brex_named_data_modules_like_any_other_object(tmp_path, brex_path):
    # Ref §3.11: `"-022a-" not in name` meant a BREX data module (info code
    # 022A) was never validated, including against itself. s1kd validates
    # BREX objects like any other -- the directory filter must not single
    # them out.
    objects = tmp_path / "objects_with_brex_dm"
    objects.mkdir()
    (objects / "clean.xml").write_text(make_xml(), encoding="utf-8")
    brex_named_file = "DMC-TEST-A-00-00-00-00A-022A-D_001-00_EN-US.XML"
    (objects / brex_named_file).write_text(make_xml("<forbiddenElement/>"), encoding="utf-8")

    checker = BrexChecker()
    checker.set_xml_dir(str(objects))
    checker.override_brex_list([brex_path])

    results = checker.validate()

    assert brex_named_file in results
    assert len(results[brex_named_file][brex_path]['0']) == 1
