"""Fixture matrix for brex_checker_rework.md §4.5: every `allowedObjectFlag`
against each kind of XPath result a compiled `objectPath` can produce --
node-set, boolean, string and numeric. Ref §3.2 (flag 0/1 boolean handling)
and §3.9 (element-vs-string comparison for flag 2).

Exercises `_check_object_flag_0/1/2` directly (same style as
test_xpath_error_handling.py / test_flag1_value_checking.py) so each result
type can be constructed precisely with a plain XPath 2.0 function call,
independent of BREX rule-selection plumbing.
"""

import elementpath
import pytest
from lxml import etree

from acd.brex_checker import BrexChecker

NODESET_ROOT = etree.fromstring(b"<root><foo>bar</foo></root>")


def make_value(xpath, root, flag, **overrides):
    value = {
        'xpath': xpath,
        'Brex': 'dummy_brex',
        'ObjectFlag': flag,
        'objectUse': f'rule for {xpath}',
        'contextRules': '',
        'values_allowed': [],
        'regex_allowed': [],
        'ranges_allowed': [],
        'ruleId': None,
        'brDecisionIdentNumber': None,
        'brSeverityLevel': None,
        'value_tailoring': [],
    }
    value.update(overrides)
    value['selector'] = elementpath.Selector(xpath, namespaces={}, parser=elementpath.XPath2Parser)
    value['selector_error'] = None
    return value


def make_violations():
    return {'dummy_brex': {'0': [], '1': [], '2': [], 'xpathError': []}}


@pytest.fixture
def checker():
    return BrexChecker()


# ---------------------------------------------------------------------------
# Flag 0: "must not be present / condition must not hold"
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("xpath,root,violated", [
    ("//foo", NODESET_ROOT, True),          # node-set, non-empty -> violation
    ("//missing", NODESET_ROOT, False),     # node-set, empty -> no violation
    ("//foo = 'bar'", NODESET_ROOT, True),  # boolean true -> violation
    ("//foo = 'zzz'", NODESET_ROOT, False),  # boolean false -> no violation
    ("string(//foo)", NODESET_ROOT, True),   # non-empty string -> violation
    ("string(//missing)", NODESET_ROOT, False),  # empty string -> no violation
    ("count(//foo)", NODESET_ROOT, True),    # non-zero number -> violation
    ("count(//missing)", NODESET_ROOT, False),  # zero -> no violation
], ids=[
    "nodeset-hit", "nodeset-empty",
    "boolean-true", "boolean-false",
    "string-nonempty", "string-empty",
    "number-nonzero", "number-zero",
])
def test_flag_0_result_types(checker, xpath, root, violated):
    value = make_value(xpath, root, '0')
    violations = make_violations()

    checker._check_object_flag_0('schema', violations, root, value)

    assert len(violations['dummy_brex']['0']) == (1 if violated else 0)


def test_flag_0_nodeset_violation_carries_a_real_node_xpath(checker):
    value = make_value("//foo", NODESET_ROOT, '0')
    violations = make_violations()

    checker._check_object_flag_0('schema', violations, NODESET_ROOT, value)

    entry = violations['dummy_brex']['0'][0]
    assert entry['NodeXpath'] == '/root[1]/foo[1]'
    assert entry['Line'] == 1


def test_flag_0_boolean_violation_has_no_backing_node(checker):
    value = make_value("//foo = 'bar'", NODESET_ROOT, '0')
    violations = make_violations()

    checker._check_object_flag_0('schema', violations, NODESET_ROOT, value)

    entry = violations['dummy_brex']['0'][0]
    assert entry['Line'] == "(Boolean condition -> Interpret XPath)"
    assert entry['NodeXpath'] is None
    assert entry['Object'] is None


@pytest.mark.parametrize("xpath", ["string(//foo)", "count(//foo)"], ids=["string", "number"])
def test_flag_0_scalar_violation_has_no_backing_node(checker, xpath):
    # Ref brex_checker_rework.md §4.5: before this fixture surfaced it, a
    # bare numeric objectPath crashed with TypeError (int has no len()) and
    # a bare string objectPath was silently miscounted by iterating over
    # its characters instead of treating the whole string as one result.
    value = make_value(xpath, NODESET_ROOT, '0')
    violations = make_violations()

    checker._check_object_flag_0('schema', violations, NODESET_ROOT, value)

    entry = violations['dummy_brex']['0'][0]
    assert entry['Line'] == "(Scalar condition -> Interpret XPath)"
    assert entry['NodeXpath'] is None
    assert entry['Object'] is None


# ---------------------------------------------------------------------------
# Flag 1: "must be present / condition must hold"
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("xpath,root,violated", [
    ("//foo", NODESET_ROOT, False),          # node-set, non-empty -> satisfied
    ("//missing", NODESET_ROOT, True),       # node-set, empty -> violation
    ("//foo = 'bar'", NODESET_ROOT, False),  # boolean true -> satisfied
    ("//foo = 'zzz'", NODESET_ROOT, True),   # boolean false -> violation
    ("string(//foo)", NODESET_ROOT, False),      # non-empty string -> satisfied
    ("string(//missing)", NODESET_ROOT, True),   # empty string -> violation
    ("count(//foo)", NODESET_ROOT, False),       # non-zero number -> satisfied
    ("count(//missing)", NODESET_ROOT, True),    # zero -> violation
], ids=[
    "nodeset-present", "nodeset-missing",
    "boolean-true", "boolean-false",
    "string-nonempty", "string-empty",
    "number-nonzero", "number-zero",
])
def test_flag_1_result_types(checker, xpath, root, violated):
    value = make_value(xpath, root, '1')
    violations = make_violations()

    checker._check_object_flag_1('schema', violations, root, value)

    assert len(violations['dummy_brex']['1']) == (1 if violated else 0)


# ---------------------------------------------------------------------------
# Flag 2: "present, value constrained" -- a scalar function result is itself
# the value to check (no separate node-set of matched elements exists).
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("xpath,allowed,violated", [
    ("//foo", ["bar"], False),          # node-set element, matching text -> ok
    ("//foo", ["nope"], True),          # node-set element, wrong text -> violation
    ("string(//foo)", ["bar"], False),  # scalar string, matching -> ok
    ("string(//foo)", ["nope"], True),  # scalar string, wrong -> violation
    ("count(//foo)", ["1"], False),     # scalar number, matching (stringified) -> ok
    ("count(//foo)", ["9"], True),      # scalar number, wrong -> violation
], ids=[
    "nodeset-matches", "nodeset-mismatches",
    "string-matches", "string-mismatches",
    "number-matches", "number-mismatches",
])
def test_flag_2_result_types(checker, xpath, allowed, violated):
    value = make_value(xpath, NODESET_ROOT, '2', values_allowed=allowed)
    violations = make_violations()

    checker._check_object_flag_2('schema', violations, NODESET_ROOT, value)

    assert len(violations['dummy_brex']['2']) == (1 if violated else 0)


def test_flag_2_boolean_result_is_never_value_checked(checker):
    # A boolean-returning objectPath under flag 2 has no "value" to compare
    # against objectValue -- it can only ever report a match, never a value
    # violation (mirrors the existing type(result) is not bool guard).
    value = make_value("//foo = 'bar'", NODESET_ROOT, '2', values_allowed=["irrelevant"])
    violations = make_violations()

    checker._check_object_flag_2('schema', violations, NODESET_ROOT, value)

    assert violations['dummy_brex']['2'] == []


def test_flag_2_scalar_string_result_reports_the_whole_value_not_characters(checker):
    value = make_value("string(//foo)", NODESET_ROOT, '2', values_allowed=["nope"])
    violations = make_violations()

    checker._check_object_flag_2('schema', violations, NODESET_ROOT, value)

    # Before the fix, a 3-character string produced 3 bogus per-character
    # violations ("b", "a", "r") instead of one violation against "bar".
    assert len(violations['dummy_brex']['2']) == 1
    assert "bar" in violations['dummy_brex']['2'][0]['Description']
