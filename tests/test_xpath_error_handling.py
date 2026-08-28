import pytest
from lxml import etree

from acd.brex_checker import BrexChecker

BAD_XPATH = "][invalid xpath("


def make_violations(brex="dummy_brex"):
    return {brex: {'0': [], '1': [], '2': [], 'xpathError': []}}


def make_value(**overrides):
    value = {
        'xpath': BAD_XPATH,
        'Brex': 'dummy_brex',
        'ObjectFlag': '0',
        'objectUse': 'Some rule description',
        'contextRules': '',
        'values_allowed': [],
        'regex_allowed': [],
        'ranges_allowed': [],
    }
    value.update(overrides)
    return value


@pytest.fixture
def checker():
    return BrexChecker()


@pytest.fixture
def root():
    return etree.fromstring(b"<root><child/></root>")


def test_flag_0_records_xpath_error_and_does_not_raise(checker, root):
    value = make_value()
    violations = checker._check_object_flag_0("schema", make_violations(), root, value)
    assert len(violations['dummy_brex']['xpathError']) == 1
    error_entry = violations['dummy_brex']['xpathError'][0]
    assert error_entry['Xpath'] == BAD_XPATH
    assert error_entry['Description'] == value['objectUse']
    assert violations['dummy_brex']['0'] == []


def test_flag_1_records_xpath_error_and_does_not_raise(checker, root):
    value = make_value(ObjectFlag='1')
    violations = checker._check_object_flag_1("schema", make_violations(), root, value)
    assert len(violations['dummy_brex']['xpathError']) == 1
    assert violations['dummy_brex']['1'] == []


def test_flag_2_records_xpath_error_and_does_not_raise(checker, root):
    value = make_value(ObjectFlag='2', values_allowed=['A'])
    violations = checker._check_object_flag_2("schema", make_violations(), root, value)
    assert len(violations['dummy_brex']['xpathError']) == 1
    assert violations['dummy_brex']['2'] == []
