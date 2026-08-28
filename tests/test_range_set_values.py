import pytest

from acd.xml_processing import is_in_range
from acd.xml_processing import is_in_set


@pytest.mark.parametrize("value,value_range,expected", [
    ("a", "a~c", True),
    ("b", "a~c", True),
    ("c", "a~c", True),
    ("d", "a~c", False),
    ("aa05", "aa01~aa09", True),
    ("aa10", "aa01~aa09", False),
    ("0050", "0001~0099", True),
    ("0100", "0001~0099", False),
    ("50", "20~100", True),   # numeric: fails lexicographically ("50" > "100" as strings)
    ("100", "20~100", True),
    ("10", "20~100", False),
    ("accpnl67", "accpnl51~accpnl99", True),
    ("accpnl05", "accpnl51~accpnl99", False),
], ids=[
    "letter-lower-bound", "letter-mid", "letter-upper-bound", "letter-out-of-range",
    "zero-padded-alnum-in-range", "zero-padded-alnum-out-of-range",
    "numeric-zero-padded-in-range", "numeric-zero-padded-out-of-range",
    "numeric-mid-in-range", "numeric-upper-bound-in-range", "numeric-below-range",
    "prefixed-numeric-in-range", "prefixed-numeric-out-of-range",
])
def test_is_in_range(value, value_range, expected):
    assert is_in_range(value, value_range) is expected


@pytest.mark.parametrize("value,value_set,expected", [
    ("A", "A|B|C", True),
    ("D", "A|B|C", False),
    ("01", "01|02", True),
    ("03", "01|02", False),
    ("b", "a~c", True),
    ("z", "a~c", False),
], ids=[
    "member-of-set", "not-in-set",
    "literal-set-member", "literal-set-non-member",
    "single-range-no-pipe", "single-range-no-pipe-miss",
])
def test_is_in_set(value, value_set, expected):
    assert is_in_set(value, value_set) is expected
