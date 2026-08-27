import pytest
import regex

from acd.xml_processing import translate_xsd_regex_to_python


@pytest.mark.parametrize("xsd_pattern,value,expected", [
    (r"00[A-Z-[IO]]{1,3}|00[0]{1}", "00ABC", True),
    (r"00[A-Z-[IO]]{1,3}|00[0]{1}", "00IOI", False),
    (r"[0-9]{2}", "XX12XX", False),
    (r"[0-9]{2}", "12", True),
], ids=["subtraction-accepts", "subtraction-rejects", "whole-value-rejects-extra-chars", "whole-value-accepts"])
def test_translated_pattern_fullmatch(xsd_pattern, value, expected):
    translated = translate_xsd_regex_to_python(xsd_pattern)
    assert bool(regex.fullmatch(translated, value, regex.V1)) is expected


def test_character_class_subtraction_uses_double_dash():
    assert translate_xsd_regex_to_python(r"[A-Z-[IO]]") == "[A-Z--[IO]]"


@pytest.mark.parametrize("xsd_pattern,value,expected", [
    (r"\i\c*", "name1", True),
    (r"\i\c*", "1name", False),
])
def test_name_character_escapes_are_expanded(xsd_pattern, value, expected):
    translated = translate_xsd_regex_to_python(xsd_pattern)
    assert "\\i" not in translated and "\\c" not in translated
    assert bool(regex.fullmatch(translated, value, regex.V1)) is expected


def test_is_block_escape_is_translated_to_in_prefix():
    translated = translate_xsd_regex_to_python(r"\p{IsBasicLatin}+")
    assert translated == r"\p{In_BasicLatin}+"
    assert bool(regex.fullmatch(translated, "abc", regex.V1))
