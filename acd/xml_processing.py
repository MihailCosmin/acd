from io import BytesIO

from os.path import isfile

from re import search
from re import sub

from urllib.parse import quote

from lxml import etree
from .filepath import clean_path

XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance"

# XML 1.0 Name production, per the XSD regex spec (Appendix G):
# NameStartChar := ":" | [A-Z] | "_" | [a-z] | [#xC0-#xD6] | ...
_XSD_NAME_START_CHAR = (
    r":A-Z_a-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D"
    r"\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF"
    r"\uF900-\uFDCF\uFDF0-\uFFFD\U00010000-\U000EFFFF"
)
# NameChar := NameStartChar | "-" | "." | [0-9] | #xB7 | ...
_XSD_NAME_CHAR = _XSD_NAME_START_CHAR + r"\-.0-9\u00B7\u0300-\u036F\u203F-\u2040"

_XSD_MULTICHAR_ESCAPES = {
    "i": f"[{_XSD_NAME_START_CHAR}]",
    "I": f"[^{_XSD_NAME_START_CHAR}]",
    "c": f"[{_XSD_NAME_CHAR}]",
    "C": f"[^{_XSD_NAME_CHAR}]",
}

def translate_xsd_regex_to_python(pattern: str) -> str:
    """
    Translates an XSD (xsd:pattern / valueAllowed) regular expression into
    the syntax understood by the `regex` module compiled with the `V1` flag.

    Handles the constructs that differ between the two dialects:
    - Character-class subtraction, `[A-Z-[IO]]`, which XSD writes with a
      single `-` but `regex` V1 set operations require doubled, `[A-Z--[IO]]`.
    - The `\\i` / `\\I` / `\\c` / `\\C` XML-name-character escapes, which
      `regex` does not know, expanded into their explicit Unicode classes.
    - Block escapes, `\\p{IsBasicLatin}` / `\\P{IsBasicLatin}`, whose `Is`
      block-name prefix `regex` spells `In_`.

    Args:
        pattern (str): The XSD regular expression as read from `valueAllowed`

    Returns:
        str: An equivalent pattern that `regex.fullmatch(pattern, value, regex.V1)`
        can compile

    """

    pattern = sub(r'\\([pP])\{Is', r'\\\1{In_', pattern)

    out = []
    depth = 0
    i = 0
    length = len(pattern)
    while i < length:
        char = pattern[i]
        if char == "\\" and i + 1 < length:
            nxt = pattern[i + 1]
            replacement = _XSD_MULTICHAR_ESCAPES.get(nxt)
            out.append(replacement if replacement is not None else char + nxt)
            i += 2
            continue
        if char == "[":
            depth += 1
            out.append(char)
            i += 1
            continue
        if char == "]":
            depth = max(0, depth - 1)
            out.append(char)
            i += 1
            continue
        if char == "-" and depth > 0 and i + 1 < length and pattern[i + 1] == "[":
            out.append("--")
            i += 1
            continue
        out.append(char)
        i += 1
    return "".join(out)

def _as_float(value: str):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

def is_in_range(value: str, value_range: str) -> bool:
    """
    Tests whether `value` falls in an S1000D `objectValue` range (`first~last`),
    port of `s1kd_tools.c` `is_in_range` (`tools/common/s1kd_tools.c:378-407`).

    A range with no `~` is a single literal, matched by exact equality. Otherwise
    the range is split on `~` (only the first two tokens are used, matching the
    reference `strtok` behaviour), and the bounds are compared against `value`
    numerically when `value` and both bounds parse as numbers, lexicographically
    otherwise (e.g. `20~100` must be numeric, since `100` sorts before `20`
    lexicographically).

    Args:
        value (str): The value being checked
        value_range (str): A single `first~last` range, or a plain literal

    Returns:
        bool: True if `value` is in the range (or equals the literal)

    """

    if "~" not in value_range:
        return value == value_range

    first, last = value_range.split("~")[:2]
    f, l, v = _as_float(first), _as_float(last), _as_float(value)
    if f is not None and l is not None and v is not None:
        return f <= v <= l
    return first <= value <= last

def is_in_set(value: str, value_set: str) -> bool:
    """
    Tests whether `value` is in an S1000D `objectValue` set (`a|b|c`, where each
    member may itself be a range), port of `s1kd_tools.c` `is_in_set`
    (`tools/common/s1kd_tools.c:409-434`).

    Args:
        value (str): The value being checked
        value_set (str): The `valueAllowed` of a `range`-form `objectValue`,
            e.g. `a~c`, `A|B|C`, `01|02`, `aa01~aa09`

    Returns:
        bool: True if `value` matches any member of the set

    """

    if "|" not in value_set:
        return is_in_range(value, value_set)
    return any(is_in_range(value, member) for member in value_set.split("|"))

def delete_first_line(xml_content: str, overwrite: bool = False) -> str:
    """
    If the first line of the schema matches the regular expression, it is removed

    Args:
        file_to_delete_line (str): Content of the schema (before linearization)

    Returns:
        str: String containing the content of the schema,
        but with the changes mentioned in the description

    """
    xml_filename = None
    if isfile(xml_content):
        xml_filename = xml_content
        with open(clean_path(xml_content), "r", encoding="utf-8") as _:
            xml_content = _.read()

    if search(r'(<\?)(.*?)(\?>)', xml_content):
        return sub(r'(<\?)(.*?)(\?>)', "", xml_content)
    if "encoding" in xml_content.splitlines()[1:]:  # We need to remove first line only if encoding is specified in the XML
        return "\n".join(xml_content.splitlines()[1:])

    if xml_filename is not None:
        extension = "." + xml_filename.split(".")[-1]
        if not overwrite:
            xml_filename = xml_filename.replace(extension, f"_fixed{extension}")
        with open(clean_path(xml_filename), "w", encoding="utf-8") as _:
            _.write(xml_content)
    return xml_content

def linearize_xml(xml_content: str) -> str:
    """
    Linearizes a given xml document (writes content into a single line),
    removes all tab characters,
    removes every area where two or more than
    two white spaces appear after each other,
    removes every carriage return character (\r) and
    removes every white space between > and < characters

    Args:
        file_to_linearize (str): Content of the xml document

    Returns:
        str: String containing the content of the xml document,
        but with the changes mentioned in the description

    """

    xml_content = sub('[\n\r\t]+', " ", xml_content)
    xml_content = sub(r' {2,}', " ", xml_content)
    xml_content = xml_content.replace("> <", "><")
    return xml_content

def get_schema_from_xml(xml_content: str) -> str:
    """
    Reads the root element's xsi:noNamespaceSchemaLocation attribute via a
    namespace-aware parsed lookup (equivalent to the XPath
    /*/@xsi:noNamespaceSchemaLocation), so the result no longer depends on
    where that attribute sits among the root element's attributes.

    Args:
        xml_content (str): Content of the xml document (linearized or not)

    Returns:
        str: A string that contains the schema location, or None if the
        root element does not declare one

    """

    content = xml_content.encode("utf-8") if isinstance(xml_content, str) else xml_content
    try:
        for _, root in etree.iterparse(BytesIO(content), events=("start",), recover=True, huge_tree=True):
            values = root.xpath('/*/@xsi:noNamespaceSchemaLocation', namespaces={'xsi': XSI_NAMESPACE})
            return str(values[0]) if values else None
    except etree.XMLSyntaxError:
        return None
    return None

def get_xml_attribute(xml: str, xpath: str, attribute: str) -> str:
    """
    Gets an attribute of an XML element

    Args:
        xml (str): XML path
        xpath (str): XPath to the element
        attribute (str): Attribute to get

    Returns:
        str: Attribute value

    """

    return etree.parse(xml).xpath(xpath)[0].attrib[attribute]

def set_xml_attribute(xml: str, xpath: str, attribute: str, value: str) -> None:
    """
    Sets an attribute of an XML element

    Args:
        xml (str): XML path
        xpath (str): XPath to the element
        attribute (str): Attribute to set
        value (str): Value to set the attribute to

    """

    xml_tree = etree.parse(xml)

    xml_tree.xpath(xpath)[0].attrib[attribute] = value

    with open(clean_path(xml), "w", encoding="utf-8") as _:
        _.write(etree.tostring(xml_tree, pretty_print=True).decode("utf-8"))

def get_xml_tag_content(xml: str, xpath: str) -> str:
    """
    Gets the content of an XML tag

    Args:
        xml (str): XML path
        xpath (str): XPath to the element

    Returns:
        str: Tag content

    """

    return etree.parse(xml).xpath(xpath)[0].text

def set_xml_tag_content(xml: str, xpath: str, content: str) -> None:
    """
    Sets the content of an XML tag

    Args:
        xml (str): XML path
        xpath (str): XPath to the element
        content (str): Content to set the tag to

    """

    xml_tree = etree.parse(xml)

    xml_tree.xpath(xpath)[0].text = content

    with open(clean_path(xml), "w", encoding="utf-8") as _:
        _.write(etree.tostring(xml_tree, pretty_print=True).decode("utf-8"))

def sanitize_entity_system_uris(xml_content: str) -> str:
    """
    Percent-encodes the SYSTEM literal of <!ENTITY ... SYSTEM "..."> declarations
    (e.g. graphic/CGM file names used as NDATA entities). CGM file names are
    frequently authored with spaces, parentheses, etc., which are not valid
    unescaped characters in a URI reference. libxml2 validates the SYSTEM
    literal as a URI while parsing the DTD internal subset and raises
    "Invalid URI" for these otherwise well-formed documents, so the literal
    needs to be percent-encoded before parsing. The entity name itself
    (used elsewhere in the document) is left untouched.

    Args:
        xml_content (str): Content of the xml document

    Returns:
        str: xml content with SYSTEM literals percent-encoded
    """

    entity_system_regex = r'(<!ENTITY\s+\S+\s+SYSTEM\s+")([^"]*)(")'

    def encode_system_literal(match):
        prefix, literal, suffix = match.groups()
        return prefix + quote(literal, safe="/:._-") + suffix

    return sub(entity_system_regex, encode_system_literal, xml_content)

def replace_special_characters(xml_content: str):
    """
    replace_special_characters

    """

    xml_content = xml_content.replace("\u00a0", " ")
    xml_content = xml_content.replace("&nbsp;", " ")
    xml_content = xml_content.replace("&#xa0;", " ")
    xml_content = xml_content.replace("\xa0", " ")
    xml_content = xml_content.replace("&#160;", " ")
    xml_content = xml_content.replace(" ", " ")
    xml_content = xml_content.replace("&#177;", "±")
    xml_content = xml_content.replace("\u00b1", "±")
    xml_content = xml_content.replace("&#xb1;", "±")
    xml_content = xml_content.replace("&plusmn;", "±")
    xml_content = sanitize_entity_system_uris(xml_content)

    return xml_content
