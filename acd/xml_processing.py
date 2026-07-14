from os.path import isfile

from re import search
from re import sub

from urllib.parse import quote

from lxml import etree
from .filepath import clean_path

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

def get_schema_from_xml(linearized_file: str) -> str:
    """
    Searches if a schema url is given inside the xml document.
    If a url matches the regular expression,
    the url is returned.

    Args:
        linearized_file (str): Linearized content of the xml document,

    Returns:
        tuple: A string that contains the schema name or None

    """

    schema_url_regex = r'(xsi:noNamespaceSchemaLocation=")(.*?)(">)'
    if search(schema_url_regex, linearized_file):
        schema = str(search(schema_url_regex, linearized_file).group(2))
        return schema
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
