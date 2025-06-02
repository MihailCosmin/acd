from datetime import datetime

import sys

from os import listdir
from os.path import join
from os.path import expanduser
from os.path import dirname
from os.path import basename
from os.path import isfile
from os.path import isdir

# from re import search  # To be replaced by regex.search, see below

from io import StringIO
from json import dump
from json import dumps

import elementpath

from regex import search
from re import findall
from regex import V1

from lxml import etree
from tqdm import tqdm

from os import environ
from os import system

from saxonche import PySaxonProcessor
from saxonche import PyXdmNode

from .xml_processing import get_schema_from_xml
from .xml_processing import delete_first_line
from .s1000d import get_brex_ref
from .s1000d import ref_dict_to_str
from .s1000d import find_document_by_reference


NS_DICT = {'rdf': r'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
            'xsi': r'http://www.w3.org/2001/XMLSchema-instance'}

class BrexNotFound(Exception):
    pass

class NoBrexDefined(Exception):
    pass

def clean_xpath(xpath):
    """Clean the xpath extra tabs, spaces and new lines"""
    xpath = xpath.strip().replace("\n", "").replace("\t", "")
    while "  " in xpath:
        xpath = xpath.replace("  ", " ")
    return xpath

class BrexChecker():
    def __init__(self, saxon: bool = False):
        """_summary_

        NOTE: enable optional paramater for saxon
        Args:
            saxon (bool, optional): _description_. Defaults to False.
        """
        self._xml_path = None
        self._xml_content = None
        self._xml_dir = None
        self._saxon = saxon

        self._brex_list = (None, None)
        self._brex_dir_path = (None, None)
    
    def set_xml_dir(self, dir_path: str) -> None:
        """_summary_

        Args:
            dir_path (str): _description_
        """
        self._xml_dir = dir_path

    def set_xml(self, xml: str):
        """Function with which the user can set the xml to be checked
        Args:
            xml (str): xml file path
        """
        with open(xml, "r", encoding="utf-8") as f:
            self._xml_content = f.read()
        self._xml_path = xml
        if self._brex_dir_path[0] is None and self._brex_dir_path[1] is not True:
            self._brex_dir_path = (dirname(xml), False)

    def _init_brex_list(self):
        if self._brex_list[0] is None and self._brex_list[1] in (None, True):
            self._brex_list = ([], True)
            xml = self._xml_path
            while True:
                if ref_dict_to_str(get_brex_ref(xml)) in xml:
                    break
                brex_ref = ref_dict_to_str(get_brex_ref(xml))
                xml = find_document_by_reference(brex_ref, self._brex_dir_path[0])
                if xml is None:
                    break
                self._brex_list[0].append(xml)
        elif self._brex_list[0] is None and self._brex_list[1] is not True:
            self._brex_list = ([], False)
            xml = self._xml_path
            while True:
                if ref_dict_to_str(get_brex_ref(xml)) in xml:
                    break
                brex_ref = ref_dict_to_str(get_brex_ref(xml))
                xml = find_document_by_reference(brex_ref, self._brex_dir_path[0])
                if xml is None:
                    break
                self._brex_list[0].append(xml)
        if len(self._brex_list[0]) == 0:
            raise NoBrexDefined(f"Brex files couldn't be found\n\
                    Please use set_brex_path method to input the directory containing ALL brex data modules or \
                    use override_brex_list if the brex data modules are in different directories.\
                    expected brex: {ref_dict_to_str(get_brex_ref(self._xml_path))}".replace("                ", ""))
        else:
            for brex in self._brex_list[0]:
                if not isfile(brex):
                    raise BrexNotFound(f"Referenced Brex: {brex} is not in {self._brex_dir_path[0]}.\n\
                    Please use set_brex_path method to input the directory containing ALL brex data modules or \
                    use override_brex_list if the brex data modules are in different directories.".replace("                ", ""))

    def set_brex_path(self, brex_path: str):
        """Function with which the user can set a path where the brex files are
        located in case they are located in another directory than the xml.
        Function call can be omitted when the Brex has the same directory, the xml has.
        Args:
            brex (str): brex file path
        """
        if isdir(brex_path):
            self._brex_dir_path = (brex_path, True)
        else:
            raise BrexNotFound(f"The given path {brex_path} seems to be leading to a file. \
                Please make sure to input the path of the directory containing ALL brex data modules or \
                use override_brex_list if the brex data modules are in different directories.".replace("                ", ""))

    def override_brex_list(self, _brex_list: list):
        """The user can specify a list with specific paths and brex files

        Args:
            _brex_list (list): list of strings containing paths of different brex paths
        """
        for brex_elem in _brex_list:
            if isfile(brex_elem) is False:
                raise BrexNotFound(f"Brex could not be found in given directory {brex_elem}. \
                                     Please specify the absolute path.")
        self._brex_list = (_brex_list, True)

    def _get_object_rule_nodes(self, brex: str) -> any:
        """Return all nodes in a set matching an XPath expression

        Args:
            xpath (str): xpath expression
            brex (str): path of the brex

        Returns:
            any: Set of nodes
        """
        with open(brex, "r", encoding="utf-8") as _:
            brex_content = _.read()
        brex_content = delete_first_line(brex_content)
        brex_content = etree.parse(StringIO(brex_content))
        brothers = brex_content.findall('//structureObjectRuleGroup/structureObjectRule/objectPath')
        return brothers

    def _show_rules(self, brex: str, debug: bool = False) -> any:
        """Creates a, in nested dictionaries structured, JSON file containing all necessary information about the brex rules i.e.
        xpath, objectflag, objectUse, objectValues et Al.

        Args:
            brex (str): brex_path

        Returns:
            any: Nested Dictionary
        """
        nodes_to_check = self._get_object_rule_nodes(brex)       
        allowed_object_flag_dict = []
        for counter, x in enumerate(nodes_to_check):
            values_allowed = []
            regex_allowed = []
            for objectValue in x.getparent().xpath('objectValue'):
                for key, value in objectValue.attrib.items():
                    if key == "valueForm" and value == "single":
                        values_allowed.append(objectValue.attrib["valueAllowed"])
                        break
                    elif key == "valueForm" and value == "pattern":
                        regex_allowed.append(objectValue.attrib["valueAllowed"])
                        break
                    elif key == "valueForm" and value == "range":
                        values_range = objectValue.attrib["valueAllowed"]
                        values_range = findall(r"([a-z]*)(\d+)", values_range)
                        if len(values_range[0]) == 2:
                            name = values_range[0][0]
                            list_start = int(values_range[0][1])
                            list_end = int(values_range[1][1])
                        else:
                            name = ""
                            list_start = int(values_range[0])
                            list_end = int(values_range[1])
                        range_list = []
                        while(list_start <= list_end):
                            range_list.append(f"{name}{str(list_start)}")
                            list_start += 1
                        values_allowed += range_list
                        break
            try:
                context_rules = x.getparent().getparent().getparent().attrib['rulesContext']
            except KeyError:
                context_rules = ""
            allowed_object_flag_dict.append({
                    'xpath': str(nodes_to_check[counter].text),
                    'Brex': str(brex),
                    'ObjectFlag': str(x.attrib['allowedObjectFlag']),
                    'objectUse': str(x.getparent().xpath('objectUse')[0].text),
                    'contextRules': context_rules,
                    'values_allowed': values_allowed,
                    'regex_allowed': regex_allowed
                }
            )
        if debug:
            with open(join(expanduser("~/Desktop"), f'brex_{basename(brex)}.json'), 'w', encoding="utf-8") as _:
                for elem in allowed_object_flag_dict:
                    _.write(dumps(elem, indent=4, ensure_ascii=False))
        return allowed_object_flag_dict

    def regex_builder(self, attribute_name: str, attribute_value: str, xpath):
        """If case since there might be cases where attribute_name has no attribute_value
        Args:
            attribute_name (str): _description_
            attribute_value (str): _description_
        Returns:
            _type_: _description_
        """
        if attribute_value is not None:
            build_regex = f'({attribute_name})(.*?)("{attribute_value}")'
        else:
            build_regex = f'({attribute_name})(.*?)(")(.*?)(")'
        return build_regex

    def _check_object_flag_0(self, schema: str, brex_violations: dict, root: any, value: any, error_0: int):
        if value['contextRules'] == schema or value['contextRules'] == "":
            if self._saxon:
                with PySaxonProcessor(license=False) as proc:
                    xp = proc.new_xpath_processor()
                    for prefix, uri in NS_DICT.items():
                        xp.declare_namespace(prefix, uri)
                    node = proc.parse_xml(xml_file_name=self._xml_path)
                    xp.set_context(xdm_item=node)
                    items = xp.evaluate(clean_xpath(value['xpath']))
                    if items is not None:
                        for item in items:
                            if isinstance(item, PyXdmNode):
                                match_found = search(r'(\[@)(.+?)([^a-z0-9A-Z])', clean_xpath(value['xpath']))
                                if match_found:
                                    attribute_name = match_found.group(2)
                                    attribute_value = item.get_attribute_value(attribute_name)
                                else:
                                    attribute_name = ""
                                    attribute_value = ""
                                list_xml_content = self._xml_content.split("\n")
                                build_regex = self.regex_builder(attribute_name, attribute_value, clean_xpath(value['xpath']))
                                for element in list_xml_content:
                                    match_found_in_list = search(build_regex, element)
                                    if match_found_in_list:
                                        brex_violations[value["Brex"]]['0'] |= {error_0: {
                                            'Line': list_xml_content.index(element) + 1,
                                            'Description': value["objectUse"],
                                            'Xpath': value['xpath']}
                                        }
                                        error_0 += 1
                    proc.exception_clear()
            else:
                selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
                if type(selector.select(root)) is not bool:
                    for element in selector.select(root):
                        if ' and ' in value['xpath']:
                            line_no = "(Origin traced back to multiple lines -> Interpret XPath)"
                        else:
                            try:
                                line_no = element.sourceline
                            except AttributeError:
                                if search(r'(/@)([a-zA-Z]+)', value['xpath'], V1):
                                    attrib_name = search(r'(/@)([a-zA-Z]+)', value['xpath'], V1).group(2)
                                    split_xml = self._xml_content.split("\n")
                                    for ind, elem in enumerate(split_xml):
                                        if attrib_name in elem:
                                            line_no = ind + 1
                                else:
                                    line_no = "x"
                        brex_violations[value["Brex"]]['0'] |= {error_0: {
                            'Line': line_no,  # was line_no
                            'Description': value["objectUse"],
                            'Xpath': value['xpath']}
                        }
                        error_0 += 1
        return brex_violations

    def _check_object_flag_1(self, schema: str, brex_violations: dict, root: any, value: any, error_1: int):
        if value['contextRules'] == schema or value['contextRules'] == "":
            selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
            if not selector.select(root) or selector.select(root) == []:
                brex_violations[value["Brex"]]['1'] |= {error_1: {
                            'Description': value["objectUse"],
                            'Xpath': value['xpath']}
                            }
                error_1 += 1
        return brex_violations

    def _check_object_flag_2(self, schema: str, brex_violations: dict, root: any, value: any, error_2: int):
        if ('values_allowed' in value or 'regex_allowed' in value) and (value['contextRules'] == schema or value['contextRules'] == ""):
            selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
            if type(selector.select(root)) is not bool:
                for element in selector.select(root):
                    valid_elem = False
                    if element not in value["values_allowed"]:
                        if len(value["regex_allowed"]) > 0:
                            try:
                                if any([bool(search(regex, element, V1)) for regex in value["regex_allowed"]]):
                                    valid_elem = True
                            except TypeError:
                                regex2 = search(r"(@)([a-zA-Z]+)(^[a-zA-Z])", value['xpath'], V1)
                                if any([bool(search(regex, element.attrib[regex2.group(2)], V1)) for regex in value["regex_allowed"]]):
                                    valid_elem = True
                    else:
                        valid_elem = True
                    if not valid_elem:
                        if (r'] and ' or r'and \[' or r'] and \[' or r'\) and' or r'and \(' or r'\) and \(') in value['xpath']:
                            line_no = "(Origin traced back to multiple lines -> Read XPath)"
                        else:
                            try:
                                line_no = element.sourceline
                            except AttributeError:
                                if search(r'(/@)([a-zA-Z]+)', value['xpath'], V1) is not None:
                                    attrib_name = search(r'(/@)([a-zA-Z]+)', value['xpath'], V1).group(2)
                                    split_xml = self._xml_content.split("\n")
                                    for ind, elem in enumerate(split_xml):
                                        if attrib_name in elem and element in elem:
                                            line_no = ind + 1
                                else:
                                    line_no = "x"
                        brex_violations[value["Brex"]]['2'] |= {error_2: {
                            'Line': line_no,
                            'Description': f'Element/Attribute ({element}) did not match the object values.',
                            'Xpath': value['xpath'],
                            'Single Values': [value["values_allowed"]],
                            'Pattern Values': [value["regex_allowed"]],
                            'ObjectUse': value["objectUse"]}}
                        error_2 += 1
        return brex_violations

    def _check_rules(self, debug: bool = False, include_tqdm: bool = False) -> dict:
        """Traverses through every node of the brex and checks the rules through the given xpaths.
        For objectFlag 0 we also get the line of the error
        For objectFlag 1 we only get the Description of the rule that was violated
        For objectFlag 2 we get a list containing all 'single' values and a list containing all 'pattern' values
                         and we might get the line of the error

        Returns:
            any: Dictionary with all errors
        """
        with open(self._xml_path, "r", encoding="utf-8") as _:
            self.xml_content = _.read()
        schema = get_schema_from_xml(self.xml_content)
        brex_violations_dict = {}
        for brex in self._brex_list[0]:
            brex_violations_dict[brex] = {
                '0': {},
                '1': {},
                '2': {}
            }
        root = etree.parse(self._xml_path)
        all_content_rules = []
        for brex in self._brex_list[0]:
            content_rules = self._show_rules(brex, debug=debug)
            all_content_rules += content_rules

        if debug:
            with open(join(expanduser("~/Desktop"), "All_content_rules.txt"), 'w', encoding="utf-8") as _:
                for rule in all_content_rules:
                    _.write(str(rule) + "\n")
        error_0, error_1, error_2 = 1, 1, 1
        container = tqdm(all_content_rules) if include_tqdm else all_content_rules
        for value in container:
            if value["ObjectFlag"] == '0':
                if schema != "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/ddn.xsd":
                    brex_violations_dict |= self._check_object_flag_0(schema, brex_violations_dict, root, value, error_0)
            if value["ObjectFlag"] == '1':
                if schema != "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/ddn.xsd":    
                    brex_violations_dict |= self._check_object_flag_1(schema, brex_violations_dict, root, value, error_1)
            if value["ObjectFlag"] == '2':
                if value["values_allowed"] != [] or value["regex_allowed"] != []:
                    if schema != "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/ddn.xsd":
                        brex_violations_dict |= self._check_object_flag_2(schema, brex_violations_dict, root, value, error_2)
        return brex_violations_dict

    def _append_summary(self, object_flag_dict: dict) -> dict:
        """Counts the number of Brex Errors for a xml and appends this
        information as another key-value-pair into the dictionary

        Args:
            object_flag_dict (dict): _description_

        Returns:
            dict: _description_
        """
        error_count = 0
        for x, values in object_flag_dict.items():
            values_length = 0
            for value in values.values():
                values_length += len(value)
            error_count += values_length
        return f"{error_count} Errors"
    
    def validate(self, debug: bool = False, include_tqdm: bool = False) -> dict:
        """Check xml against all brexes and dump the results into a JSon file
        """
        if self._xml_dir:
            if debug:
                with open(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_dir)}.json'), 'w', encoding="utf-8") as _:
                    _.write("{")
            files = [_ for _ in listdir(self._xml_dir) if ".xml" in _.lower() and "-022a-" not in _.lower() ]
            container = tqdm(files) if include_tqdm else files
            for _xml in container:
                self.set_xml(join(self._xml_dir, _xml))
                self._init_brex_list()
                result = self._check_rules(debug=debug, include_tqdm=include_tqdm)
                summary = self._append_summary(result)
                if debug:
                    with open(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_dir)}.json'), 'a', encoding="utf-8") as _:
                        dump({_xml: result, "Summary": summary}, _, indent=4)
                self._brex_list = (None, None)
                self._brex_dir_path = (None, None)
            if debug:
                with open(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_dir)}.json'), 'a', encoding="utf-8") as _:
                    _.write("}")
        else:
            self._init_brex_list()
            result = self._check_rules(debug=debug)
            summary = self._append_summary(result)
            result["Summary"] = summary
            if debug:
                with open(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_path)}.json'), 'w', encoding="utf-8") as _:
                    dump(result, _, indent=4)
        return result

