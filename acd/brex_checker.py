from datetime import datetime

from copy import deepcopy

import sys

from os import listdir
from os.path import join
from os.path import expanduser
from os.path import dirname
from os.path import basename
from os.path import isfile
from os.path import isdir
from os.path import abspath

from .filepath import clean_path

# from re import search  # To be replaced by regex.search, see below

from io import StringIO
from json import dump
from json import dumps

import elementpath

from regex import search
from regex import fullmatch
from regex import V1

from lxml import etree
from tqdm import tqdm

from os import environ
from os import system

from saxonche import PySaxonProcessor
from saxonche import PyXdmNode

from .xml_processing import get_schema_from_xml
from .xml_processing import delete_first_line
from .xml_processing import translate_xsd_regex_to_python
from .xml_processing import is_in_set
from .s1000d import get_brex_ref
from .s1000d import ref_dict_to_str
from .s1000d import find_document_by_reference
from .default_brex import default_brex_dmc
from .default_brex import default_brex_path
from .default_brex import find_default_brex_fallback


NS_DICT = {'rdf': r'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
            'xsi': r'http://www.w3.org/2001/XMLSchema-instance'}

SNS_MODES = ("normal", "strict", "unstrict")

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
        self._use_default_brex = False
        self._brex_fallbacks = []

        self._severity_levels_path = None
        self._severity_levels_search = True
        self._severity_levels = (None, False)

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
        with open(clean_path(xml), "r", encoding="utf-8") as f:
            self._xml_content = f.read()
        self._xml_path = xml
        if self._brex_dir_path[0] is None and self._brex_dir_path[1] is not True:
            self._brex_dir_path = (dirname(xml), False)

    def _init_brex_list(self):
        if self._brex_list[0] is not None:
            return

        if self._use_default_brex:
            self._brex_fallbacks = []
            schema = get_schema_from_xml(self._xml_content)
            self._brex_list = ([default_brex_path(default_brex_dmc(schema))], True)
            return

        self._brex_list = ([], True)
        xml = self._xml_path
        visited = {xml}
        fallbacks = []
        while True:
            brex_ref_dict = get_brex_ref(xml)
            if brex_ref_dict is None:
                break
            brex_ref = ref_dict_to_str(brex_ref_dict)
            if brex_ref in xml:
                break
            resolved = find_document_by_reference(brex_ref, self._brex_dir_path[0])
            if resolved is None:
                # The referenced BREX isn't on disk anywhere we looked: fall
                # back to the built-in default BREX if the reference names
                # one of them (search_brex_fname_from_default_brex).
                fallback_dmc = find_default_brex_fallback(brex_ref_dict)
                if fallback_dmc is None:
                    break
                resolved = default_brex_path(fallback_dmc)
                fallbacks.append({
                    'Reference': brex_ref,
                    'UsedBuiltinBrex': fallback_dmc,
                    'BuiltinBrexPath': resolved
                })
            if resolved in visited:
                break
            visited.add(resolved)
            self._brex_list[0].append(resolved)
            xml = resolved
        self._brex_fallbacks = fallbacks

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
        self._brex_fallbacks = []

    def use_default_brex(self, enabled: bool = True):
        """Equivalent to `s1kd-brexcheck -B`/`--default-brex`: ignore any
        `brexDmRef`/`brexref` the object carries and check only against the
        built-in default BREX matching the object's declared S1000D schema
        version (see `default_brex.default_brex_dmc`), with no further
        `brexDmRef` layering. Overrides `set_brex_path` / `override_brex_list`
        while enabled.

        Args:
            enabled (bool, optional): if True, switch to default-BREX-only
                mode; if False, return to normal `brexDmRef` resolution.
                Defaults to True.
        """
        self._use_default_brex = enabled
        self._brex_list = (None, None)
        self._brex_fallbacks = []

    def set_severity_levels_path(self, path: str):
        """Explicitly set the path to the `.brseveritylevels` file used to resolve
        `brSeverityLevel` values, e.g.:
        `<brSeverityLevels><brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>...`

        This overrides the default behaviour of searching the checked XML's
        directory and its parents for a file named `.brseveritylevels`
        (see `_find_severity_levels_file`); once called, that search is skipped
        and this exact path is used instead.

        Args:
            path (str): path to the severity levels XML file
        """
        self._severity_levels_path = path
        self._severity_levels = (None, False)

    def set_severity_levels_search(self, enabled: bool):
        """Enable or disable the default parent-directory search for a
        `.brseveritylevels` file (see `_find_severity_levels_file`). Enabled by
        default; this is the override for callers who need to turn it off, e.g.
        to ignore a `.brseveritylevels` file that exists but should not apply to
        this run. Has no effect once `set_severity_levels_path` has been called.

        Args:
            enabled (bool): if False, no severity-levels file is auto-discovered
                and every violation fails, as if no severity levels applied at all
        """
        self._severity_levels_search = enabled
        self._severity_levels = (None, False)

    def _find_severity_levels_file(self) -> str:
        """Search the checked XML's directory, then its parent directories in
        turn, for a file named `.brseveritylevels`. Adapts the generic
        `find_config` helper (`s1kd_tools.c:30-56`) that s1kd tools use to
        auto-discover CSDB-wide configuration files, searching from the checked
        file's directory instead of the process's current directory.

        Returns:
            str: full path to the first `.brseveritylevels` file found; None if
                `self._xml_path` is not set yet, or no file is found before the
                filesystem root is reached
        """
        if not self._xml_path:
            return None
        current = abspath(dirname(self._xml_path))
        while True:
            candidate = join(current, ".brseveritylevels")
            if isfile(candidate):
                return candidate
            parent = dirname(current)
            if parent == current:
                return None
            current = parent

    def _get_severity_levels(self) -> dict:
        """Parse the `.brseveritylevels` file into `{value: {'fail': bool, 'type': str}}`,
        caching the result. Port of the `brsl` lookup table in `s1kd-brexcheck.c`.

        Resolves the file from `set_severity_levels_path` if one was set explicitly;
        otherwise, unless disabled via `set_severity_levels_search`, auto-discovers
        it via `_find_severity_levels_file`.

        Returns:
            dict: severity-level lookup table; empty if no file was set or found
        """
        if self._severity_levels[1]:
            return self._severity_levels[0]
        levels = {}
        path = self._severity_levels_path
        if path is None and self._severity_levels_search:
            path = self._find_severity_levels_file()
        if path is not None:
            with open(clean_path(path), "r", encoding="utf-8") as _:
                content = _.read()
            tree = etree.parse(StringIO(content))
            for level in tree.findall('.//brSeverityLevel'):
                value = level.get('value')
                if value is None:
                    continue
                levels[value] = {
                    'fail': level.get('fail') != 'no',
                    'type': "".join(level.itertext()) or None
                }
        self._severity_levels = (levels, True)
        return levels

    def _is_severity_failure(self, severity: str) -> bool:
        """Decide whether a violation at the given business-rule severity level
        counts as a failure.

        Port of `is_failure` (`s1kd-brexcheck.c:569-605`): a violation always fails
        unless a `.brseveritylevels` file is set or auto-discovered, defines this
        exact severity value, and marks it `fail="no"`.

        Args:
            severity (str): resolved `brSeverityLevel` value, or None

        Returns:
            bool: True if the violation should count as a failing error
        """
        if severity is None:
            return True
        level = self._get_severity_levels().get(severity)
        if level is None:
            return True
        return level['fail']

    def _get_object_rule_nodes(self, brex: str) -> any:
        """Return all nodes in a set matching an XPath expression

        Args:
            xpath (str): xpath expression
            brex (str): path of the brex

        Returns:
            any: Set of nodes
        """
        with open(clean_path(brex), "r", encoding="utf-8") as _:
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
        default_br_severity_level = None
        if len(nodes_to_check) > 0:
            default_br_severity_level = nodes_to_check[0].getroottree().getroot().get('defaultBrSeverityLevel')
        allowed_object_flag_dict = []
        for counter, x in enumerate(nodes_to_check):
            values_allowed = []
            regex_allowed = []
            ranges_allowed = []
            for objectValue in x.getparent().xpath('objectValue'):
                for key, value in objectValue.attrib.items():
                    if key == "valueForm" and value == "single":
                        values_allowed.append(objectValue.attrib["valueAllowed"])
                        break
                    elif key == "valueForm" and value == "pattern":
                        regex_allowed.append(translate_xsd_regex_to_python(objectValue.attrib["valueAllowed"]))
                        break
                    elif key == "valueForm" and value == "range":
                        ranges_allowed.append(objectValue.attrib["valueAllowed"])
                        break
            try:
                context_rules = x.getparent().getparent().getparent().attrib['rulesContext']
            except KeyError:
                context_rules = ""
            br_decision_ref = x.getparent().find('brDecisionRef')
            br_decision_ident_number = br_decision_ref.get('brDecisionIdentNumber') if br_decision_ref is not None else None
            br_severity_level = x.getparent().get('brSeverityLevel')
            if br_severity_level is None:
                br_severity_level = default_br_severity_level
            allowed_object_flag_dict.append({
                    'xpath': str(nodes_to_check[counter].text),
                    'Brex': str(brex),
                    'ObjectFlag': str(x.attrib['allowedObjectFlag']),
                    'objectUse': str(x.getparent().xpath('objectUse')[0].text),
                    'contextRules': context_rules,
                    'values_allowed': values_allowed,
                    'regex_allowed': regex_allowed,
                    'ranges_allowed': ranges_allowed,
                    'brDecisionIdentNumber': br_decision_ident_number,
                    'brSeverityLevel': br_severity_level
                }
            )
        if debug:
            with open(clean_path(join(expanduser("~/Desktop"), f'brex_{basename(brex)}.json')), 'w', encoding="utf-8") as _:
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

    def _check_object_flag_0(self, schema: str, brex_violations: dict, root: any, value: any):
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
                                        brex_violations[value["Brex"]]['0'].append({
                                            'Line': list_xml_content.index(element) + 1,
                                            'Description': value["objectUse"],
                                            'Xpath': value['xpath'],
                                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                                            'BrSeverityLevel': value.get('brSeverityLevel'),
                                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                                        )
                    proc.exception_clear()
            else:
                try:
                    selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
                    result = selector.select(root)
                except elementpath.ElementPathError as e:
                    brex_violations[value["Brex"]]['xpathError'].append({
                        'Description': value["objectUse"],
                        'Xpath': value['xpath'],
                        'Error': str(e),
                        'BrDecisionIdentNumber': value.get('brDecisionIdentNumber')}
                    )
                    return brex_violations
                if isinstance(result, bool):
                    if result:
                        brex_violations[value["Brex"]]['0'].append({
                            'Line': "(Boolean condition -> Interpret XPath)",
                            'Description': value["objectUse"],
                            'Xpath': value['xpath'],
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                        )
                else:
                    for element in result:
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
                        brex_violations[value["Brex"]]['0'].append({
                            'Line': line_no,
                            'Description': value["objectUse"],
                            'Xpath': value['xpath'],
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                        )
        return brex_violations

    def _check_object_flag_1(self, schema: str, brex_violations: dict, root: any, value: any):
        if value['contextRules'] == schema or value['contextRules'] == "":
            try:
                selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
                result = selector.select(root)
            except elementpath.ElementPathError as e:
                brex_violations[value["Brex"]]['xpathError'].append({
                    'Description': value["objectUse"],
                    'Xpath': value['xpath'],
                    'Error': str(e),
                    'BrDecisionIdentNumber': value.get('brDecisionIdentNumber')}
                )
                return brex_violations
            if isinstance(result, bool):
                violation = not result
            elif isinstance(result, (str, int, float)):
                violation = not result
            else:
                violation = len(result) == 0
            if violation:
                brex_violations[value["Brex"]]['1'].append({
                            'Description': value["objectUse"],
                            'Xpath': value['xpath'],
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                            )
            elif not isinstance(result, (bool, str, int, float)) and (
                    value["values_allowed"] or value["regex_allowed"] or value["ranges_allowed"]):
                brex_violations[value["Brex"]]['2'].extend(self._check_object_values(value, result))
        return brex_violations

    def _check_object_values(self, value: any, elements: any) -> list:
        """Check a set of matched nodes against a rule's `objectValue` children.

        Shared by any flag whose matched nodes must additionally satisfy a value
        constraint (port of `check_objects_values`, `s1kd-brexcheck.c:275-304`,
        which applies value checking to a rule's matched node-set regardless of
        `allowedObjectFlag`). Ref §3.8.

        Args:
            value (any): rule dict from `_show_rules`, carrying `values_allowed`
                / `regex_allowed` / `ranges_allowed`
            elements (any): node-set matched by `value['xpath']`

        Returns:
            list: one violation dict per element whose value matches none of the
                allowed values, patterns or ranges
        """
        violations = []
        for element in elements:
            valid_elem = False
            if isinstance(element, etree._Element):
                element_value = element.text or ""
            else:
                element_value = element if isinstance(element, str) else str(element)
            if element_value not in value["values_allowed"]:
                if len(value["regex_allowed"]) > 0:
                    if any(bool(fullmatch(regex, element_value, V1)) for regex in value["regex_allowed"]):
                        valid_elem = True
                if not valid_elem and len(value["ranges_allowed"]) > 0:
                    if any(is_in_set(element_value, value_range) for value_range in value["ranges_allowed"]):
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
                                if attrib_name in elem and element_value in elem:
                                    line_no = ind + 1
                        else:
                            line_no = "x"
                violations.append({
                    'Line': line_no,
                    'Description': f'Element/Attribute ({element_value}) did not match the object values.',
                    'Xpath': value['xpath'],
                    'Single Values': [value["values_allowed"]],
                    'Pattern Values': [value["regex_allowed"]],
                    'Range Values': [value["ranges_allowed"]],
                    'ObjectUse': value["objectUse"],
                    'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                    'BrSeverityLevel': value.get('brSeverityLevel'),
                    'Fail': self._is_severity_failure(value.get('brSeverityLevel'))})
        return violations

    def _check_object_flag_2(self, schema: str, brex_violations: dict, root: any, value: any):
        if ('values_allowed' in value or 'regex_allowed' in value or 'ranges_allowed' in value) and (value['contextRules'] == schema or value['contextRules'] == ""):
            try:
                selector = elementpath.Selector(value['xpath'], namespaces=NS_DICT)
                result = selector.select(root)
            except elementpath.ElementPathError as e:
                brex_violations[value["Brex"]]['xpathError'].append({
                    'Description': value["objectUse"],
                    'Xpath': value['xpath'],
                    'Error': str(e),
                    'BrDecisionIdentNumber': value.get('brDecisionIdentNumber')}
                )
                return brex_violations
            if type(result) is not bool:
                brex_violations[value["Brex"]]['2'].extend(self._check_object_values(value, result))
        return brex_violations

    def _get_sns_rules_group(self) -> any:
        """Merge the `snsRules` element from every active BREX into one root.

        Port of `check_brex_sns` (`s1kd-brexcheck.c:1144-1173`), which builds a
        combined `snsRulesGroup` document from the `snsRules` of every BREX
        passed to the tool so a code can be checked against rules defined in
        any of them.

        Returns:
            any: `snsRulesGroup` lxml element (childless if no BREX defines `snsRules`)
        """
        group = etree.Element("snsRulesGroup")
        for brex in self._brex_list[0]:
            with open(clean_path(brex), "r", encoding="utf-8") as _:
                brex_content = _.read()
            brex_content = delete_first_line(brex_content)
            brex_tree = etree.parse(StringIO(brex_content))
            sns_rules = brex_tree.find(".//snsRules")
            if sns_rules is not None:
                group.append(deepcopy(sns_rules))
        return group

    def _sns_should_check(self, code: str, tag: str, ctx: any, sns_mode: str = "normal") -> bool:
        """Decide whether an SNS code level needs to be checked.

        Port of `should_check` (`s1kd-brexcheck.c:1038-1054`):

        - `strict`: always check every level; a placeholder code ("0" for
          sub/sub-sub-system, "00"/"0000" otherwise) is not treated as
          shorthand and must itself match a defined `snsCode`.
        - `unstrict`: check a level only if the current scope defines any
          rule for that level at all; if it defines none, any code (whether
          or not it looks like a placeholder) is accepted without checking.
        - `normal` (default): a non-placeholder code is always checked; a
          placeholder code is only checked if the current scope actually
          defines rules for that level.

        Args:
            code (str): the SNS code value from the data module's dmCode
            tag (str): the SNS rule element to look for (snsSystem/snsSubSystem/snsSubSubSystem/snsAssy)
            ctx (any): current scope to search within (an snsRulesGroup or a matched snsXxx node)
            sns_mode (str): one of `SNS_MODES` ("normal", "strict", "unstrict")

        Returns:
            bool: True if this level should be checked
        """
        if sns_mode == "strict":
            return True
        if sns_mode == "unstrict":
            return ctx.find(f".//{tag}") is not None
        if tag in ("snsSubSystem", "snsSubSubSystem"):
            non_placeholder = code != "0"
        else:
            non_placeholder = code not in ("00", "0000")
        return non_placeholder or ctx.find(f".//{tag}") is not None

    def _check_sns_rules(self, sns_rules_group: any, dmod_root: any, sns_mode: str = "normal") -> any:
        """Check a data module's SNS code against the merged SNS rules.

        Port of `check_brex_sns_rules` (`s1kd-brexcheck.c:1057-1144`): walks
        `systemCode` -> `subSystemCode` -> `subSubSystemCode` -> `assyCode` down
        `snsSystem` / `snsSubSystem` / `snsSubSubSystem` / `snsAssy`, stopping at
        the first failing level.

        Args:
            sns_rules_group (any): merged `snsRulesGroup` element from `_get_sns_rules_group`
            dmod_root (any): root element of the data module being checked
            sns_mode (str): one of `SNS_MODES` ("normal", "strict", "unstrict");
                see `_sns_should_check`

        Returns:
            any: dict describing the first failing level, or None if the SNS code is valid
                 (or the object being checked is not a data module, or has no `dmCode`)
        """
        if dmod_root.tag != "dmodule":
            return None

        dm_code = dmod_root.find(".//dmIdent/dmCode")
        if dm_code is None:
            return None

        system_code = dm_code.get("systemCode", "")
        sub_system_code = dm_code.get("subSystemCode", "")
        sub_sub_system_code = dm_code.get("subSubSystemCode", "")
        assy_code = dm_code.get("assyCode", "")

        levels = (
            ("systemCode", "snsSystem", system_code,
             system_code),
            ("subSystemCode", "snsSubSystem", sub_system_code,
             f"{system_code}-{sub_system_code}"),
            ("subSubSystemCode", "snsSubSubSystem", sub_sub_system_code,
             f"{system_code}-{sub_system_code}{sub_sub_system_code}"),
            ("assyCode", "snsAssy", assy_code,
             f"{system_code}-{sub_system_code}{sub_sub_system_code}-{assy_code}"),
        )

        ctx = sns_rules_group
        for code_name, tag, code, invalid_value in levels:
            if not self._sns_should_check(code, tag, ctx, sns_mode):
                continue
            match = ctx.xpath(f".//{tag}[snsCode=$code]", code=code)
            if not match:
                return {"code": code_name, "invalidValue": invalid_value}
            ctx = match[0]

        return None

    def _get_notation_rules_group(self) -> any:
        """Merge the `notationRuleList` element from every active BREX into one root.

        Port of the notation-rule loading step in `check_brex_notations`
        (`s1kd-brexcheck.c:1229-1256`), which builds a combined
        `notationRuleGroup` document from the `notationRuleList` of every
        BREX passed to the tool so an entity's notation can be checked
        against rules defined in any of them.

        Returns:
            any: `notationRuleGroup` lxml element (childless if no BREX defines `notationRuleList`)
        """
        group = etree.Element("notationRuleGroup")
        for brex in self._brex_list[0]:
            with open(clean_path(brex), "r", encoding="utf-8") as _:
                brex_content = _.read()
            brex_content = delete_first_line(brex_content)
            brex_tree = etree.parse(StringIO(brex_content))
            notation_rule_list = brex_tree.find(".//notationRuleList")
            if notation_rule_list is not None:
                group.append(deepcopy(notation_rule_list))
        return group

    def _check_entity_notation(self, entity_name: str, notation_name: str, notation_rule_group: any) -> any:
        """Check a single unparsed (`NDATA`) entity's notation against the notation rules.

        Port of `check_entity` (`s1kd-brexcheck.c:1176-1201`). A notation is
        accepted if some `notationRule/notationName` names it with
        `@allowedNotationFlag != "0"`. Otherwise the entity is reported
        against the `objectUse` of the first `notationRule` in the merged
        rule group -- the C original's fallback XPath,
        `(//notationRule[notationName=X]|//notationRule)[1]`, unions a
        subset with its own superset, so it always resolves to the first
        `notationRule` in document order regardless of whether it actually
        names this notation.

        Args:
            entity_name (str): the `<!ENTITY>` name (for reporting only)
            notation_name (str): the NDATA notation the entity declares --
                for an unparsed entity, libxml2/lxml store the NDATA target
                name in the entity's content, which is what `entity->content`
                reads in the C original
            notation_rule_group (any): merged `notationRuleGroup` element
                from `_get_notation_rules_group`

        Returns:
            any: dict describing the violation, or None if the notation is allowed
        """
        allowed = notation_rule_group.xpath(
            ".//notationRule[notationName=$name and notationName/@allowedNotationFlag != '0']",
            name=notation_name,
        )
        if allowed:
            return None

        rules = notation_rule_group.xpath(".//notationRule")
        rule = rules[0] if rules else None
        object_use = None
        if rule is not None:
            use_node = rule.find("objectUse")
            if use_node is not None:
                object_use = "".join(use_node.itertext()) or None

        return {
            "Entity": entity_name,
            "Notation": notation_name,
            "Description": object_use or f"Notation '{notation_name}' is not allowed.",
        }

    def _check_notation_rules(self, notation_rule_group: any, dmod_tree: any) -> list:
        """Check every unparsed entity declared in the object's internal DTD subset.

        Port of `check_brex_notation_rules` (`s1kd-brexcheck.c:1204-1227`):
        walks the `ENTITY` declarations of the internal DTD subset, and for
        each external, unparsed (`NDATA`) entity checks its notation via
        `_check_entity_notation`. Objects with no internal DTD subset (the
        common case for XSD-validated S1000D 4.x+ content) are not checked,
        matching the original's `if (!(dtd = dmod_doc->intSubset)) return 0;`.

        Args:
            notation_rule_group (any): merged `notationRuleGroup` element
                from `_get_notation_rules_group`
            dmod_tree (any): parsed `ElementTree` of the object being checked

        Returns:
            list: violation records, one per entity naming a disallowed notation
        """
        internal_dtd = dmod_tree.docinfo.internalDTD
        if internal_dtd is None:
            return []

        violations = []
        for entity in internal_dtd.iterentities():
            # Unparsed (NDATA) entities are the only ones with both a system
            # identifier and content set (content holds the NDATA notation
            # name); internal entities have no system_url, external parsed
            # entities have no content -- this mirrors the C original's
            # etype == XML_EXTERNAL_GENERAL_UNPARSED_ENTITY check.
            if entity.system_url is None or entity.content is None:
                continue
            violation = self._check_entity_notation(entity.name, entity.content, notation_rule_group)
            if violation is not None:
                violations.append(violation)
        return violations

    def _check_rules(self, debug: bool = False, include_tqdm: bool = False, sns_mode: str = "normal") -> dict:
        """Traverses through every node of the brex and checks the rules through the given xpaths.
        For objectFlag 0 we also get the line of the error
        For objectFlag 1 we only get the Description of the rule that was violated
        For objectFlag 2 we get a list containing all 'single' values and a list containing all 'pattern' values
                         and we might get the line of the error

        Args:
            debug (bool): dump intermediate rule/error data for inspection
            include_tqdm (bool): show a progress bar while checking content rules
            sns_mode (str): one of `SNS_MODES` ("normal", "strict", "unstrict");
                see `_sns_should_check`

        Returns:
            any: Dictionary with all errors
        """
        schema = get_schema_from_xml(self._xml_content)
        brex_violations_dict = {}
        for brex in self._brex_list[0]:
            brex_violations_dict[brex] = {
                '0': [],
                '1': [],
                '2': [],
                'xpathError': []
            }
        brex_violations_dict["brexFallback"] = list(self._brex_fallbacks)
        root = etree.parse(self._xml_path)

        dmod_root = root.getroot()
        if dmod_root.tag == "dmodule":
            sns_rules_group = self._get_sns_rules_group()
            sns_error = self._check_sns_rules(sns_rules_group, dmod_root, sns_mode)
            brex_violations_dict["sns"] = [] if sns_error is None else [{
                "code": sns_error["code"],
                "invalidValue": sns_error["invalidValue"],
                "Description": f"{sns_error['code']} is not valid according to the SNS rules.",
            }]

        notation_rule_group = self._get_notation_rules_group()
        brex_violations_dict["notations"] = self._check_notation_rules(notation_rule_group, root)

        all_content_rules = []
        for brex in self._brex_list[0]:
            content_rules = self._show_rules(brex, debug=debug)
            all_content_rules += content_rules

        if debug:
            with open(clean_path(join(expanduser("~/Desktop"), "All_content_rules.txt")), 'w', encoding="utf-8") as _:
                for rule in all_content_rules:
                    _.write(str(rule) + "\n")
        container = tqdm(all_content_rules) if include_tqdm else all_content_rules
        for value in container:
            if value["ObjectFlag"] == '0':
                brex_violations_dict |= self._check_object_flag_0(schema, brex_violations_dict, root, value)
            if value["ObjectFlag"] == '1':
                brex_violations_dict |= self._check_object_flag_1(schema, brex_violations_dict, root, value)
            if value["ObjectFlag"] == '2':
                if value["values_allowed"] != [] or value["regex_allowed"] != [] or value["ranges_allowed"] != []:
                    brex_violations_dict |= self._check_object_flag_2(schema, brex_violations_dict, root, value)
        return brex_violations_dict

    def _append_summary(self, object_flag_dict: dict) -> str:
        """Counts the number of actual Brex violations (flags 0, 1 and 2, plus SNS and
        notation rules) for a xml. `xpathError` entries are diagnostics about a rule
        that could not be evaluated, not violations, and are excluded from the count.

        A content-rule violation whose resolved `brSeverityLevel` is marked `fail="no"`
        in the `.brseveritylevels` file (see `_is_severity_failure`) is reported as a
        warning instead of an error and does not count towards the failing total. SNS
        and notation-rule violations have no associated severity level and always
        count as errors.

        Args:
            object_flag_dict (dict): mapping of brex path to its '0'/'1'/'2'/'xpathError' violation
                lists, plus optional 'sns' / 'notations' keys holding SNS and notation violations,
                and a 'brexFallback' key listing any built-in BREX substitutions (informational
                only -- a substitution is not itself a violation and does not count towards the total)

        Returns:
            str: human-readable violation count, e.g. "3 Errors" or "3 Errors, 1 Warnings"
        """
        error_count = 0
        warning_count = 0
        for key, brex_result in object_flag_dict.items():
            if key == "brexFallback":
                continue
            if key in ("sns", "notations"):
                error_count += len(brex_result)
                continue
            for flag in ('0', '1', '2'):
                for violation in brex_result[flag]:
                    if violation.get('Fail', True):
                        error_count += 1
                    else:
                        warning_count += 1
        if warning_count:
            return f"{error_count} Errors, {warning_count} Warnings"
        return f"{error_count} Errors"
    
    def validate(self, debug: bool = False, include_tqdm: bool = False, sns_mode: str = "normal") -> dict:
        """Check xml against all brexes and dump the results into a JSon file

        Args:
            debug (bool): dump intermediate rule/error data for inspection
            include_tqdm (bool): show a progress bar while checking files/rules
            sns_mode (str): SNS shorthand mode, one of `SNS_MODES`. Port of
                `should_check` (`s1kd-brexcheck.c:1038`):

                - `"normal"` (default): optional levels default to `0` / `00` /
                  `0000`, i.e. a placeholder code is only checked if the BREX
                  actually defines rules for that level.
                - `"strict"`: no shorthand — every level's code must match a
                  `snsCode` defined by the BREX, including placeholders.
                - `"unstrict"`: any code is valid at a level the BREX defines
                  no rules for, whether or not it looks like a placeholder.

        Raises:
            ValueError: if `sns_mode` is not one of `SNS_MODES`
        """
        if sns_mode not in SNS_MODES:
            raise ValueError(f"sns_mode must be one of {SNS_MODES}, got {sns_mode!r}")
        if self._xml_dir:
            files = [_ for _ in listdir(self._xml_dir) if ".xml" in _.lower() and "-022a-" not in _.lower()]
            had_explicit_brex_list = self._brex_list[0] is not None
            had_explicit_brex_dir_path = self._brex_dir_path[1] is True
            initial_brex_list = self._brex_list
            initial_brex_dir_path = self._brex_dir_path
            results = {}
            container = tqdm(files) if include_tqdm else files
            for _xml in container:
                self.set_xml(join(self._xml_dir, _xml))
                self._init_brex_list()
                result = self._check_rules(debug=debug, include_tqdm=include_tqdm, sns_mode=sns_mode)
                result["Summary"] = self._append_summary(result)
                results[_xml] = result
                self._brex_list = initial_brex_list if had_explicit_brex_list else (None, None)
                self._brex_dir_path = initial_brex_dir_path if had_explicit_brex_dir_path else (None, None)
            if debug:
                with open(clean_path(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_dir)}.json')), 'w', encoding="utf-8") as _:
                    dump(results, _, indent=4)
            return results
        else:
            self._init_brex_list()
            result = self._check_rules(debug=debug, sns_mode=sns_mode)
            summary = self._append_summary(result)
            result["Summary"] = summary
            if debug:
                with open(clean_path(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_path)}.json')), 'w', encoding="utf-8") as _:
                    dump(result, _, indent=4)
            return result

