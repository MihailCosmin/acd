from datetime import datetime

from copy import deepcopy

from decimal import Decimal

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
        self._brex_search_paths = []
        self._brex_recursive_search = True
        self._use_default_brex = False
        self._brex_fallbacks = []

        self._severity_levels_path = None
        self._severity_levels_search = True
        self._severity_levels = (None, False)

        self._xinclude = False
        self._resolve_entities = True
        self._load_external_dtd = False
        self._allow_network = False
        self._ignore_empty = False

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
            resolved = find_document_by_reference(brex_ref, self._brex_dir_path[0],
                                                   recursive=self._brex_recursive_search)
            if resolved is None:
                for search_path in self._brex_search_paths:
                    resolved = find_document_by_reference(brex_ref, search_path,
                                                           recursive=self._brex_recursive_search)
                    if resolved is not None:
                        break
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

    def add_brex_search_path(self, brex_path: str):
        """Add an additional directory to search for referenced BREX data
        modules, equivalent to `s1kd-brexcheck`'s repeatable `-I`/`--include`
        option. Can be called multiple times to register several search
        paths.

        Each path added here is only searched if the referenced BREX was not
        found in the primary path (`set_brex_path`, or the checked XML's own
        directory when `set_brex_path` was not called); paths are then tried
        in the order they were added, stopping at the first match.

        Args:
            brex_path (str): directory to add to the BREX search path list
        """
        if isdir(brex_path):
            self._brex_search_paths.append(brex_path)
        else:
            raise BrexNotFound(f"The given search path {brex_path} seems to be leading to a file. \
                Please make sure to input the path of a directory containing brex data modules.".replace("                ", ""))

    def clear_brex_search_paths(self):
        """Remove all additional BREX search paths added via
        `add_brex_search_path`.
        """
        self._brex_search_paths = []

    def set_brex_recursive_search(self, enabled: bool):
        """Enable or disable recursive search (equivalent to
        `s1kd-brexcheck`'s `-r`/`--recursive`) of the primary BREX directory
        (`set_brex_path`, or the checked XML's own directory) and of every
        path added via `add_brex_search_path`. Enabled by default; disable it
        to only look directly inside each search directory, ignoring
        subdirectories.

        Args:
            enabled (bool): whether to search subdirectories recursively
        """
        self._brex_recursive_search = enabled

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

    def set_xinclude(self, enabled: bool = True):
        """Equivalent to `s1kd-brexcheck`'s `--xinclude`: resolve `xi:include`
        elements in the checked object and every BREX file before checking,
        via lxml's `ElementTree.xinclude()`. Mirrors `read_xml_doc`'s
        `xmlXIncludeProcessFlags` call (`s1kd_tools.c:538-539`), which is
        applied uniformly to every CSDB object it reads. Disabled by default,
        matching libxml2's `XML_PARSE_XINCLUDE` default.

        Args:
            enabled (bool): whether to process XInclude directives
        """
        self._xinclude = enabled

    def set_resolve_entities(self, enabled: bool = True):
        """Control whether entity references are substituted with their
        declared content (lxml's `resolve_entities` parser option),
        equivalent to `--noent`. Entities declared in the internal DTD
        subset are read the same way regardless (see `_check_notation_rules`,
        which reads them directly off `docinfo.internalDTD`); this only
        affects whether an entity *reference* elsewhere in the content
        survives in the parsed tree as a placeholder or is replaced in place
        before rules are checked. Enabled by default -- lxml's own default,
        already stricter than `s1kd-brexcheck`'s default of leaving
        references unresolved unless `--noent` is given.

        Args:
            enabled (bool): whether to substitute entity references with content
        """
        self._resolve_entities = enabled

    def set_entity_resolution(self, load_external_dtd: bool = True, allow_network: bool = False):
        """Enable resolving entities declared in an *external* DTD subset
        (`SYSTEM`/`PUBLIC` entities), equivalent to `s1kd-brexcheck`'s
        `--dtdload` (and, if `allow_network` is also set, `--net`). Both are
        off by default: an external DTD is not fetched, and network access is
        never allowed unless explicitly requested here, so that parsing a
        checked object cannot trigger unexpected file or network access on
        its own.

        Args:
            load_external_dtd (bool): fetch and parse the object's external DTD subset
            allow_network (bool): allow the parser to resolve `http(s)://`
                DTD/entity references over the network; only takes effect
                when `load_external_dtd` is also True. Defaults to False.
        """
        self._load_external_dtd = load_external_dtd
        self._allow_network = allow_network and load_external_dtd

    def set_xml_catalog(self, catalog_path: str):
        """Register an XML catalog file for resolving external DTD/entity/
        schema references, equivalent to `--xml-catalog <file>`
        (`xmlLoadCatalog`). lxml has no direct catalog-loading binding, so
        this appends the path to the `XML_CATALOG_FILES` environment
        variable, which libxml2 reads the first time it needs to consult the
        global catalog -- the standard way to drive libxml2's catalog
        resolution from Python. Can be called multiple times to register
        several catalogs, same as repeating the command-line flag.

        Note: libxml2 only reads `XML_CATALOG_FILES` once per process, on its
        first catalog lookup, so a catalog registered after that point may
        not take effect within the same process.

        Args:
            catalog_path (str): path to an XML (or SGML) catalog file
        """
        if not isfile(catalog_path):
            raise BrexNotFound(f"The given catalog path {catalog_path} does not point to a file.")
        entries = environ.get("XML_CATALOG_FILES", "").split()
        if catalog_path not in entries:
            entries.append(catalog_path)
            environ["XML_CATALOG_FILES"] = " ".join(entries)

    def set_ignore_empty(self, enabled: bool = True):
        """Equivalent to `-e`/`--ignore-empty`: silently skip a checked object
        that is empty or not well-formed XML, instead of raising. In
        directory mode (`set_xml_dir`) the file is left out of the results
        entirely, matching `s1kd-brexcheck`'s `continue`; for a single object
        (`set_xml`/`validate()`), the skip is reported as
        `{"Skipped": True, "Summary": "..."}` instead of raising.

        Args:
            enabled (bool): whether to skip empty/non-XML input instead of raising
        """
        self._ignore_empty = enabled

    def _build_xml_parser(self) -> etree.XMLParser:
        """Build the lxml parser used for both the checked object and every
        BREX file, honouring the parser options set via `set_resolve_entities`
        / `set_entity_resolution`. Mirrors `DEFAULT_PARSE_OPTS`
        (`s1kd_tools.c:14`), which `read_xml_doc` applies uniformly to every
        CSDB object it reads.

        Returns:
            etree.XMLParser: configured parser
        """
        return etree.XMLParser(
            resolve_entities=self._resolve_entities,
            load_dtd=self._load_external_dtd,
            no_network=not self._allow_network,
            huge_tree=True,
        )

    def _finish_parse(self, tree: any) -> any:
        """Apply XInclude processing to a freshly parsed tree, if enabled via
        `set_xinclude`. Equivalent to `read_xml_doc`'s
        `xmlXIncludeProcessFlags` call (`s1kd_tools.c:538-539`).

        Args:
            tree (any): parsed `ElementTree`

        Returns:
            any: the same tree, with XInclude directives resolved in place if enabled
        """
        if self._xinclude:
            tree.xinclude()
        return tree

    def _parse_xml_file(self, path: str) -> any:
        """Parse an XML file from disk with the configured parser options.

        Args:
            path (str): path to the XML file

        Returns:
            any: parsed `ElementTree`
        """
        return self._finish_parse(etree.parse(path, parser=self._build_xml_parser()))

    def _parse_xml_text(self, content: str) -> any:
        """Parse XML held in a string with the configured parser options.

        Args:
            content (str): XML content

        Returns:
            any: parsed `ElementTree`
        """
        return self._finish_parse(etree.parse(StringIO(content), parser=self._build_xml_parser()))

    def _is_valid_xml_file(self, path: str) -> bool:
        """Return whether `path` parses as well-formed XML with the
        configured parser options. Used by `set_ignore_empty` to decide
        whether a checked object should be silently skipped, equivalent to
        the `read_xml_doc(...) == NULL` check in `s1kd-brexcheck.c:2151-2160`.

        Args:
            path (str): path to the object to test

        Returns:
            bool: True if the file parses as XML; False if it is missing,
                empty, or not well-formed
        """
        try:
            self._parse_xml_file(path)
            return True
        except (etree.XMLSyntaxError, OSError):
            return False

    def _get_object_rule_nodes(self, brex: str, schema: str = None) -> any:
        """Return all `objectPath` nodes whose enclosing `contextRules` is
        unqualified or targets the given schema, selected with the descendant
        axis so nested/grouped rules at any depth are found. Uses real XPath
        (`Element.xpath`) rather than the restricted ElementPath `findall`,
        which also avoids lxml's "This search incorrectly ignores the root
        element" FutureWarning on a leading `//`.

        Args:
            brex (str): path of the brex
            schema (str): the object's declared schema; rules whose
                `rulesContext`/`context` names a different schema are
                excluded at selection time. `None` returns every rule.

        Returns:
            any: Set of nodes
        """
        with open(clean_path(brex), "r", encoding="utf-8") as _:
            brex_content = _.read()
        brex_content = delete_first_line(brex_content)
        root = self._parse_xml_text(brex_content).getroot()
        if schema is None:
            nodes = root.xpath('//contextRules//structureObjectRule/objectPath')
            # S1000D <= 3.0 spelling
            nodes += root.xpath('//contextrules//objrule/objpath')
        else:
            nodes = root.xpath(
                '//contextRules[not(@rulesContext) or @rulesContext=$schema]'
                '//structureObjectRule/objectPath',
                schema=schema,
            )
            # S1000D <= 3.0 spelling
            nodes += root.xpath(
                '//contextrules[not(@context) or @context=$schema]//objrule/objpath',
                schema=schema,
            )
        return nodes

    def _show_rules(self, brex: str, schema: str = None, debug: bool = False) -> any:
        """Creates a, in nested dictionaries structured, JSON file containing all necessary information about the brex rules i.e.
        xpath, objectflag, objectUse, objectValues et Al.

        Args:
            brex (str): brex_path
            schema (str): the object's declared schema, passed through to
                `_get_object_rule_nodes` to filter rules at selection time

        Returns:
            any: Nested Dictionary
        """
        nodes_to_check = self._get_object_rule_nodes(brex, schema)
        default_br_severity_level = None
        if len(nodes_to_check) > 0:
            default_br_severity_level = nodes_to_check[0].getroottree().getroot().get('defaultBrSeverityLevel')
        allowed_object_flag_dict = []
        for counter, x in enumerate(nodes_to_check):
            values_allowed = []
            regex_allowed = []
            ranges_allowed = []
            for objectValue in x.getparent().xpath('objectValue|objval'):
                # S1000D <= 3.0 spells these @valtype and @val1[~@val2] instead
                # of @valueForm and @valueAllowed (a range is written as two
                # attributes rather than one "first~last" string).
                value_form = objectValue.get('valueForm', objectValue.get('valtype'))
                value_allowed = objectValue.get('valueAllowed')
                if value_allowed is None and objectValue.get('val1') is not None:
                    value_allowed = objectValue.get('val1')
                    val2 = objectValue.get('val2')
                    if val2 is not None:
                        value_allowed = f"{value_allowed}~{val2}"
                if value_form == "single":
                    values_allowed.append(value_allowed)
                elif value_form == "pattern":
                    regex_allowed.append(translate_xsd_regex_to_python(value_allowed))
                elif value_form == "range":
                    ranges_allowed.append(value_allowed)
            context_group = next(x.iterancestors('contextRules', 'contextrules'), None)
            context_rules = (
                context_group.get('rulesContext', context_group.get('context', ''))
                if context_group is not None else ''
            )
            br_decision_ref = x.getparent().find('brDecisionRef')
            br_decision_ident_number = br_decision_ref.get('brDecisionIdentNumber') if br_decision_ref is not None else None
            br_severity_level = x.getparent().get('brSeverityLevel')
            if br_severity_level is None:
                br_severity_level = default_br_severity_level
            # Register every namespace in scope at this objectPath node (lxml's
            # nsmap includes prefixes declared on ancestors), rather than relying
            # on a hard-coded rdf+xsi dictionary. The default namespace (lxml key
            # None) is remapped to '' as elementpath expects. NS_DICT is kept as
            # a base so rdf/xsi stay resolvable even if a rule's local scope
            # happens not to declare them. Ref §3.12.
            namespaces = dict(NS_DICT)
            namespaces.update(
                {(prefix or ''): uri for prefix, uri in x.nsmap.items()}
            )
            allowed_object_flag_dict.append({
                    'xpath': str(nodes_to_check[counter].text),
                    'Brex': str(brex),
                    'ObjectFlag': x.get('allowedObjectFlag', x.get('objappl')),
                    'objectUse': str(x.getparent().xpath('objectUse|objuse')[0].text),
                    'contextRules': context_rules,
                    'values_allowed': values_allowed,
                    'regex_allowed': regex_allowed,
                    'ranges_allowed': ranges_allowed,
                    'brDecisionIdentNumber': br_decision_ident_number,
                    'brSeverityLevel': br_severity_level,
                    'namespaces': namespaces
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

    def _select_with_nodes(self, selector: any, root: any) -> tuple:
        """Evaluate a compiled `elementpath.Selector` the same way `Selector.select`
        does, while also returning the raw XPath node backing each item of a
        node-set result. `Selector.select` (via `XPathToken.get_results`) reduces
        every node to its plain value (an lxml element, or a bare string for an
        attribute/text result), discarding the node's parent/position -- exactly
        the information needed to compute a violating node's canonical XPath and
        a copy of its owning element (categories D2/D3). This re-implements that
        reduction from the lower-level, un-formatted `root_token.select()` so the
        formatted half of the return value stays identical to plain `.select()`.

        Args:
            selector (any): compiled rule selector (`elementpath.Selector`)
            root (any): document root to evaluate the selector against

        Returns:
            tuple: `(result, nodes)`. `result` is exactly what `selector.select(root)`
                would return. `nodes` is `None` when `result` is a bare scalar (no
                node backs a computed boolean/number), otherwise a list of raw
                `elementpath` XPath node objects (or `None` per position for a
                non-node item) aligned with `result`.
        """
        context = elementpath.XPathContext(root, schema=selector.parser.schema)
        raw_items = list(selector.root_token.select(context))

        values = []
        nodes = []
        for item in raw_items:
            if isinstance(item, elementpath.xpath_nodes.XPathNode):
                values.append(item.value)
                nodes.append(item)
            else:
                values.append(item)
                nodes.append(None)

        if len(raw_items) == 1 and not isinstance(
                raw_items[0], (elementpath.xpath_nodes.ElementNode, elementpath.xpath_nodes.DocumentNode)):
            if isinstance(raw_items[0], (bool, int, float, Decimal)):
                return raw_items[0], None
            elif selector.root_token.label in ('function', 'literal'):
                return values[0], None

        return values, nodes

    def _resolve_owning_element(self, node: any) -> any:
        """Resolve a raw XPath node (from `_select_with_nodes`) to the lxml
        element that backs it, walking up to the parent for an attribute/text
        result (which has no element of its own). Shared by
        `_node_xpath_and_copy` (categories D2/D3) and `_node_line_number`
        (category D1).

        Args:
            node (any): raw XPath node from `_select_with_nodes`'s `nodes`
                list, or `None`

        Returns:
            any: the backing `lxml.etree._Element`, or `None` when `node` is
                `None` or does not resolve to one
        """
        if node is None:
            return None
        element = getattr(node, 'obj', None)
        if not isinstance(element, etree._Element):
            parent = getattr(node, 'parent', None)
            element = getattr(parent, 'obj', None) if parent is not None else None
        if not isinstance(element, etree._Element):
            return None
        return element

    def _node_line_number(self, node: any) -> any:
        """Real line number of a violating node, read from the parsed tree's
        `sourceline` (lxml's binding to libxml2's `xmlGetLineNo`) instead of
        scanning the raw XML text for the attribute name. An attribute or
        text result has no `sourceline` of its own, so it is reported against
        its owning element's line, same as `_node_xpath_and_copy` does for
        the node's XPath/copy. Ref §3.12, category D1.

        Args:
            node (any): raw XPath node from `_select_with_nodes`'s `nodes`
                list, or `None` when no such node is available

        Returns:
            any: the 1-based line number (`int`), or `None` when it cannot be
                resolved (no backing node, or the node carries no line info)
        """
        element = self._resolve_owning_element(node)
        if element is None:
            return None
        return element.sourceline

    def _node_xpath_and_copy(self, node: any, deep_copy_nodes: bool = False) -> tuple:
        """Resolve a raw XPath node (from `_select_with_nodes`) into the two
        fields `s1kd-brexcheck`'s `dump_nodes_xml` attaches to every violation:
        the node's canonical XPath (port of `xpath_of`, `s1kd_tools.c:59-144`,
        using `elementpath`'s own equivalent node-path computation instead of
        re-walking the tree) and a copy of its owning element, serialised to an
        XML string. An attribute or text result is reported against its owning
        element (`if (node->type == XML_ATTRIBUTE_NODE) node = node->parent;` in
        the C original), since a bare attribute/text value has no subtree of its
        own to copy.

        Args:
            node (any): raw XPath node from `_select_with_nodes`'s `nodes` list,
                or `None` when the violation has no backing node (e.g. a flag-1
                "required but missing" violation, or a boolean-valued rule)
            deep_copy_nodes (bool): copy the full subtree (all descendants),
                equivalent to `-8`/`--deep-copy-nodes`. Defaults to a shallow
                copy of just the element's own tag and attributes, matching
                `xmlCopyNode(node, 2)` (properties only, no children).

        Returns:
            tuple: `(canonical_xpath, xml_snippet)`, both `None` when `node` is
                `None` or does not resolve to an lxml element
        """
        if node is None:
            return None, None

        try:
            canonical_xpath = node.extended_path
        except AttributeError:
            canonical_xpath = None

        element = self._resolve_owning_element(node)
        if element is None:
            return canonical_xpath, None

        try:
            if deep_copy_nodes:
                copy_elem = deepcopy(element)
            else:
                copy_elem = etree.Element(element.tag, nsmap=element.nsmap)
                for key, val in element.attrib.items():
                    copy_elem.set(key, val)
            xml_snippet = etree.tostring(copy_elem, encoding="unicode")
        except (TypeError, ValueError):
            xml_snippet = None

        return canonical_xpath, xml_snippet

    def _check_object_flag_0(self, schema: str, brex_violations: dict, root: any, value: any, xml_text: str = None,
                              deep_copy_nodes: bool = False):
        if value['contextRules'] == schema or value['contextRules'] == "":
            if self._saxon:
                with PySaxonProcessor(license=False) as proc:
                    xp = proc.new_xpath_processor()
                    for prefix, uri in value.get('namespaces', NS_DICT).items():
                        if prefix:
                            xp.declare_namespace(prefix, uri)
                    if xml_text is not None:
                        node = proc.parse_xml(xml_text=xml_text)
                    else:
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
                                            'NodeXpath': None,
                                            'Object': None,
                                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                                            'BrSeverityLevel': value.get('brSeverityLevel'),
                                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                                        )
                    proc.exception_clear()
            else:
                try:
                    selector = elementpath.Selector(value['xpath'], namespaces=value.get('namespaces', NS_DICT))
                    result, nodes = self._select_with_nodes(selector, root)
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
                            'NodeXpath': None,
                            'Object': None,
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                        )
                else:
                    for idx, element in enumerate(result):
                        node = nodes[idx] if nodes else None
                        if ' and ' in value['xpath']:
                            line_no = "(Origin traced back to multiple lines -> Interpret XPath)"
                        else:
                            line_no = self._node_line_number(node)
                            if line_no is None:
                                line_no = "x"
                        node_xpath, node_copy = self._node_xpath_and_copy(node, deep_copy_nodes)
                        brex_violations[value["Brex"]]['0'].append({
                            'Line': line_no,
                            'Description': value["objectUse"],
                            'Xpath': value['xpath'],
                            'NodeXpath': node_xpath,
                            'Object': node_copy,
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                        )
        return brex_violations

    def _check_object_flag_1(self, schema: str, brex_violations: dict, root: any, value: any,
                              deep_copy_nodes: bool = False):
        if value['contextRules'] == schema or value['contextRules'] == "":
            try:
                selector = elementpath.Selector(value['xpath'], namespaces=value.get('namespaces', NS_DICT))
                result, nodes = self._select_with_nodes(selector, root)
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
                            'NodeXpath': None,
                            'Object': None,
                            'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                            'BrSeverityLevel': value.get('brSeverityLevel'),
                            'Fail': self._is_severity_failure(value.get('brSeverityLevel'))}
                            )
            elif not isinstance(result, (bool, str, int, float)) and (
                    value["values_allowed"] or value["regex_allowed"] or value["ranges_allowed"]):
                brex_violations[value["Brex"]]['2'].extend(
                    self._check_object_values(value, result, nodes, deep_copy_nodes))
        return brex_violations

    def _check_object_values(self, value: any, elements: any, nodes: any = None,
                              deep_copy_nodes: bool = False) -> list:
        """Check a set of matched nodes against a rule's `objectValue` children.

        Shared by any flag whose matched nodes must additionally satisfy a value
        constraint (port of `check_objects_values`, `s1kd-brexcheck.c:275-304`,
        which applies value checking to a rule's matched node-set regardless of
        `allowedObjectFlag`). Ref §3.8.

        Args:
            value (any): rule dict from `_show_rules`, carrying `values_allowed`
                / `regex_allowed` / `ranges_allowed`
            elements (any): node-set matched by `value['xpath']`
            nodes (any): raw XPath nodes aligned with `elements`, from
                `_select_with_nodes`, used to compute `NodeXpath`/`Object`
                (categories D2/D3); `None` when no such alignment is available
            deep_copy_nodes (bool): copy the full subtree instead of just the
                element's own tag and attributes, see `_node_xpath_and_copy`

        Returns:
            list: one violation dict per element whose value matches none of the
                allowed values, patterns or ranges
        """
        violations = []
        for idx, element in enumerate(elements):
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
                node = nodes[idx] if nodes else None
                if (r'] and ' or r'and \[' or r'] and \[' or r'\) and' or r'and \(' or r'\) and \(') in value['xpath']:
                    line_no = "(Origin traced back to multiple lines -> Read XPath)"
                else:
                    line_no = self._node_line_number(node)
                    if line_no is None:
                        line_no = "x"
                node_xpath, node_copy = self._node_xpath_and_copy(node, deep_copy_nodes)
                violations.append({
                    'Line': line_no,
                    'Description': f'Element/Attribute ({element_value}) did not match the object values.',
                    'Xpath': value['xpath'],
                    'NodeXpath': node_xpath,
                    'Object': node_copy,
                    'Single Values': [value["values_allowed"]],
                    'Pattern Values': [value["regex_allowed"]],
                    'Range Values': [value["ranges_allowed"]],
                    'ObjectUse': value["objectUse"],
                    'BrDecisionIdentNumber': value.get('brDecisionIdentNumber'),
                    'BrSeverityLevel': value.get('brSeverityLevel'),
                    'Fail': self._is_severity_failure(value.get('brSeverityLevel'))})
        return violations

    def _check_object_flag_2(self, schema: str, brex_violations: dict, root: any, value: any,
                              deep_copy_nodes: bool = False):
        if ('values_allowed' in value or 'regex_allowed' in value or 'ranges_allowed' in value) and (value['contextRules'] == schema or value['contextRules'] == ""):
            try:
                selector = elementpath.Selector(value['xpath'], namespaces=value.get('namespaces', NS_DICT))
                result, nodes = self._select_with_nodes(selector, root)
            except elementpath.ElementPathError as e:
                brex_violations[value["Brex"]]['xpathError'].append({
                    'Description': value["objectUse"],
                    'Xpath': value['xpath'],
                    'Error': str(e),
                    'BrDecisionIdentNumber': value.get('brDecisionIdentNumber')}
                )
                return brex_violations
            if type(result) is not bool:
                brex_violations[value["Brex"]]['2'].extend(
                    self._check_object_values(value, result, nodes, deep_copy_nodes))
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
            brex_tree = self._parse_xml_text(brex_content)
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
            brex_tree = self._parse_xml_text(brex_content)
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

    def _remove_deleted_elements(self, node: any) -> None:
        """Recursively drop elements marked as deleted from a parsed tree.

        Port of `rem_delete_nodes`/`rem_delete_elems` (`s1kd_tools.c:1054-1088`),
        `s1kd-brexcheck`'s `-^`/`--remove-deleted` option: an element carrying
        `@changeType="delete"` (or the legacy `@change="delete"` spelling the C
        original also checks) is removed along with its whole subtree before any
        rule is checked, so content staged for deletion in a change-marked
        revision does not trigger BREX violations. Children are only visited when
        the element itself is kept, matching the C original.

        Args:
            node (any): element to inspect, e.g. the checked document's root
        """
        change = node.get('change', node.get('changeType'))
        if change == 'delete':
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
            return
        for child in list(node):
            self._remove_deleted_elements(child)

    def _check_rules(self, debug: bool = False, include_tqdm: bool = False, sns_mode: str = "normal",
                      remove_deleted: bool = False, deep_copy_nodes: bool = False) -> dict:
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
            remove_deleted (bool): equivalent to `s1kd-brexcheck -^`/`--remove-deleted`;
                drop elements marked `@changeType="delete"` (see `_remove_deleted_elements`)
                before every check (content rules, SNS, notations)
            deep_copy_nodes (bool): equivalent to `-8`/`--deep-copy-nodes`; the `Object`
                field of every content-rule violation record holds a full recursive copy
                of the violating element instead of just its own tag and attributes
                (see `_node_xpath_and_copy`)

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
        root = self._parse_xml_file(self._xml_path)

        xml_text = None
        if remove_deleted:
            self._remove_deleted_elements(root.getroot())
            xml_text = etree.tostring(root, encoding="unicode")

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
            content_rules = self._show_rules(brex, schema=schema, debug=debug)
            all_content_rules += content_rules

        if debug:
            with open(clean_path(join(expanduser("~/Desktop"), "All_content_rules.txt")), 'w', encoding="utf-8") as _:
                for rule in all_content_rules:
                    _.write(str(rule) + "\n")
        container = tqdm(all_content_rules) if include_tqdm else all_content_rules
        for value in container:
            if value["ObjectFlag"] == '0':
                brex_violations_dict |= self._check_object_flag_0(
                    schema, brex_violations_dict, root, value, xml_text, deep_copy_nodes)
            if value["ObjectFlag"] == '1':
                brex_violations_dict |= self._check_object_flag_1(
                    schema, brex_violations_dict, root, value, deep_copy_nodes)
            has_values = value["values_allowed"] != [] or value["regex_allowed"] != [] or value["ranges_allowed"] != []
            # S1000D <= 3.0 rules commonly omit @objappl entirely for a
            # value-only constraint (no presence/absence semantics); s1kd's
            # is_invalid falls through to the value check in that case too.
            if has_values and value["ObjectFlag"] in ('2', None):
                brex_violations_dict |= self._check_object_flag_2(
                    schema, brex_violations_dict, root, value, deep_copy_nodes)
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
    
    def validate(self, debug: bool = False, include_tqdm: bool = False, sns_mode: str = "normal",
                 remove_deleted: bool = False, deep_copy_nodes: bool = False) -> dict:
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
            remove_deleted (bool): equivalent to `s1kd-brexcheck -^`/`--remove-deleted`;
                drop elements marked `@changeType="delete"` before checking. See
                `_remove_deleted_elements`.
            deep_copy_nodes (bool): equivalent to `-8`/`--deep-copy-nodes`; every
                content-rule violation's `Object` field holds a full recursive
                copy of the violating element (all descendants) instead of just
                its own tag and attributes. See `_node_xpath_and_copy`.

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
                xml_path = join(self._xml_dir, _xml)
                if self._ignore_empty and not self._is_valid_xml_file(xml_path):
                    continue
                self.set_xml(xml_path)
                self._init_brex_list()
                result = self._check_rules(debug=debug, include_tqdm=include_tqdm, sns_mode=sns_mode,
                                            remove_deleted=remove_deleted, deep_copy_nodes=deep_copy_nodes)
                result["Summary"] = self._append_summary(result)
                results[_xml] = result
                self._brex_list = initial_brex_list if had_explicit_brex_list else (None, None)
                self._brex_dir_path = initial_brex_dir_path if had_explicit_brex_dir_path else (None, None)
            if debug:
                with open(clean_path(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_dir)}.json')), 'w', encoding="utf-8") as _:
                    dump(results, _, indent=4)
            return results
        else:
            if self._ignore_empty and not self._is_valid_xml_file(self._xml_path):
                result = {"Skipped": True, "Summary": "Skipped (empty or non-XML file)"}
                if debug:
                    with open(clean_path(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_path)}.json')), 'w', encoding="utf-8") as _:
                        dump(result, _, indent=4)
                return result
            self._init_brex_list()
            result = self._check_rules(debug=debug, sns_mode=sns_mode, remove_deleted=remove_deleted,
                                        deep_copy_nodes=deep_copy_nodes)
            summary = self._append_summary(result)
            result["Summary"] = summary
            if debug:
                with open(clean_path(join(expanduser("~/Desktop"), f'Errors_{basename(self._xml_path)}.json')), 'w', encoding="utf-8") as _:
                    dump(result, _, indent=4)
            return result

