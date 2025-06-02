"""Module for S1000D function
See s1000d_xml_samples.xml for examples of S1000D XML
"""

from os import listdir
from os import walk
from os.path import join
from os.path import splitext
from os.path import isdir
from os.path import isfile
from os.path import sep
from os.path import basename
from os.path import dirname
from os.path import expanduser

import sys

from re import search
from re import findall

from requests import get

from json import dump

from lxml import etree

sys.path.insert(0, join(dirname(dirname(__file__)), 'txt'))
from txt import add_leading

from .xml_processing import delete_first_line
from .xml_processing import get_xml_attribute
from .xml_processing import set_xml_attribute
from .xml_processing import get_schema_from_xml
from .xml_processing import linearize_xml
from .constants import S1000D_VERSION_REGEX
from .constants import DM_REF_REGEX
from .constants import DELIVERY_LIST_ITEM_REGEX
from .constants import OLD_TO_NEW

def get_s1000d_version(xml: str) -> float:
    """Returns the S1000D version of the XML file

    Args:
        xml (str): File path to XML file

    Returns:
        str: A string containing the S1000D version or None if not found
    """
    if isfile(xml):
        with open(xml, "r", encoding="utf-8") as _:
            while (line := _.readline().rstrip()):
                if "s1000d" in line:
                    if search(S1000D_VERSION_REGEX, line):
                        return float(search(S1000D_VERSION_REGEX, line).group(2).replace("-", "."))  # we convert 4-0 to 4.0
                    else:
                        raise Exception(f"Unknown S1000D version pattern: {line}")
    elif search(S1000D_VERSION_REGEX, xml):
        return float(search(S1000D_VERSION_REGEX, xml).group(2).replace("-", "."))  # we convert 4-0 to 4.0
    raise Exception(f"Could not find version for : {xml}")

def get_references(directory: any, json_dump: bool = False, http_mode: bool = False) -> dict:
    """Get all data module references from inside all S1000D data modules in a directory

    Args:
        dir (any): Path to directory containing S1000D data modules or list of http links
        json_dump (bool, optional): Dump the result to a JSON file. Defaults to False.
        http_mode (bool, optional): If True, will get the references from the HTTP server. Defaults to False.

    Returns:
        dict: Dictionary containing all references for each data module
    """
    references = {}
    if not http_mode:
        for document in listdir(directory):
            if splitext(document)[-1] in [".xml", ".XML"]:
                s1000d_version = get_s1000d_version(join(directory, document))
                if s1000d_version is not None:
                    references[join(directory, document)] = get_s1000d_refs(join(directory, document), s1000d_version)
                else:
                    raise Exception(f"Could not find version for {join(directory, document)}.\nPlease check the file.")
    else:
        for document in directory:
            s1000d_version = get_s1000d_version(str(get(document).content))
            if s1000d_version is not None:
                references[document] = get_s1000d_refs(str(get(document).content), s1000d_version)
            else:
                raise Exception(f"Could not find version for {document}.\nPlease check the file.")
    if json_dump and not http_mode:
        parent_dir = directory.split(sep)[-1]
        with open(join(directory, f"{parent_dir}_references.json"), "w", encoding="utf-8") as _:
            dump(references, _, indent=4)
    return references


def get_s1000d_refs(xml: str, s1000d_version: str) -> list:
    """get_s1000d_refs

    Args:
        xml (str): File path to XML file
        s1000d_version (str): Version of S1000D

    Returns:
        list: List of references
        Note: Some references might not have issue_info, dm_title or issue_date
        For some dm_title the techName or infoName might be empty or None
    """
    if s1000d_version >= 4.0:  # if s1000d_version in ("4-0", "4-1", "4-2", "5-0"):
        return get_4plus_refs(xml)
    if s1000d_version <= 3.0:  # if s1000d_version in ("2-3", "3-0"):
        return get_2and3_refs(xml)

    return f"Unknown S1000D version: {s1000d_version}"

def get_4plus_refs(xml: str) -> list:
    """Get data module references from inside a S1000D 4.2 data module

    Args:
        xml (str): File path to XML file

    Returns:
        list: List of references
        Note: Some references might not have issue_info, dm_title or issue_date
        For some dm_title the techName or infoName might be empty or None
    """
    refs = []

    if isfile(xml):
        with open(xml, "r", encoding="utf-8") as _:
            xml = _.read().replace("\n", " ").replace("> <", "><")
    elif "www.s1000d.org" in str(xml):
        xml = str(xml).replace("\r\n", " ").replace("\n", " ").replace("\t", "").replace("> <", "><")
        xml = str(xml).replace(r"\r\n", " ").replace(r"\n", " ").replace(r"\t", "").replace("> <", "><")

    match_found = search(DM_REF_REGEX, xml)
    if match_found:
        for match in findall(DM_REF_REGEX, xml):
            try:
                parsed_ref = etree.fromstring(match)
            except etree.XMLSyntaxError:
                print(f"Error parsing: {xml}")
                print(f"Error parsing: {match}")

            dm_code = dict(parsed_ref.xpath("//dmCode")[0].attrib)
            issue_info = dict(parsed_ref.xpath("//issueInfo")[0].attrib) if parsed_ref.xpath("//issueInfo") else {
                "inWork": "",
                "issueNumber": ""
            }
            dm_title = {
                "techName": parsed_ref.xpath("//techName")[0].text,
                "infoName": parsed_ref.xpath("//infoName")[0].text
            } if parsed_ref.xpath("//techName") else {
                "techName": "",
                "infoName": ""
            }

            issue_date = dict(parsed_ref.xpath("//issueDate")[0].attrib) if parsed_ref.xpath("//issueDate") else {
                "day": "",
                "month": "",
                "year": ""
            }

            refs += [dm_code | issue_info | dm_title | issue_date]

    return refs

def get_2and3_refs(xml: str) -> list:
    """Get data module references from inside a S1000D 2.3 data module

    Args:
        xml (str): File path to XML file

    Returns:
        list: List of references
        Note: Some references might not have issue_info, dm_title or issue_date
        For some dm_title the techName or infoName might be empty or None
    """
    refs = []

    if isfile(xml):
        with open(xml, "r", encoding="utf-8") as _:
            xml = _.read().replace("\n", " ").replace("> <", "><")
    elif "www.s1000d.org" in str(xml):
        xml = str(xml).replace("\r\n", " ").replace("\n", " ").replace("\t", "").replace("> <", "><")
        xml = str(xml).replace(r"\r\n", " ").replace(r"\n", " ").replace(r"\t", "").replace("> <", "><")

    match_found = search(DM_REF_REGEX, xml)
    if match_found:
        for match in findall(DM_REF_REGEX, xml):
            parsed_ref = etree.fromstring(match)
            dm_code = {OLD_TO_NEW[child.tag]: child.text for child in parsed_ref.xpath("//avee")[0]}
            refs += [dm_code]
    return refs

def get_brex_ref(xml: str, to_string: bool = False) -> dict:
    """Get BREX reference from inside a S1000D data module

    Args:
        xml (str): File path to XML file

    Returns:
        dict: BREX reference
    """
    if get_s1000d_version(xml) is not None:
        for ref in get_s1000d_refs(xml, get_s1000d_version(xml)):
            if ref["infoCode"] == "022":
                if to_string:
                    return ref_dict_to_str(ref)
                return ref
    else:
        raise Exception(f"Could not find version for {xml}.\nPlease check the file.")
    return None

def ref_dict_to_str(ref: dict) -> str:
    """Converts a reference dictionary to a datamodule filename string

    Args:
        ref (dict): Reference dictionary

    Returns:
        str: Data module filename string
    """
    str_ref = f"DMC-{ref['modelIdentCode']}\
                -{ref['systemDiffCode']}\
                -{ref['systemCode']}\
                -{ref['subSystemCode']}\
                {ref['subSubSystemCode']}\
                -{ref['assyCode']}\
                -{ref['disassyCode']}\
                {ref['disassyCodeVariant']}\
                -{ref['infoCode']}\
                {ref['infoCodeVariant']}\
                -{ref['itemLocationCode']}".replace(" ", "")
    try:
        if ref["issueNumber"] != "" and ref["issueNumber"] is not None:
            str_ref += f"_{ref['issueNumber']}"
        if ref["inWork"] != "" and ref["inWork"] is not None:
            str_ref += f"-{ref['inWork']}"
    except KeyError:
        pass

    return str_ref

def find_document_by_reference(filename_part: str, directory: str, extension: str = ".xml") -> str:
    """Finds the full path of a document based on a filename part and the extension

    Args:
        filename_part (str): Filename part
        directory (str): Directory to search in
        extension (str): extension

    Returns:
        str: Full path to document
    """
    for root, _, files in walk(directory):
        for file_ in files:
            if filename_part in file_ and file_.lower().endswith(extension.lower()):
                return join(root, file_)

def get_dm_codes_from_dir(directory: str, json_dump: bool = False) -> dict:
    dm_codes = {}
    for root, _, files in walk(directory):
        for file_ in files:
            if file_.lower().endswith(".xml"):
                from_filename = get_dm_code_from_filename(file_)
                from_xml = get_dm_code_from_xml(join(root, file_))
                dm_codes[join(root, file_)] = {
                    "from_filename": from_filename,
                    "from_xml": from_xml,
                    "are_identical": from_filename == from_xml
                }

    if json_dump:
        parent_dir = directory.split(sep)[-1]
        with open(join(directory, f"{parent_dir}_dmcodes.json"), "w", encoding="utf-8") as _:
            dump(dm_codes, _, indent=4)

    return dm_codes

def get_dm_code_from_filename(filename: str) -> dict:
    """Get code (dmCode / pmCode / ddnCode) from filename (without path)

    Args:
        filename (str): Filename

    Returns:
        dict: code
    """
    filename_list = filename.split("-")

    if "DDN" in filename:
        return {
            "modelIdentCode": filename_list[1],
            "receiverIdent": filename_list[2],
            "senderIdent": filename_list[3],
            "seqNumber": filename_list[5].split(".")[0],
            "yearOfDataIssue": filename_list[4]
        }
    elif "PMC" in filename:
        return {
            "modelIdentCode": filename_list[1],
            "pmIssuer": filename_list[2],
            "pmNumber": filename_list[3],
            "pmVolume": filename_list[4].split("_")[0],
        }
    elif "DMC" in filename:
        return {
            "modelIdentCode": filename_list[1],
            "systemDiffCode": filename_list[2],
            "systemCode": filename_list[3],
            "subSystemCode": filename_list[4][0],
            "subSubSystemCode": filename_list[4][1],
            "assyCode": filename_list[5],
            "disassyCode": filename_list[6][0:2],
            "disassyCodeVariant": filename_list[6][2],
            "infoCode": filename_list[7][0:3],
            "infoCodeVariant": filename_list[7][3],
            "itemLocationCode": filename_list[8].split("_")[0]
        }
    return None

def get_dm_code_from_xml(xml: str) -> dict:
    """Get code (dmCode / pmCode / ddnCode) from XML

    Args:
        xml (str): Path of the XML file

    Returns:
        dict: dmCode as a dictionary
    """

    code = None
    xml_filename = basename(xml)

    s1000d_version = get_s1000d_version(xml)

    with open(xml, "r", encoding="utf-8") as _:
        xml = delete_first_line(_.read().replace("\n", " ").replace("> <", "><"))
    if "DDN" in xml_filename:
        code = dict(
            etree.fromstring(xml).xpath("//identAndStatusSection/ddnAddress/ddnIdent/ddnCode")[0].attrib
        ) if s1000d_version >= 4.0 else {
            OLD_TO_NEW[child.tag]: child.text for child in etree.fromstring(xml).xpath("//idstatus/ddnaddres/ddnc")[0]
        }
    elif "PMC" in xml_filename:
        code = dict(
            etree.fromstring(xml).xpath("//identAndStatusSection/pmAddress/pmIdent/pmCode")[0].attrib
        ) if s1000d_version >= 4.0 else {
            OLD_TO_NEW[child.tag]: child.text for child in etree.fromstring(xml).xpath("//idstatus/pmaddres/pmc")[0]
        }
    elif "DMC" in xml_filename:
        code = dict(
            etree.fromstring(xml).xpath("//identAndStatusSection/dmAddress/dmIdent/dmCode")[0].attrib
        ) if s1000d_version >= 4.0 else {
            OLD_TO_NEW[child.tag]: child.text for child in etree.fromstring(xml).xpath("//idstatus/dmaddres/dmc/avee")[0]
        }

    return code

def validate_references(directory: str, json_dump: bool = False) -> dict:
    """Validates the references of all documents in a directory

        Args:
            directory (str): Directory to search in
            json_dump (bool): Dump the results to a json file

        Returns:
            dict: Results
    """

    dm_codes_dict = get_dm_codes_from_dir(directory, True)
    references = get_references(directory, True)
    reference_validation = {}
    for key in references:
        reference_validation[key] = {}
        for ref in references[key]:
            ref_filepath = find_document_by_reference(
                ref_dict_to_str(ref_dict_to_dm_code_dict(ref)),
                directory
            )
            dm_code_message = "" if ref_filepath is None else \
                ", but the DM code does not match the filename" if not dm_codes_dict[ref_filepath]["are_identical"] \
                else ""
            filename_message = "Resolves to a valid document" if ref_filepath else "Does not resolve to a valid document"

            error = "Error: " if "but" in dm_code_message or "Does not" in filename_message else ""
            reference_validation[key][ref_dict_to_str(ref_dict_to_dm_code_dict(ref))] = error
            reference_validation[key][ref_dict_to_str(ref_dict_to_dm_code_dict(ref))] += \
                "Resolves to a valid document" + dm_code_message if ref_filepath else "Does not resolve to a valid document"

    if json_dump:
        parent_dir = directory.split(sep)[-1]
        with open(join(directory, f"{parent_dir}_reference_validation.json"), "w", encoding="utf-8") as _:
            dump(reference_validation, _, indent=4)

    return reference_validation

def ref_dict_to_dm_code_dict(ref: dict) -> dict:
    """Converts a reference dictionary to a dm_code dictionary

    Args:
        ref (dict): Reference dictionary

    Returns:
        dict: DM code dictionary
    """

    if "receiverIdent" in ref:  # DDN
        return ref
    elif "pmIssuer" in ref:  # PMC
        return ref
    elif "infoCode" in ref:  # DMC
        return {
            "modelIdentCode": ref["modelIdentCode"],
            "systemDiffCode": ref["systemDiffCode"],
            "systemCode": ref["systemCode"],
            "subSystemCode": ref["subSystemCode"],
            "subSubSystemCode": ref["subSubSystemCode"],
            "assyCode": ref["assyCode"],
            "disassyCode": ref["disassyCode"],
            "disassyCodeVariant": ref["disassyCodeVariant"],
            "infoCode": ref["infoCode"],
            "infoCodeVariant": ref["infoCodeVariant"],
            "itemLocationCode": ref["itemLocationCode"]
        }

def ref_dict_to_xml(ref: dict) -> str:
    """ref_dict_to_xml converts a reference dictionary to an XML string

    Args:
        ref (dict): Reference dictionary

    Returns:
        str: XML string
    """
    xml_ref = '<dmRef><dmRefIdent><dmCode '
    attributes = ''
    for key, value in ref.items():
        attributes += f'{key}="{value}" '
    xml_ref += attributes[:-1] + '/></dmRefIdent></dmRef>'

    return xml_ref

def get_ddn(directory: str) -> str:
    """Find the DDN in a directory

    Args:
        directory (str): Directory to search in

    Returns:
        str: Full path of the DDN
    """
    for root, _, files in walk(directory):
        for file_ in files:
            if "DDN" in file_:
                return join(root, file_)

def validate_ddn(directory: str, json_dump: bool = False) -> dict:
    """Validates the DDN of all documents in a directory

        Args:
            directory (str): Directory to search in
            json_dump (bool): Dump the results to a json file

        Returns:
            dict: Results
    """

    results = {}

    with open(get_ddn(directory), "r", encoding="utf-8") as _:
        xml = delete_first_line(_.read().replace("\n", " ").replace("> <", "><"))
    dir_list = listdir(directory)
    dcn_items = []
    if search(DELIVERY_LIST_ITEM_REGEX, xml):
        for fname in findall(DELIVERY_LIST_ITEM_REGEX, xml):
            results[fname[1]] = []
            dcn_items.append(fname[1])
            if f"{fname[9]}-{fname[7]}" not in fname[1]:  # Issue info is not in the filename
                results[fname[1]].append("Error: Does not have the correct issue info in its filename")
            if fname[4] not in fname[1]:  # ECN
                results[fname[1]].append("Error: Does not have the correct ECN in its filename")
            if fname[1] not in dir_list:  # File exists
                results[fname[1]].append("Error: Is not in the directory")

            if len(results[fname[1]]) == 0:
                results[fname[1]].append("Valid")
    results["Files not in DDN"] = []
    if any(file_ not in dcn_items for file_ in dir_list):
        for file_ in dir_list:
            if file_ not in dcn_items:
                results["Files not in DDN"].append(file_)

    if json_dump:
        with open(join(directory, "ddn_validation.json"), "w", encoding="utf-8") as _:
            dump(results, _, indent=4)

    return results

def increase_issue_number(dmodule: str) -> None:
    """Increases the issue number of a document module
        Args:
            dmodule (str): Document module to increase the issue number of

    TODO: Not to be used, use set_xml_attribute instead
    # INFO: Probably there is no need to have specific functions for each xml value or attribute
    # We could just use the general functions get_xml_value and get_xml_attribute and set_xml_value and set_xml_attribute
    """
    set_xml_attribute(
        dmodule,
        "./identAndStatusSection/dmAddress/dmIdent/issueInfo",
        "issueNumber",
        add_leading(str(int(get_xml_attribute(
            dmodule,
            "./identAndStatusSection/dmAddress/dmIdent/issueInfo",
            "issueNumber"
        )) + 1), "0", 3))


def set_inwork(dmodule: str, in_work: str) -> None:
    """Sets the inWork attribute of a document module

    Args:
        dmodule (str): Document module to increase the issue number of
        in_work (str): Value to set the inWork attribute to

    TODO: Not to be used, use set_xml_attribute instead
    # INFO: Probably there is no need to have specific functions for each xml value or attribute
    # We could just use the general functions get_xml_value and get_xml_attribute and set_xml_value and set_xml_attribute
    """
    set_xml_attribute(
        dmodule,
        "./identAndStatusSection/dmAddress/dmIdent/issueInfo",
        "inWork",
        in_work
    )

def read_dmodule(dmodule: str, json_dump: bool = False, show: bool = False) -> dict:
    """Reads a S1000d v4.0+ procedure or description and returns a dictionary with the procedure steps that have an id for reference

    Args:
        dmodule (str): Data module to read (xml filepath)
        json_dump (bool, optional): If it should dump the dictionary to a json stored on your Desktop. Defaults to False.
        show (bool, optional): If it should print the procedure to console. Defaults to False.

    Returns:
        dict: Dictionary with the procedure steps that have an id for reference
    """
    with open(dmodule, "r", encoding="utf-8") as _:
        xml = linearize_xml(delete_first_line(_.read()))
        schema = get_schema_from_xml(xml).split("/")[-1].split(".")[0]
    if "proced" == schema:
        return read_procedure(dmodule, json_dump, show)
    if "descript" == schema:
        return read_description(dmodule, json_dump, show)
    return None
    
    
def read_procedure(dmodule: str, json_dump: bool = False, show: bool = False) -> dict:
    """Reads a S1000d v4.0+ procedure and returns a dictionary with the procedure steps that have an id for reference

    Args:
        dmodule (str): Data module to read (xml filepath)
        json_dump (bool, optional): If it should dump the dictionary to a json stored on your Desktop. Defaults to False.
        show (bool, optional): If it should print the procedure to console. Defaults to False.

    Returns:
        dict: Dictionary with the procedure steps that have an id for reference
    """
    references = {}
    if isfile(dmodule):
        xml_tree = etree.parse(dmodule)
    else:
        xml_tree = etree.fromstring(dmodule)
    for ind, step in enumerate(xml_tree.xpath(f"//mainProcedure/proceduralStep")):
        for ind2, child in enumerate(step.xpath("*")):
            if child.tag == "title" or child.tag == "para":
                if show:
                    print(f"{ind + 1}. {child.text}")  # 1. Title
                if child.getparent().tag == "proceduralStep":
                    if "id" in child.getparent().attrib:
                        references[f"{ind + 1}. {child.text}"] = {
                            "id": child.getparent().attrib["id"],
                            "number": f"{ind + 1}.",
                            "content": child.text
                        }
            if child.tag == "proceduralStep":
                for ind3, step2 in enumerate(child.xpath("*")):
                    if step2.tag == "title" or step2.tag == "para":
                        if show:
                            print(f"{chr(ind2 + 64)}. {step2.text}")  # A. Title
                        if step2.getparent().tag == "proceduralStep":
                            if "id" in step2.getparent().attrib:
                                references[f"{ind + 1}.{chr(ind2 + 64)}. {step2.text}"] = {
                                    "id": step2.getparent().attrib["id"],
                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.",
                                    "content": step2.text
                                }
                    if step2.tag == "proceduralStep":
                        for ind4, step3 in enumerate(step2.xpath("*")):
                            if step3.tag == "title" or step3.tag == "para":
                                if show:
                                    print(f"({ind3}) {step3.text}")  # (1) Title
                                if step3.getparent().tag == "proceduralStep":
                                    if "id" in step3.getparent().attrib:
                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}) {step3.text}"] = {
                                            "id": step3.getparent().attrib["id"],
                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3})",
                                            "content": step3.text
                                        }
                            if step3.tag == "proceduralStep":
                                for ind5, step4 in enumerate(step3.xpath("*")):
                                    if step4.tag == "title" or step4.tag == "para":
                                        if show:
                                            print(f"({chr(ind4 + 96)}) {step4.text}")  # (a) Title
                                        if step4.getparent().tag == "proceduralStep":
                                            if "id" in step4.getparent().attrib:
                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}) {step4.text}"] = {
                                                    "id": step4.getparent().attrib["id"],
                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)})",
                                                    "content": step4.text
                                                }
                                    if step4.tag == "proceduralStep":
                                        for ind6, step5 in enumerate(step4.xpath("*")):
                                            if step5.tag == "title" or step5.tag == "para":
                                                if show:
                                                    print(f"{ind5}. {step5.text}")  # 1. Title
                                                if step5.getparent().tag == "proceduralStep":
                                                    if "id" in step5.getparent().attrib:
                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5} {step5.text}"] = {
                                                            "id": step5.getparent().attrib["id"],
                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}",
                                                            "content": step5.text
                                                        }
                                            if step5.tag == "proceduralStep":
                                                for ind7, step6 in enumerate(step5.xpath("*")):
                                                    if step6.tag == "title" or step6.tag == "para":
                                                        if show:
                                                            print(f"{chr(ind6 + 96)}. {step6.text}")  # a. Title
                                                        if step6.getparent().tag == "proceduralStep":
                                                            if "id" in step6.getparent().attrib:
                                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}. {step6.text}"] = {
                                                                    "id": step6.getparent().attrib["id"],
                                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.",
                                                                    "content": step6.text
                                                                }
                                                    if step6.tag == "proceduralStep":
                                                        for ind8, step7 in enumerate(step6.xpath("*")):
                                                            if step7.tag == "title" or step7.tag == "para":
                                                                if show:
                                                                    print(f"({ind7}) {step7.text}")  # (1) Title
                                                                if step7.getparent().tag == "proceduralStep":
                                                                    if "id" in step7.getparent().attrib:
                                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}) {step7.text}"] = {
                                                                            "id": step7.getparent().attrib["id"],
                                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7})",
                                                                            "content": step7.text
                                                                        }
                                                            if step7.tag == "proceduralStep":
                                                                for ind9, step8 in enumerate(step7.xpath("*")):
                                                                    if step8.tag == "title" or step8.tag == "para":
                                                                        if show:
                                                                            print(f"({chr(ind8 + 96)}) {step8.text}")  # (a) Title
                                                                        if step8.getparent().tag == "proceduralStep":
                                                                            if "id" in step8.getparent().attrib:
                                                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).({chr(ind8 + 96)}) {step8.text}"] = {
                                                                                    "id": step8.getparent().attrib["id"],
                                                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).({chr(ind8 + 96)})",
                                                                                    "content": step8.text
                                                                                }
                                                                    if step8.tag == "proceduralStep":
                                                                        for ind10, step9 in enumerate(step8.xpath("*")):
                                                                            if step9.tag == "title" or step9.tag == "para":
                                                                                if show:
                                                                                    print(f"{ind9}. {step9.text}")  # 1. Title
                                                                                if step9.getparent().tag == "proceduralStep":
                                                                                    if "id" in step9.getparent().attrib:
                                                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).{chr(ind8 + 96)}.{ind9}. {step9.text}"] = {
                                                                                            "id": step9.getparent().attrib["id"],
                                                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).{chr(ind8 + 96)}.{ind9}.",
                                                                                            "content": step9.text
                                                                                        }

    if json_dump:
        with open(join(expanduser("~/Desktop"), "references.json"), "w", encoding="utf-8") as f:
            dump(references, f, indent=4)
    return references

def read_description(dmodule: str, json_dump: bool = False, show: bool = False) -> dict:
    """Reads a S1000d v4.0+ description and returns a dictionary with the description steps that have an id for reference

    Args:
        dmodule (str): Data module to read (xml filepath)
        json_dump (bool, optional): If it should dump the dictionary to a json stored on your Desktop. Defaults to False.
        show (bool, optional): If it should print the description to console. Defaults to False.

    Returns:
        dict: Dictionary with the description steps that have an id for reference
    """
    # TODO: Fix the numbering - it's not correct. Create cases, either with strict rules or flexible rules.
    # RECOMENDATION: Use strict rules. If title and para on same level then same numbering should be used. 1. TITLE 2. PARA
    # Don't use 1. TITLE, a. PARA. To achieve this, the author should restructure the description.
    references = {}    
    xml_tree = etree.parse(dmodule)
    for ind, step in enumerate(xml_tree.xpath(f"//description/levelledPara")):
        for ind2, child in enumerate(step.xpath("*")):
            if child.tag == "title" or child.tag == "para":
                if show:
                    print(f"{ind + 1}. {child.text}")  # 1. Title
                if child.getparent().tag == "levelledPara":
                    if "id" in child.getparent().attrib:
                        references[f"{ind + 1}. {child.text}"] = {
                            "id": child.getparent().attrib["id"],
                            "number": f"{ind + 1}.",
                            "content": child.text
                        }
            if child.tag == "levelledPara":
                for ind3, step2 in enumerate(child.xpath("*")):
                    if step2.tag == "title" or step2.tag == "para":
                        if show:
                            print(f"{chr(ind2 + 64)}. {step2.text}")  # A. Title
                        if step2.getparent().tag == "levelledPara":
                            if "id" in step2.getparent().attrib:
                                references[f"{ind + 1}.{chr(ind2 + 64)}. {step2.text}"] = {
                                    "id": step2.getparent().attrib["id"],
                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.",
                                    "content": step2.text
                                }
                    if step2.tag == "levelledPara":
                        for ind4, step3 in enumerate(step2.xpath("*")):
                            if step3.tag == "title" or step3.tag == "para":
                                if show:
                                    print(f"({ind3}) {step3.text}")  # (1) Title
                                if step3.getparent().tag == "levelledPara":
                                    if "id" in step3.getparent().attrib:
                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}) {step3.text}"] = {
                                            "id": step3.getparent().attrib["id"],
                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3})",
                                            "content": step3.text
                                        }
                            if step3.tag == "levelledPara":
                                for ind5, step4 in enumerate(step3.xpath("*")):
                                    if step4.tag == "title" or step4.tag == "para":
                                        if show:
                                            print(f"({chr(ind4 + 96)}) {step4.text}")  # (a) Title
                                        if step4.getparent().tag == "levelledPara":
                                            if "id" in step4.getparent().attrib:
                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}) {step4.text}"] = {
                                                    "id": step4.getparent().attrib["id"],
                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)})",
                                                    "content": step4.text
                                                }
                                    if step4.tag == "levelledPara":
                                        for ind6, step5 in enumerate(step4.xpath("*")):
                                            if step5.tag == "title" or step5.tag == "para":
                                                if show:
                                                    print(f"{ind5}. {step5.text}")  # 1. Title
                                                if step5.getparent().tag == "levelledPara":
                                                    if "id" in step5.getparent().attrib:
                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5} {step5.text}"] = {
                                                            "id": step5.getparent().attrib["id"],
                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}",
                                                            "content": step5.text
                                                        }
                                            if step5.tag == "levelledPara":
                                                for ind7, step6 in enumerate(step5.xpath("*")):
                                                    if step6.tag == "title" or step6.tag == "para":
                                                        if show:
                                                            print(f"{chr(ind6 + 96)}. {step6.text}")  # a. Title
                                                        if step6.getparent().tag == "levelledPara":
                                                            if "id" in step6.getparent().attrib:
                                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}. {step6.text}"] = {
                                                                    "id": step6.getparent().attrib["id"],
                                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.",
                                                                    "content": step6.text
                                                                }
                                                    if step6.tag == "levelledPara":
                                                        for ind8, step7 in enumerate(step6.xpath("*")):
                                                            if step7.tag == "title" or step7.tag == "para":
                                                                if show:
                                                                    print(f"({ind7}) {step7.text}")  # (1) Title
                                                                if step7.getparent().tag == "levelledPara":
                                                                    if "id" in step7.getparent().attrib:
                                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}) {step7.text}"] = {
                                                                            "id": step7.getparent().attrib["id"],
                                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7})",
                                                                            "content": step7.text
                                                                        }
                                                            if step7.tag == "levelledPara":
                                                                for ind9, step8 in enumerate(step7.xpath("*")):
                                                                    if step8.tag == "title" or step8.tag == "para":
                                                                        if show:
                                                                            print(f"({chr(ind8 + 96)}) {step8.text}")  # (a) Title
                                                                        if step8.getparent().tag == "levelledPara":
                                                                            if "id" in step8.getparent().attrib:
                                                                                references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).({chr(ind8 + 96)}) {step8.text}"] = {
                                                                                    "id": step8.getparent().attrib["id"],
                                                                                    "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).({chr(ind8 + 96)})",
                                                                                    "content": step8.text
                                                                                }
                                                                    if step8.tag == "levelledPara":
                                                                        for ind10, step9 in enumerate(step8.xpath("*")):
                                                                            if step9.tag == "title" or step9.tag == "para":
                                                                                if show:
                                                                                    print(f"{ind9}. {step9.text}")  # 1. Title
                                                                                if step9.getparent().tag == "levelledPara":
                                                                                    if "id" in step9.getparent().attrib:
                                                                                        references[f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).{chr(ind8 + 96)}.{ind9}. {step9.text}"] = {
                                                                                            "id": step9.getparent().attrib["id"],
                                                                                            "number": f"{ind + 1}.{chr(ind2 + 64)}.({ind3}).({chr(ind4 + 96)}).{ind5}.{chr(ind6 + 96)}.({ind7}).{chr(ind8 + 96)}.{ind9}.",
                                                                                            "content": step9.text
                                                                                        }

    if json_dump:
        with open(join(expanduser("~/Desktop"), "references.json"), "w", encoding="utf-8") as f:
            dump(references, f, indent=4)
    return references

if __name__ == "__main__":
    # read_dmodule(
    #     r"C:\Users\munteanu\Desktop\SITEC\Publisher CMP 21-77-05\DMC-CTTAE29N-A-21-77-05-02A-018A-D_001-01_SX-US.XML",
    #     # r"C:\Users\munteanu\Desktop\SITEC\Publisher CMP 21-77-05\DMC-CTTAE29N-A-21-77-05-02A-710A-D_001-01_SX-US.XML",
    #     True,
    #     True
    # )
    validate_ddn(r"C:\Users\munteanu\Downloads\Updated_CMPs\Work_Folder_SITEC_05", True)
    # validate_references(r"C:\Users\munteanu\Downloads\Updated_CMPs\Work_Folder_SITEC_05", True)
