# Workflow
# 1. Import Class
# 		from acd import ConsumablesList
# 2. Create Instance of class
# 		instance = ConsumablesList()
# 3. Call set_xml function and pass path to xml
# 		instance.set_xml(r"")
# 4. Call set_excel function and pass path to "MASTER LOM ISSUE"
# 		instance.set_excel(r"")
# If you wish to only create the tbody:
# 5.1. Call create_table_function
#       instance.create_table_file()
# If you wish to validate an exisiting cons table:
# 5.2 Call validate_cons_table function
#       instance.validate_cons_table()

# import heartrate; heartrate.trace(browser=True)

from io import StringIO

from re import findall
from re import sub

from sys import exit as done

from os.path import join
from os.path import expanduser
from os.path import basename
from os.path import dirname
from os.path import normpath

from traceback import format_exc

from lxml import etree

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal

from typing import Tuple
# from typing import Dict  # Cosmin - Was not used

from regex import search


from .xml_processing import linearize_xml  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position
from .xml_processing import delete_first_line  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position
from .xml_processing import replace_special_characters  # noqa # pylint: disable=unused-import, import-error, ungrouped-imports, wrong-import-position
from .ataispec2200 import NoXmlSet


import pandas as pd

FILEPATH = dirname(__file__)

class NoExcelSet(Exception):
    pass

class NoOriginalTableFound(Exception):
    pass

class DictError(Exception):
    pass

class ConsumablesList():
    def __init__(self) -> None:
        self.xml_path = None
        self.excel_path = None
        self.xml_list = None
        self.main_dict = None
        self.old_compare_add = None
        self.export_path = expanduser("~/Desktop")

    def set_xml(self, xml_path: str):
        """Function with which the user can set the xml to be checked.

        Args:
            xml_path (str): xml file path.
        """
        self.xml_path = xml_path

    def set_excel(self, excel_path: str):
        """Function with which the user can set the excel where the connbr are collected.

        Args:
            excel_path (str): excel file path
        """
        self.excel_path = excel_path

    def set_export_path(self, export_path: str):
        """Function with huch the user can specify a path different to the default path (desktop),
        where to export the files the script produces.
        Args:
            export_path (str): path to where the generated files should be exported.
        """
        self.export_path = export_path

    def read_xml_file(self) -> str:
        """Function to open and read an xml file. Content is saved into variable.

        Raises:
            NoXmlSet: If path is incorrect or xml can't be found we raise an error.

        Returns:
            str: variable which contains the content of the xml.
        """
        if self.xml_path is None:
            raise NoXmlSet("No xml found to validate. Use the set_xml function to set the path of your xml \
                        before you use the validate_consumables function.".replace("                        ", ""))

        with open(self.xml_path, "r", encoding="utf-8") as _:
            xml_content = _.read()

        return xml_content

    def prepare_xml_content(self) -> str:
        """Modifies the xml content such that it is better processable.
        1. Linearizes the content (All content in first line)
        2. Deltes the first line (i.e "<?xml version="1.0" encoding="UTF-8"?>")
        3. Replaces special characters that are non unicode

        Returns:
            str: xml content with mentioned modifications.
        """
        xml_content = self.read_xml_file()
        xml_content = linearize_xml(xml_content)
        xml_content = delete_first_line(xml_content)
        xml_content = replace_special_characters(xml_content)
        return xml_content

    def replace_entities(self) -> str:
        """Adds content of inmedISOEntities to the xml to handle special characters
        for further lxml processing.

        Returns:
            str: xml content with mentioned modifications.
        """
        xml_content = self.prepare_xml_content()
        with open(join(FILEPATH, "inmedISOEntities.ent"), "r", encoding="utf-8") as _:
            entities = _.read()
        entities = linearize_xml(entities)
        xml_content = sub(r"(]>.*?<cmm)", entities + r'\1', xml_content)
        return xml_content

    def extract_pageblocks(self) -> list:
        """Function converts content of xml to an lxml etree object and uses
        xpath to extract all pageblocks and save them as elements in a list.

        Returns:
            list: Containing lxml etree objects which represent the pageblocks.
        """
        xml_content = self.replace_entities()
        xml_content = etree.parse(StringIO(xml_content))
        all_pageblocks = xml_content.xpath(r"//pgblk")
        return all_pageblocks

    def extract_consumable_code(self) -> list:
        """Function iterates trough each procedure
        and saves the connbrs in a list.

        Returns:
            list: Containing all connbrs as strings.
        """
        all_pageblocks = self.extract_pageblocks()
        con_codes = []
        for pgblk in all_pageblocks:
            check_task = pgblk.xpath(r".//task")
            for task in check_task:
                table_con = task.xpath(r"./topic[title = 'Procedure']")
                if table_con:
                    for con in table_con[0].xpath(r".//con"):
                        connbr_text = ""
                        try:
                            connbr_text = str(con.xpath("connbr")[0].text)
                        except IndexError:
                            pass
                        con_codes.append(connbr_text)
        return con_codes

    def prepare_consumables(self) -> list:
        """Function that removes all duplicate connbrs from the list
        where they are saved in and sort the elements in increasing order

        Returns:
            list: Containing all connbrs as strings with the mentioned modifications.
        """
        con_codes = self.extract_consumable_code()
        con_codes_new = []
        for elem in con_codes:
            if elem in con_codes_new:
                pass
            else:
                con_codes_new.append(elem)
        con_codes = con_codes_new
        con_codes = sorted(con_codes)
        return con_codes

    def extract_data_from_excel(self) -> dict:
        """Function opens given excel and reads it. Then we use the connbrs to find the row in the excel that has the desired information.
        Afterwards we save the information (Name and Material Number, Group, Specification, Vendor Code and Address, Used in Page Block)
        in a accordingly created dictionary.

        Returns:
            dict: holding the mentioned information.
        """
        con_codes = self.prepare_consumables()
        d_frame = pd.read_excel(self.excel_path)
        table = {
            'Name and Material Number': [],
            'Group': [],
            'Specification': [],
            'Vendor Code and Address': [],
            'Used in Page Block': []
        }
        for elem in con_codes:
            for ind, row in d_frame.iterrows():
                group_list = []
                spec_list = []
                sup_code_and_address_list = []
                pageblock_list = []
                if elem == row["Material No."]:
                    mat_name = str(row["Material Name"])
                    material_no = str(row['Material No.'])
                    group_list.append(str(row["Group"]).split("\n"))
                    spec_list.append(str(row["Specification"]).split("\n"))
                    for ind, elem in enumerate(spec_list[0]):
                        if '/' in elem:
                            spec_list[ind] = elem.split(' / ')
                            spec_list[ind][0] += ' /'
                    sup_code_and_address_list.append(str(row["Address"]).split("\n"))
                    sup_code_and_address_list[0].insert(0, row["Suppliers Code"])

                    pageblock_list.append(str(row["Used in Page Block"]).split("\n"))

                    table['Name and Material Number'].append([f"{mat_name} ({material_no})"])
                    table['Group'].append(group_list)
                    table['Specification'].append(spec_list)
                    table['Vendor Code and Address'].append(sup_code_and_address_list)
                    pageblock_list[0][0] = sub(",", "", pageblock_list[0][0])
                    if ' ' in pageblock_list[0][0]:
                        pageblock_list = [pageblock_list[0][0].split(' ')]
                    table['Used in Page Block'].append(pageblock_list)
                    break
        return table

    def format_table(self):
        """Function removes unnecessary nesting of lists and performs
        the following corrections in the dictionary:
        1. Removes all nan values
        2. Removes every leading trailing whitespaces from the dictionary
        3. Removes empty entries
        Afterwards the dictionary is saved in the main_dict variable defined in init.
        """
        table = self.extract_data_from_excel()
        for ind, elem in enumerate(table['Group']):
            table['Group'][ind] = table['Group'][ind][0]
            table['Specification'][ind] = table['Specification'][ind][0]
            table['Vendor Code and Address'][ind] = table['Vendor Code and Address'][ind][0]
            table['Used in Page Block'][ind] = table['Used in Page Block'][ind][0]

        # Removes all nan values
        table = {key: [[x for x in sublist if str(x) != 'nan'] for sublist in value] for key, value in table.items()}

        # Removes every leading trailing whitespaces from the dictionary
        table = {key: [[value.strip() if isinstance(value, str) else value for value in values] for values in table[key]] for key in table}

        # Removes empty entries
        for key, lists in table.items():
            new_lists = [[elem for elem in lst if elem] for lst in lists]
            table[key] = new_lists

        self.main_dict = table

    def create_table_skeleton(self) -> str:
        """The function creates the skeleton of the table which doesn't hold any
        important information and serves solely for the next processing steps.
        In the right there is a representation of the skeleton where:
        content 0: Name and Material Number
        content 1: Group
        content 2: Specification
        content 3: Vendor Code and Address
        content 4: Used in Page Block

        # Skeleton:
        # <tboody>
        #   <row>
        #       <entry rotate="0" valign="middle">content0</entry>
        #       <entry rotate="0" valign="middle">content1</entry>
        #       <entry rotate="0" valign="middle">content2</entry>
        #       <entry rotate="0" valign="middle">content3</entry>
        #       <entry rotate="0" valign="middle">content4</entry>
        #   </row>
        # </tbody>

        Raises:

            DictError: We raise an error in case if main_dict is still None.

        Returns:
            str: which has the content of the table skeleton.
        """
        self.format_table()
        if self.main_dict is None:
            raise DictError("An error occured when creating the dictionary for the consumables")

        tbody_xml = "<tbody>content\n</tbody>"
        rows = ""
        entrys = ""
        for _, _ in enumerate(self.main_dict['Group']):  # _, _ was ind, elem
            rows += "\n<row>\ncontent</row>"
        for i in range(5):
            entrys += f'<entry rotate="0" valign="middle">content{i}</entry>'

        tbody_xml = sub("content", rows, tbody_xml)
        tbody_xml = sub("content", entrys, tbody_xml)
        return tbody_xml

    def fill_skeleton(self) -> str:
        """Summary

        Returns:
            str: _description_
        """
        tbody_xml = self.create_table_skeleton()

        for ind, _ in enumerate(self.main_dict['Group']):  # _ was elem
            amount_para = len(self.main_dict['Name and Material Number'][ind])
            para = ""
            for i in range(amount_para):
                name = search(r"(.*?)( \()(.*?)(\))", self.main_dict['Name and Material Number'][ind][i]).group(1)
                mat_no = search(r"(.*?)( \()(.*?)(\))", self.main_dict['Name and Material Number'][ind][i]).group(3)
                para += f"<para><con><connbr>{mat_no}</connbr><conname>{name}</conname></con></para>"
            if "content0" in tbody_xml:
                tbody_xml = tbody_xml.replace("content0", para, 1)

            amount_para = len(self.main_dict['Group'][ind])
            para = ""
            for i in range(amount_para):
                para += f"<para>{self.main_dict['Group'][ind][i]}</para>"
            if para == "":
                tbody_xml = tbody_xml.replace("content1", "", 1)
            elif "content1" in tbody_xml:
                tbody_xml = tbody_xml.replace("content1", para, 1)

            amount_para = len(self.main_dict['Specification'][ind])
            para = ""
            for i in range(amount_para):
                para += f"<para>{self.main_dict['Specification'][ind][i]}</para>"
            if para == "":
                tbody_xml = tbody_xml.replace("content2", "", 1)
            elif "content2" in tbody_xml:
                tbody_xml = tbody_xml.replace("content2", para, 1)

            amount_para = len(self.main_dict['Vendor Code and Address'][ind])
            para_address = ""
            for i in range(amount_para):
                para_address += f"<para>{self.main_dict['Vendor Code and Address'][ind][i]}</para>"

            # new_address = [para_address] # Cosmin - comented, this is not used

            amount_para = len(self.main_dict['Used in Page Block'][ind])
            para_pgblk = ""
            for i in range(amount_para):
                para_pgblk += f"<para>{self.main_dict['Used in Page Block'][ind][i]}</para>"

            if para_address == "":
                tbody_xml = tbody_xml.replace("content3", "", 1)
            elif "content3" in tbody_xml:
                tbody_xml = tbody_xml.replace("content3", para_address, 1)

            if para_pgblk == "":
                tbody_xml = tbody_xml.replace("content4", "", 1)
            elif "content4" in tbody_xml:
                tbody_xml = tbody_xml.replace("content4", para_pgblk, 1)

        return tbody_xml

    def prepare_tbody(self) -> etree._Element:
        """Function that does modifications to the filled skeleton of the table
        to avoid unicode broblems. Afterwards the skeleton is transformed to a
        lxml etree object

        Returns:
            etree._Element: which represents the created table body
        """
        tbody_xml = self.fill_skeleton()

        tbody_xml = linearize_xml(tbody_xml)
        tbody_xml = tbody_xml.replace("&", "myand")
        tbody_xml = etree.fromstring(tbody_xml)
        return tbody_xml

    def center_pageblocks(self) -> any:
        """Centers text in the last entry of a row (pageblock) in tbody_xml.

        Returns:
            any: tbody_xml with centered text in the last entry of a row.
        """
        tbody_xml = self.prepare_tbody()
        try:
            tbody_xml = etree.tostring(tbody_xml, encoding="unicode")
        except TypeError:
            pass
        for row in findall(r"<row>.*?</row>", tbody_xml):
            last_entrys = findall(r"<entry.*?</entry>", row)[-1]
            last_paras = findall(r"<para>.*?</para>", last_entrys)
            try:
                if search(r"<para>\d+</para>", last_paras[-1]):
                    updated_last_entrys = sub(r"(<entry.*?)valign=('|\")middle('|\")", r'\1valign="middle" align="center"', last_entrys)
                    tbody_xml = tbody_xml.replace(last_entrys, updated_last_entrys)
            except IndexError:
                pass
        return tbody_xml

    def find_same_successors(self) -> Tuple[str, dict]:
        """Creates a dictionary which holds the information of
        how many identical successors each row has

        Returns:
            Tuple[str, dict]: string contains the tbody, dict contains the successor information
        """
        tbody_xml = self.center_pageblocks()

        succ_dict = {}
        tbody_xml = etree.fromstring(tbody_xml)
        rows = tbody_xml.findall('.//row')
        for ind, row in enumerate(rows):
            count = 0
            fourth_entry = row.xpath('.//entry[4]')[0]
            fourth_entry = ' '.join(fourth_entry.itertext())
            fifth_entry = row.xpath('.//entry[5]')[0]
            fifth_entry = ' '.join(fifth_entry.itertext())

            if fourth_entry.lower() != "local purchase":
                for next_row in rows[ind + 1:]:
                    fourth_entry_next = next_row.xpath('.//entry[4]')[0]
                    fourth_entry_next = ' '.join(fourth_entry_next.itertext())
                    fifth_entry_next = next_row.xpath('.//entry[5]')[0]
                    fifth_entry_next = ' '.join(fifth_entry_next.itertext())
                    if fourth_entry == fourth_entry_next:
                        count += 1
                    else:
                        break
                succ_dict[ind] = count
        return tbody_xml, succ_dict

    def add_morerows(self) -> str:
        """Function iterates through each row and uses the dictionary with
        the successor information to add the "morerows" attribute to the
        right entrys.
        Entrys following the "morerows" attribute will be deleted if they
        meet specific conditions.

        Returns:
            str: String containing the tbody with the changes for the "morerows" attribute
        """
        tbody_xml, succ_dict = self.find_same_successors()

        rows = tbody_xml.findall('.//row')
        last = 0
        for ind, row in enumerate(rows):
            for i, entry in enumerate(row.findall("entry")):
                if i in [3, 4]:
                    try:
                        last = succ_dict[ind - 1]
                    except KeyError:
                        last = succ_dict[0]
                    try:
                        if succ_dict[ind] > 0 and last == 0:
                            # Add morerows
                            entry.set("morerows", str(succ_dict[ind]))
                        elif succ_dict[ind] >= 0 and last != 0:
                            # Delete
                            row.remove(entry)
                        else:
                            # Do nothing
                            pass
                    except KeyError:
                        pass

        return tbody_xml

    def create_table_file(self) -> str:
        """Return a string representation of the created table,
        with "deleteAddress" and "deletePgblk" removed,
        and special characters escaped.

        Returns:
            str: The string representation of the table.
        """
        tbody_xml = self.add_morerows()
        try:
            tbody_xml = etree.tostring(tbody_xml, encoding="unicode")
        except TypeError:
            pass
        with open(join(self.export_path, f"cons_tbody_{basename(normpath(self.xml_path))}.xml"), "w", encoding="utf-8") as _:
            _.write(
                tbody_xml
                .replace('<entry rotate="0" valign="middle">deleteAddress</entry>', '')
                .replace('<entry rotate="0" valign="middle">deletePgblk</entry>', '')
                .replace("&", "&amp;")
                .replace("myand", "&amp;"))
        return tbody_xml

    def check_xml_path(self):
        """Check if the xml path is set, and raise NoXmlSet error if not set

        Raises:
            NoXmlSet: If xml path is not set
        """
        if self.xml_path is None:
            raise NoXmlSet("No xml found to validate. Use the set_xml function to set the path of your xml \
                        before you use the validate_consumables function.".replace("                        ", ""))

    def prepare_xml(self) -> str:
        """Function prepares the XML content read from a file specified by the xml_path attribute of the object.
        1. Linearizes the content (All content in first line)
        2. Deltes the first line (i.e "<?xml version="1.0" encoding="UTF-8"?>")
        3. Replaces special characters that are non unicode

        Returns:
            str: Representation of the xml content after perfoming the mentioned modifications.
        """
        with open(self.xml_path, "r", encoding="utf-8") as _:
            xml_content = _.read()
        xml_content = linearize_xml(xml_content)
        xml_content = delete_first_line(xml_content)
        xml_content = replace_special_characters(xml_content)
        return xml_content

    def insert_ISOEntities(self) -> str:
        """This function is used to insert replacements for all ISO entities.
        The function opens the inmedISOEntities.ent file and replaces the entities in the xml content.

        Returns:
            str: of the xml content with the entity replacements.
        """
        xml_content = self.prepare_xml()
        # Insert replacements for all entitys
        with open(join(FILEPATH, "inmedISOEntities.ent"), "r", encoding="utf-8") as _:
            entities = _.read()
        entities = linearize_xml(entities)
        xml_content = sub(r"(]>.*?<cmm)", entities + r'\1', xml_content)
        return xml_content

    def extract_pgblk_9000(self) -> etree._Element:
        """Function extracts the pgblk with pgblknbr '9000' from the xml file.

        Returns:
            etree._Element: Representation for the pageblock 9000
        """
        xml_content = self.insert_ISOEntities()

        xml_content = etree.fromstring(xml_content)
        pgblk_9000 = xml_content.xpath("//pgblk[@pgblknbr='9000']")[0]
        return pgblk_9000

    def extract_tbody_from_xml(self) -> str:
        """Extracts the tbody element from the given xml file.
        If the tbody element does not exist, then the function will exit instead.

        Returns:
            str: The extracted tbody as a string.
        """
        pgblk_9000 = self.extract_pgblk_9000()

        pgblk_9000 = etree.tostring(pgblk_9000, encoding="unicode")
        tree = etree.fromstring(pgblk_9000)
        # Check if table exists in xml
        task = tree.find('.//task[title="Consumables"]')
        if task is None:
            done()  # Cosmin: replaced frome exit to done to avoid name conflict
        task = etree.tostring(task, encoding="unicode")
        tree = etree.fromstring(task)
        tbody_from_xml = tree.find('.//tbody')
        tbody_from_xml = etree.tostring(tbody_from_xml, encoding="unicode")
        return tbody_from_xml

    def extract_text_from_original_table(self) -> list:
        """Extract the text (connbr, conname, Vendor Code and Address) in
        a structured way from the original tbody for further processing.
        Example: 'M011§Loctite 222|D2617 Henkel AG &amp; Co. KGAA P.O.Box 810580 81905 Muenchen Germany www.henkel.de'

        Returns:
            list: containing the strcutured representation of the text of each row from the original_table.
        """
        tbody_from_xml = self.extract_tbody_from_xml()

        if search(r"(\<row\>.*?\</row\>)", tbody_from_xml):
            original_rows = findall(r"(\<row\>.*?\</row\>)", tbody_from_xml)

            for ind, row in enumerate(original_rows):
                entry_tags = findall(r'<entry.*?>(.*?)</entry>', row)
                entry_text = []

                for entry in entry_tags:
                    para_tags = findall(r'<para.*?>(.*?)</para>', entry)
                    para_text = ' '.join(para_tags)
                    para_text = para_text.replace("&", "&")
                    entry_text.append(para_text)

                result = '~'.join(entry_text)
                result = result.replace("Alternative:", "")
                result = result\
                    .replace("<con>", "")\
                    .replace("</con>", "")\
                    .replace("<connbr>", "")\
                    .replace("</connbr>", "§")\
                    .replace("<conname>", "")\
                    .replace("</conname>", "|")
                result = sub(r'<.*?>.*?<.*?>~', '', result)
                result = result.split("|")
                elem_1 = result[0]
                try:
                    elem_2 = result[1].split("~")[-2]
                except IndexError:
                    pass
                original_rows[ind] = elem_1 + "|" + elem_2
        return original_rows

    def extract_text_from_created_table(self, created_tbody: str) -> list:
        """Extract the text (connbr, conname, Vendor Code and Address) in
        a structured way from the original tbody for further processing.
        Example: 'M011§Loctite 222|D2617 Henkel AG &amp; Co. KGAA P.O.Box 810580 81905 Muenchen Germany www.henkel.de'

        Args:
            created_tbody (str): 

        Returns:
            list: containing the strcutured representation of the text of each row from the created_table.
        """
        created_tbody = created_tbody.replace("\n", "")
        if search(r"(\<row\>.*?\</row\>)", created_tbody):
            created_rows = findall(r"(\<row\>.*?\</row\>)", created_tbody)

            for ind, row in enumerate(created_rows):
                entry_tags = findall(r'<entry.*?>(.*?)</entry>', row)
                entry_text = []

                for entry in entry_tags:
                    para_tags = findall(r'<para.*?>(.*?)</para>', entry)
                    para_text = ' '.join(para_tags)
                    para_text = para_text.replace("&", "&")
                    entry_text.append(para_text)

                result = '~'.join(entry_text)
                result = result.replace("Alternative:", "")
                result = result\
                    .replace("<con>", "")\
                    .replace("</con>", "")\
                    .replace("<connbr>", "")\
                    .replace("</connbr>", "§")\
                    .replace("<conname>", "")\
                    .replace("</conname>", "|")
                result = sub(r'<.*?>.*?<.*?>~', '', result)
                result = result.split("|")
                elem_1 = result[0]
                elem_2 = result[1].split("~")[-2]
                created_rows[ind] = elem_1 + "|" + elem_2

        for ind, _ in enumerate(created_rows):  # _ was elem
            created_rows[ind] = created_rows[ind].replace("myand", "&")
        return created_rows

    def validation_step(self, created_tbody: str):
        """Uses the lists original_rows and created_rows for the validation process.
        Function checks for each element of created_rows if it matches one of the element from original_rows.
        If for one element no match was found (i.e validation failed) the connbr, conname and address are extracted
        and written to a xml file which shows all failed validations.

        Args:
            created_tbody (str):
        """
        original_rows = self.extract_text_from_original_table()
        created_rows = self.extract_text_from_created_table(created_tbody)
        # Check if row from created table is in the original table
        with open(join(self.export_path, f"validated_table_{basename(normpath(self.xml_path))}.xml"), "w", encoding="utf-8") as _:
            _.write(f"Validated xml: {basename(self.xml_path)}\n\n")
            for ind, elem in enumerate(created_rows):
                found = False
                location = ""
                for elem2 in original_rows:
                    elem_no_ws = elem.replace(" ", "").lower()
                    elem2_no_ws = elem2.replace(" ", "").lower().replace("&amp;", "&")
                    if elem_no_ws == elem2_no_ws:
                        found = True
                if found is False:
                    # Extract Number
                    number_match = search(r"(M\d{3})", elem)
                    if number_match:
                        connbr = number_match.group(1)
                    else:
                        connbr = None
                    # extract the name
                    name_match = search(r"(M\d{3}§)([^|]+)(|)", elem)
                    if name_match:
                        conname = name_match.group(2)
                    else:
                        conname = None
                    # extract the Address
                    address_match = search(r"([^§]+§)([^|]+)(.*)", elem)
                    if address_match:
                        address = address_match.group(3).replace("|", "")
                    else:
                        address = None

                    # Find location of error
                    for element in original_rows:
                        if element is not None and connbr is not None:
                            if connbr in element:
                                element = element.replace("Alternative:§", "")
                                conname_in_element = element.split('§')[1].split('|')[0]
                                address_in_element = element.split('§')[1].split('|')[1]
                                if conname != conname_in_element and address != address_in_element:
                                    location = "conname & address"
                                    break
                                elif conname != conname_in_element:
                                    location = "conname"
                                    break
                                elif address != address_in_element:
                                    location = "address"
                                    break
                    _.write("The row with the following informations failed the validation:\n")
                    _.write(f"  connbr: {connbr}\n  conname: {conname}\n")
                    _.write(f"  address:\n{address}\n-> Location of Error: {location}\n\n")

    def validate_cons_table(
            self,
            debug: bool = False,
            qt_window: QMainWindow = None,
            progress: Signal = Signal(0),
            console: Signal = Signal("")) -> int:
        """Opens the given xml file and saves the content.
        If the xml file contains a consumables list the function validation_step
        is called to perform the validation of the table.
        """
        try:
            self.check_xml_path()
            tbody_xml = self.create_table_file()
            tbody_xml = tbody_xml.replace("&", "&amp;")
            with open(self.xml_path, "r", encoding="utf-8") as _:
                xml_content = _.read()
            if search(r"<title>List of Consumables</title>", xml_content):
                self.validation_step(tbody_xml)

            if qt_window is not None:
                console.emit(
                    "Consumables table created successfully. See: " + join(
                        self.export_path,
                        f"cons_tbody_{basename(normpath(self.xml_path))}.xml") + "\n")
        except Exception as err:
            if qt_window is not None and debug:
                progress.emit(100)
                console.emit("Error: " + str(err) + "\n" + format_exc() + "\n")
                return 1
        return 0

if __name__ == "__main__":
    instance = ConsumablesList()
    # instance.set_xml(r"C:\Users\bakalarz\Desktop\01_XML_Samples\ATA\CRM-D9893-BD50-32-21-03RM_004-01_EN_TestM.xml")
    instance.set_xml(r"D:\CMM Automation\REWORK\CMM\2.1 rework\CMM-D9893-C091-32-11-21_000-01_EN.xml")
    instance.set_excel(r"D:\CMM Automation\REWORK\CMM\2.1 rework\MASTER_LOM_Issue_45_DRAFT.xlsx")
    instance.validate_cons_table()
