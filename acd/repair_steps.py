from re import findall
from re import sub
from re import search

from os.path import join
from os.path import expanduser
from os.path import basename
from os.path import dirname
from os.path import normpath

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill

from math import floor
from math import ceil

from ast import literal_eval

from lxml import etree

from traceback import format_exc

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal

from .xml_processing import delete_first_line
from .xml_processing import replace_special_characters
from .xml_processing import linearize_xml

FILEPATH = dirname(__file__)


def clean_xml_tags(xml_string: str) -> str:
    if "xmlns:xsi=" in xml_string:
        rv_attribute1 = search(r'xmlns:xsi=".*?"', xml_string).group(0)
        xml_string = xml_string.replace(rv_attribute1, "")
    if "xmlns:rdf=" in xml_string:
        rv_attribute2 = search(r'xmlns:rdf=".*?"', xml_string).group(0)
        xml_string = xml_string.replace(rv_attribute2, "")
    if "xmlns:dc=" in xml_string:
        rv_attribute3 = search(r'xmlns:dc=".*?"', xml_string).group(0)
        xml_string = xml_string.replace(rv_attribute3, "")

    xml_string = sub(r"\s+", " ", xml_string)
    return xml_string


class RepairSteps():
    def __init__(self) -> None:
        self.xml_path = None
        self.export_path = expanduser("~/Desktop")
        self.intermediate_row = "1"

    def set_xml(self, xml_path: str):
        """Function with which the user can set the xml to be checked.

        Args:
            xml_path (str): xml file path.
        """
        self.xml_path = xml_path

    def set_export_path(self, export_path: str):
        """Function to specify a path different to the default path (desktop),
        where to export the files the script produces.
        Args:
            export_path (str): path to where the generated files should be exported.
        """
        self.export_path = export_path

    def calculate_conv(self, value: float) -> str:
        if value == 0:
            return "0.0000"
        conversion = value * 0.03937
        conversion = ceil(conversion * 10000) / 10000

        conversion = str(conversion)
        while len(conversion.split('.')[1]) < 4:
            conversion += '0'
        return conversion

    def prepare_xml(self) -> str:
        """The prepare_xml() function reads an XML file and performs
        a series of preprocessing steps on it.
        Returns:
            str: It returns the resulting XML content as a string.
        """
        with open(self.xml_path, "r", encoding="utf-8") as _:
            xml_content = _.read()
            xml_content = linearize_xml(xml_content)
            xml_content = delete_first_line(xml_content)
            xml_content = replace_special_characters(xml_content)
            # Insert replacements for all entitys
            with open(join(FILEPATH, "inmedISOEntities.ent"), "r", encoding="utf-8") as _:
                entities = _.read()
            entities = linearize_xml(entities)
            xml_content = sub(r"(]>.*?<cmm)", entities + r'\1', xml_content)
        return xml_content

    def create_table(
            self,
            step: float = 0.2,
            debug: bool = False,
            qt_window: QMainWindow = None,
            progress: Signal = Signal(0),
            console: Signal = Signal("")):
        try:
            xml_content = self.prepare_xml()

            xml_content = etree.fromstring(xml_content)
            pgblk_6000 = xml_content.xpath(".//pgblk[@pgblknbr=6000]")[0]
            tables = pgblk_6000.xpath(".//table")
            tables_backup = tables
            max_progress = len(tables)
            for ind, table in enumerate(tables):
                text = ''.join(table.itertext())
                if "repair" in text.lower():
                    if "part number" in text.lower():
                        # Case with part number
                        this_table = etree.tostring(table, encoding="unicode")
                        # Get last row and limit
                        note_row = ""
                        if "NOTE:" not in findall(r"\<row\>.*?\</row\>", this_table)[-1]:
                            last_row_1 = findall(
                                r"\<row\>.*?\</row\>", this_table)[-1]
                        else:
                            note_row = findall(
                                r"\<row\>.*?\</row\>", this_table)[-1]
                            last_row_1 = findall(
                                r"\<row\>.*?\</row\>", this_table)[-2]
                        limit = float(findall(
                            r"\<para.*?\>.*?\</para\>", last_row_1)[-2].split(' ')[0].replace("<para>", ""))
                        this_table = this_table.replace(
                            last_row_1, "").replace(note_row, "")

                        tbody = search(r"\<tbody\>.*?\</tbody\>",
                                       this_table).group(0)
                        last_row_2 = findall(r"\<row\>.*?\</row\>", tbody)[-1]
                        paras = findall(r"\<para.*?\>.*?\</para\>", last_row_2)
                        start_min = float(
                            paras[-2].replace("<para>", "").replace("</para>", "").split(' ')[0])
                        start_max = float(
                            paras[-1].replace("<para>", "").replace("</para>", "").split(' ')[0])
                        added_rows = ""
                        step_name = ""
                        i = 1
                        while start_min < limit and start_max < limit:
                            start_min = round(start_min + step, 3)
                            start_min_conversion = self.calculate_conv(
                                float(start_min))
                            start_min_write = start_min
                            if len(str(start_min_write).split('.')[1]) < 3:
                                start_min_write = str(start_min_write)
                                while len(start_min_write.split('.')[1]) < 3:
                                    start_min_write = start_min_write + '0'
                            start_max = round(start_max + step, 3)
                            start_max_conversion = self.calculate_conv(
                                float(start_max))
                            start_max_write = start_max
                            if len(str(start_max_write).split('.')[1]) < 3:
                                start_max_write = str(start_max_write)
                                while len(start_max_write.split('.')[1]) < 3:
                                    start_max_write = start_max_write + '0'
                            i += 1
                            step_num = str(i)
                            if len(step_num) < 2:
                                step_num = '0' + step_num
                            added_rows += f'<row><?validrow {i + 4}?><entry colname="col1" valign="middle"><para>RS{step_num}</para>\
                                        </entry><entry colname="col1" valign="middle"><para>TBD</para></entry><entry \
                                        colname="col2" valign="middle"><para>{start_min_write} ({start_min_conversion})</para></entry><entry colname="col3" \
                                        valign="middle"><para>{start_max_write} ({start_max_conversion})</para></entry></row>'.replace("                                    ", "")
                            i = int(i)
                        this_table = sub(
                            "</tbody>", added_rows + note_row + "</tbody>", this_table)
                        tables[ind] = this_table
                    else:
                        try:
                            # Standard case
                            this_table = etree.tostring(
                                table, encoding="unicode")
                            # Get last row and limit
                            note_row = ""
                            if "NOTE:" not in findall(r"\<row\>.*?\</row\>", this_table)[-1]:
                                last_row_1 = findall(
                                    r"\<row\>.*?\</row\>", this_table)[-1]
                            else:
                                note_row = findall(
                                    r"\<row\>.*?\</row\>", this_table)[-1]
                                last_row_1 = findall(
                                    r"\<row\>.*?\</row\>", this_table)[-2]
                            limit = float(findall(
                                r"\<para\>.*?\</para\>", last_row_1)[1].split(' ')[0].replace("<para>", ""))
                            this_table = this_table.replace(
                                last_row_1, "").replace(note_row, "")

                            tbody = search(
                                r"\<tbody\>.*?\</tbody\>", this_table).group(0)
                            last_row_2 = findall(
                                r"\<row\>.*?\</row\>", tbody)[-1]
                            paras = findall(
                                r"\<para\>.*?\</para\>", last_row_2)
                            start_min = float(
                                paras[-2].replace("<para>", "").replace("</para>", "").split(' ')[0])
                            start_max = float(
                                paras[-1].replace("<para>", "").replace("</para>", "").split(' ')[0])
                            added_rows = ""
                            step_name = ""
                            first_iter = True
                            i = 1
                            while start_min < limit and start_max < limit:
                                start_min = round(start_min + step, 3)
                                start_min_conversion = self.calculate_conv(
                                    float(start_min))
                                start_min_write = start_min
                                if len(str(start_min_write).split('.')[1]) < 3:
                                    start_min_write = str(start_min_write)
                                    while len(start_min_write.split('.')[1]) < 3:
                                        start_min_write = start_min_write + '0'
                                start_max = round(start_max + step, 3)
                                start_max_conversion = self.calculate_conv(
                                    float(start_min))
                                start_max_write = start_max
                                if len(str(start_max_write).split('.')[1]) < 3:
                                    start_max_write = str(start_max_write)
                                    while len(start_max_write.split('.')[1]) < 3:
                                        start_max_write = start_max_write + '0'
                                i += 1
                                step_num = str(i)
                                if len(step_num) < 2:
                                    step_num = '0' + step_num
                                if first_iter:
                                    step_name = "FS"
                                else:
                                    step_name = "RS"
                                added_rows += f'<row><?validrow {i + 4}?><entry colname="col1" valign="middle"><para>{step_name}{step_num}</para></entry><entry \
                                            colname="col2" valign="middle"><para>{start_min_write} ({start_min_conversion})</para></entry><entry colname="col3" \
                                            valign="middle"><para>{start_max_write} ({start_max_conversion})</para></entry></row>'.replace("                                        ", "")
                                if first_iter:
                                    first_iter = False
                                    i -= 2
                            this_table = sub(
                                "</tbody>", added_rows + "</tbody>", this_table)
                            tables[ind] = this_table
                        except IndexError:
                            print(etree.tostring(table, encoding="unicode"))
                if qt_window is not None:
                    progress.emit(floor(ind / max_progress * 100))
            xml_content_new = xml_content
            pgblk_node = xml_content_new.xpath(".//pgblk[@pgblknbr=6000]")[0]
            table_nodes = pgblk_node.xpath(".//table")
            for ind, modified_table in enumerate(tables):
                table_to_replace = table_nodes[ind]
                if isinstance(modified_table, str):
                    modified_table = etree.fromstring(modified_table)
                table_to_replace.getparent().replace(table_to_replace, modified_table)
                # table_to_replace.getparent().replace(modified_table, table_to_replace)
            xml_content_new = etree.tostring(
                xml_content_new, encoding="unicode")

            xml_content_new = self.check_more_rows(xml_content_new)
            xml_content_new = self.replace_tbd_tag(xml_content_new)

            with open(join(self.export_path, f"completed_repair_step_tables_{basename(self.xml_path)}.xml"), 'w', encoding="utf-8") as _:
                _.write(xml_content_new)
            if qt_window is not None:
                console.emit(
                    "Repair steps table created successfully. See: " + join(
                        self.export_path,
                        f"completed_repair_step_tables_{basename(self.xml_path)}.xml") + "\n")
        except Exception as err:
            if qt_window is not None and debug:
                progress.emit(100)
                console.emit("Error: " + str(err) + "\n" + format_exc() + "\n")
                return 1
        return 0

    def check_more_rows(self, xml_content) -> str:
        xml_content = etree.fromstring(xml_content)
        pgblk_6000 = xml_content.xpath(".//pgblk[@pgblknbr=6000]")[0]
        tables = pgblk_6000.xpath(".//table")

        for ind, table in enumerate(tables):
            text = ''.join(table.itertext())
            if "repair" in text.lower():
                tbody = table.find(".//tbody")
                if tbody is not None:
                    rows = tbody.findall("row")
                    if len(rows) >= 3:
                        target_row = rows[2]
                        entries = target_row.findall("entry")
                        if len(entries) >= 2:
                            second_entry = entries[1]
                            number_morerows = second_entry.get("morerows")
                            if number_morerows is not None:
                                # ToDo: Move back to the parent which has multiple <row> childs and staring from the forth <row> child, for each <row> delete the second <entry> child until the amount of childs deleted is equal to the value of number_morerows
                                parent_row = target_row.getparent()
                                all_rows = parent_row.findall("row")
                                start_index = all_rows.index(target_row) + 1
                                delete_count = 0
                                for i in range(start_index, start_index + int(number_morerows)):
                                    if i >= len(all_rows):
                                        break
                                    row_to_delete_from = all_rows[i]
                                    second_entry_to_delete = row_to_delete_from.findall("entry")[
                                        1]
                                    row_to_delete_from.remove(
                                        second_entry_to_delete)
                                    delete_count += 1
                            else:
                                continue
        xml_content = etree.tostring(xml_content, encoding="unicode")
        return xml_content

    def replace_tbd_tag(self, xml_content) -> str:
        xml_content = etree.fromstring(xml_content)
        pgblk_6000 = xml_content.xpath(".//pgblk[@pgblknbr=6000]")[0]
        tables = pgblk_6000.xpath(".//table")

        for ind, table in enumerate(tables):
            text = ''.join(table.itertext())
            if "repair" in text.lower() and "TBD" in text:
                tbody = table.find(".//tbody")
                if tbody is not None:
                    # Find part number
                    row = tbody.find("row")
                    second_entry = row.findall("entry")[1]
                    para = second_entry.find("para")
                    part_num = para.text.replace("FS", "RS")
                    # Find row with <para> with text "TBD"
                    rows = tbody.findall("row")
                    for row in rows:
                        paras = row.findall(".//para")
                        for para in paras:
                            if "TBD" in para.text:
                                para.text = part_num
                                break
        xml_content = etree.tostring(xml_content, encoding="unicode")
        return xml_content


if __name__ == "__main__":
    instance = RepairSteps()
    instance.set_xml(
        r"C:\Users\bakalarz\Desktop\01_XML_Samples\ATA\CRM-D9893-AR21-32-10-02RM_EN - Kopie.xml")
    # r"C:\Users\bakalarz\Desktop\01_XML_Samples\ATA\CRM-D9893-BD50-32-21-03RM_004-01_EN - Kopie.xml")
    instance.set_export_path(r"C:\Users\bakalarz\Desktop\Export")
    instance.create_table()
