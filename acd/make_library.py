import sys

from os import listdir

from os.path import sep
from os.path import join
from os.path import isdir
from os.path import dirname
from os.path import getmtime

from datetime import datetime

from re import search
from re import findall

from tqdm import tqdm

from traceback import format_exc

from pandas import concat
from pandas import DataFrame
from pandas import ExcelWriter

import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal

from .filelist import list_files  # pylint: disable=import-error,E0401,C0413
from .excel_ import format_excel  # pylint: disable=import-error,E0401,C0413

def get_manual_series(
        manual_xml: str,
        manual_type: str,
        items: dict,
        illus: list,
        ata: str,
        year: int,
        repair_tasks: list,
        testing_subtasks: list,
        disassembly_subtasks: list,
        assembly_subtasks: list,
        special_tools: dict) -> DataFrame:
    repair_df_ = DataFrame(
        columns=['Year', 'ATA', 'Manual Type', 'Repair Tasks', 'XML Content', ''])
    testing_df_ = DataFrame(
        columns=['Year', 'ATA', 'Manual Type', 'Testing Subtask', 'XML Content', ''])
    disassembly_df_ = DataFrame(
        columns=['Year', 'ATA', 'Manual Type', 'Disassembly Subtask', 'XML Content', ''])
    assembly_df_ = DataFrame(
        columns=['Year', 'ATA', 'Manual Type', 'Assembly Subtask', 'XML Content', ''])
    with open(manual_xml, 'r', encoding='utf-8') as cmm_file:
        cmm_text = cmm_file.read().replace('\n', '')
    if search(r'<figure.*?</figure>', cmm_text) is not None:
        for figure in findall(r'<figure.*?</figure>', cmm_text):
            try:
                figure_no = search(r'(fignbr=")(.*?)(")',
                                   ''.join(figure)).group(2)
            except AttributeError:
                figure_no = "01"
            if search(r'<itemdata.*?</itemdata>', figure) is not None:
                for item in findall(r'<itemdata.*?</itemdata>', figure):
                    try:
                        item_no = search(r'(itemnbr=")(.*?)(")',
                                         ''.join(item)).group(2)
                    except AttributeError:
                        continue
                    pnr = search(r'(<pnr>)(.*?)(</pnr>)', item).group(2)
                    try:
                        kwd = search(r'(<kwd>)(.*?)(</kwd>)', item).group(2)
                    except AttributeError:
                        kwd = ""
                    adt = ""
                    if search(r'<adt>.*?</adt>', item) is not None:
                        adt = search(r'(<adt>)(.*?)(</adt>)',
                                     item).group(2) + " "
                    items[f'{figure_no}-{item_no}'] = [pnr, f'{adt}{kwd}']
    if search(r'<graphic.*?</graphic>', cmm_text) is not None:
        for graphic in findall(r'<graphic.*?</graphic>', cmm_text):
            # potentially check if an item number is in the illu title,
            # if yes, check the pn of the item and add it to the list
            sheets = []
            for sheet in findall(r'<sheet.*?</sheet>', graphic):
                matched = search(r'(gnbr=")(.*?)(")', ''.join(sheet))
                if matched:
                    sheets.append(matched.group(2))
            if search(r'(<title>)(.*?)(</title>)', ''.join(graphic)) is not None:
                illus.append(search(r'(<title>)(.*?)(</title>)',
                                    ''.join(graphic)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' ') + " - " + str(sheets))

    if search(r'(pgblknbr="6000")(.*?)(</pgblk>)', cmm_text) is not None:
        for task in findall(r'(<task)(.*?)(</task>)', search(r'(pgblknbr="6000")(.*?)(</pgblk>)', cmm_text).group(2)):
            repair_tasks.append(
                search(r'(<title>)(.*?)(</title>)', ''.join(task)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' '))

            repair_df_ = concat([repair_df_, DataFrame(
                {'Year': [year],
                 'ATA': [ata],
                 'Manual Type': [manual_type],
                 'Repair Tasks': [
                     search(r'(<title>)(.*?)(</title>)', ''.join(task)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' ')],
                 'XML Content': [''.join(task)],
                 '': [' ']
                 })], ignore_index=True)

    if search(r'(pgblknbr="1000")(.*?)(</pgblk>)', cmm_text) is not None:
        for subtask in findall(r'(<subtask)(.*?)(</subtask>)', search(r'(pgblknbr="1000")(.*?)(</pgblk>)', cmm_text).group(2)):
            testing_subtasks.append(
                search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' '))
            testing_df_ = concat([testing_df_, DataFrame(
                {'Year': [year],
                 'ATA': [ata],
                 'Manual Type': [manual_type],
                 'Testing Subtask': [
                     search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' ')],
                 'XML Content': [''.join(subtask)],
                 '': [' ']
                 })], ignore_index=True)

    if search(r'(pgblknbr="3000")(.*?)(</pgblk>)', cmm_text) is not None:
        for subtask in findall(r'(<subtask)(.*?)(</subtask>)', search(r'(pgblknbr="3000")(.*?)(</pgblk>)', cmm_text).group(2)):
            disassembly_subtasks.append(
                search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' '))
            disassembly_df_ = concat([disassembly_df_, DataFrame(
                {'Year': [year],
                 'ATA': [ata],
                 'Manual Type': [manual_type],
                 'Disassembly Subtask': [
                     search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' ')],
                 'XML Content': [''.join(subtask)],
                 '': [' ']
                 })], ignore_index=True)

    if search(r'(pgblknbr="7000")(.*?)(</pgblk>)', cmm_text) is not None:
        for subtask in findall(r'(<subtask)(.*?)(</subtask>)', search(r'(pgblknbr="7000")(.*?)(</pgblk>)', cmm_text).group(2)):
            assembly_subtasks.append(
                search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' '))
            assembly_df_ = concat([assembly_df_, DataFrame(
                {'Year': [year],
                 'ATA': [ata],
                 'Manual Type': [manual_type],
                 'Assembly Subtask': [
                     search(r'(<title>)(.*?)(</title>)', ''.join(subtask)).group(2).replace('<?Pub Dtl?>', '').replace('\xa0', ' ')],
                 'XML Content': [''.join(subtask)],
                 '': [' ']
                 })], ignore_index=True)

    if search(r'(pgblknbr="9000")(.*?)(</pgblk>)', cmm_text) is not None:
        special_tools_table = search(r"(<table.*?><title>Special Tools</title>)(.*?)(</table>)",
                                     search(r'(pgblknbr="9000")(.*?)(</pgblk>)', cmm_text).group(2))
        empty = []
        if special_tools_table is not None:
            for tool in findall(r'(<ted>)(.*?)(</ted>)', special_tools_table.group(2)):
                special_tools[search(r'(<toolnbr>)(.*?)(</toolnbr>)', ''.join(tool)).group(
                    2)] = search(r'(<toolname>)(.*?)(</toolname>)', ''.join(tool)).group(2)
    return DataFrame(
        {'Year': [year],
         'ATA': [ata],
         'Manual Type': [manual_type],
         'Illus': [illus],
         'PNs': [items],
         'Special Tools': [special_tools],
         '': [' ']
         # 'Repair Tasks': [repair_tasks],
         # 'Testing Subtasks': [testing_subtasks]
         }
    ), repair_df_, testing_df_, disassembly_df_, assembly_df_


def make_library(
        top_dir: str,
        debug: bool = False,
        qt_window: QMainWindow = None,
        progress: Signal = Signal(0),
        console: Signal = Signal("")) -> int:

    try:
        # report_df = DataFrame(columns=[
        #                       'Year', 'ATA', 'Manual Type', 'Illus', 'PNs', 'Repair Tasks', 'Special Tools', 'Testing Subtasks'])
        report_df = DataFrame(columns=[
            'Year', 'ATA', 'Manual Type', 'Illus', 'PNs', 'Special Tools', ''])

        repair_df = DataFrame(
            columns=['Year', 'ATA', 'Manual Type', 'Repair Tasks', 'XML Content', ''])
        testing_df = DataFrame(
            columns=['Year', 'ATA', 'Manual Type', 'Testing Subtask', 'XML Content', ''])
        disassembly_df = DataFrame(
            columns=['Year', 'ATA', 'Manual Type', 'Disassembly Subtask', 'XML Content', ''])
        assembly_df = DataFrame(
            columns=['Year', 'ATA', 'Manual Type', 'Assembly Subtask', 'XML Content', ''])

        max_progress = len(listdir(top_dir))
        for ind, year in enumerate(listdir(top_dir)):
            if isdir(join(top_dir, year)):
                for workpackage in listdir(join(top_dir, year)):
                    if isdir(join(top_dir, year, workpackage)):
                        try:
                            ata = search(r'\d\d\-\d\d\-\d\d', workpackage).group()
                        except AttributeError:
                            continue
                        for wp_folder in listdir(join(top_dir, year, workpackage)):
                            # illus = {}
                            illus = []
                            items = {}
                            repair_tasks = []
                            testing_subtasks = []
                            disassembly_subtasks = []
                            assembly_subtasks = []
                            special_tools = {}
                            cmm_xml = (None, None)
                            crm_xml = (None, None)
                            if "delivery" in wp_folder.lower():
                                for xml in list_files(join(top_dir, year, workpackage, wp_folder), ext_list=['xml']):
                                    xml_title = xml.split(sep)[-1]
                                    if 'meta' in xml_title.lower():
                                        continue
                                    if "cmm" in xml_title.lower() and cmm_xml[1] is None:
                                        cmm_xml = (
                                            xml, datetime.fromtimestamp(getmtime(xml)))
                                    elif "rm" in xml_title.lower() and crm_xml[1] is None:
                                        crm_xml = (
                                            xml, datetime.fromtimestamp(getmtime(xml)))
                                    elif "cmm" in xml_title.lower() and datetime.fromtimestamp(getmtime(xml)) > cmm_xml[1]:
                                        cmm_xml = (
                                            xml, datetime.fromtimestamp(getmtime(xml)))
                                        cmm_xml = (
                                            xml, datetime.fromtimestamp(getmtime(xml)))
                                    elif "rm" in xml_title.lower() and datetime.fromtimestamp(getmtime(xml)) > crm_xml[1]:
                                        crm_xml = (
                                            xml, datetime.fromtimestamp(getmtime(xml)))
                            else:
                                # get content from input files
                                pass
                            report = None
                            if cmm_xml[0] is not None:
                                report = get_manual_series(
                                    cmm_xml[0],
                                    'CMM',
                                    items,
                                    illus,
                                    ata,
                                    year,
                                    repair_tasks,
                                    testing_subtasks,
                                    disassembly_subtasks,
                                    assembly_subtasks,
                                    special_tools
                                )
                            if crm_xml[0] is not None:
                                report = get_manual_series(
                                    crm_xml[0],
                                    'CRM',
                                    items,
                                    illus,
                                    ata,
                                    year,
                                    repair_tasks,
                                    testing_subtasks,
                                    disassembly_subtasks,
                                    assembly_subtasks,
                                    special_tools
                                )
                            if report is not None:
                                report_df = concat(
                                    [report_df, report[0]], ignore_index=True)
                                repair_df = concat(
                                    [repair_df, report[1]], ignore_index=True)
                                testing_df = concat(
                                    [testing_df, report[2]], ignore_index=True)
                                disassembly_df = concat(
                                    [disassembly_df, report[3]], ignore_index=True)
                                assembly_df = concat(
                                    [assembly_df, report[4]], ignore_index=True)
                if qt_window is not None:
                    progress.emit(int(ind / max_progress * 100))
        # Create an ExcelWriter object
        excel_writer = ExcelWriter(join(top_dir, "Library.xlsx"), engine='xlsxwriter')  # pylint: disable=abstract-class-instantiated

        # Write each DataFrame to a separate sheet
        report_df.to_excel(
            excel_writer, sheet_name='Library Overview', index=False)
        repair_df.to_excel(excel_writer, sheet_name='Repair Tasks', index=False)
        testing_df.to_excel(
            excel_writer, sheet_name='Testing Subtasks', index=False)
        disassembly_df.to_excel(
            excel_writer, sheet_name='Disassembly Subtasks', index=False)
        assembly_df.to_excel(
            excel_writer, sheet_name='Assembly Subtasks', index=False)

        # Save the Excel file
        excel_writer.save()
        excel_writer.close()

        format_excel(
                join(top_dir, "Library.xlsx"),
                body_format={'bold': False, 'border': 1, 'bg_color': '#D6EAFF', 'font_color': 'black'},
                inplace=True,
                header_format={
                    'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#0070C0',
                    'font_color': 'white', 'font_size': 14},
                freeze_panes=(1, 0),
                column_widths={0: 8, 1: 12, 2: 18, 3: 60, 4: 60, 5: 60},
                column_alignments={0: 'center', 1: 'center', 2: 'center', 3: 'left', 4: 'left', 5: 'left', 6: 'left'}
        )

        if qt_window is not None:
            progress.emit(100)
            console.emit(
                "Library was created. See: " + join(top_dir, "Library.xlsx") + "\n")
    except Exception as err:
        if qt_window is not None and debug:
            progress.emit(100)
            console.emit("Error: " + str(err) + "\n" + format_exc() + "\n")
            return 1
    return 0
