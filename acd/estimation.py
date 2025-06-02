from os import remove
from os.path import join
from os.path import dirname

from re import sub
from re import search
from re import findall

import xlsxwriter

from fitz import open as pdf_open
from pikepdf import open as pike_open
from pdfreader import SimplePDFViewer

from json import dump

from pandas import DataFrame

from .filelist import list_files
from .pdf import get_pdf_content
from .txt import word_frequency

def is_fullpage_illu(text: str) -> str:
    """Checks if the page is a fullpage illustration.

    Args:
        text (str): The text to check.

    Returns:
        str: "Illustration" or "Text + Illustration"
    """
    for row in text.split("\n"):
        step1 = search(r"\(\d{1,2}\)( |\n)", text)
        step2 = search(r"\([a-z]{1,2}\)( |\n)", text)
        step3 = search(r"^\d{1,2}\.( |\n)", text)
        step4 = search(r"[A-Z]{1,2}\.( |\n)", text)
        step5 = search(r"^\d{1,2}( |\n)", text)

        if step1 or step2 or step3 or step4 or step5:
            return "Text + Illustration"

    return "Illustration"

PAGEBLOCKS = {
    1: "TESTING AND FAULT ISOLATION",
    2: "SCHEMATIC AND WIRING DIAGRAMS",
    3: "DISASSEMBLY",
    4: "CLEANING",
    5: "INSPECTION/CHECK",
    6: "REPAIR",
    7: "ASSEMBLY",
    8: "FITS AND CLEARANCES",
    9: "SPECIAL TOOLS; FIXTURES; EQUIPMENT AND CONSUMABLES",
    10: "ILLUSTRATED PARTS LIST (IPL)",
    11: "SPECIAL PROCEDURES",
    15: "STORAGE (INCLUDING TRANSPORTATION)",
}

SVG_ELEMENT_REGEX = r'<svg.*?>'
SVG_HEIGHT_REGEX = r'(height=")(.*?)(")'
SVG_WIDTH_REGEX = r'(width=")(.*?)(")'

def estimate_illustration(svg_dir: str) -> dict:
    svg_data = {}
    for svg in list_files(svg_dir, True, ["svg"]):
        with open(svg, 'r', encoding='utf-8') as svg_in:
            svg_content = svg_in.read()

        svg_element = search(SVG_ELEMENT_REGEX, svg_content).group(0)
        svg_height = search(SVG_HEIGHT_REGEX, svg_element).group(2)
        svg_width = search(SVG_WIDTH_REGEX, svg_element).group(2)

        svg_height = float(sub(r"[^\d\.]", "", svg_height))
        svg_width = float(sub(r"[^\d\.]", "", svg_width))

        svg_data[svg] = {
            "height": svg_height,
            "width": svg_width,
            "fullpage": int(svg_height) >= 20,
            "content_length": len(svg_content),
            "g_count": len(findall(r"</g>", svg_content)),
            "a_count": len(findall(r"</a>", svg_content)),
            "polyline_count": len(findall(r"<polyline", svg_content)),
            "polygon_count": len(findall(r"<polygon", svg_content)),
            "circle_count": len(findall(r"<circle", svg_content)),
            "text_count": len(findall(r"<text", svg_content)),
            "path_count": len(findall(r"<path", svg_content)),
            "rect_count": len(findall(r"<rect", svg_content)),
            "clipPath_count": len(findall(r"<clipPath", svg_content)),
            "marker_count": len(findall(r"<marker", svg_content)),
        }
    with open(join(svg_dir, "svg_data.json"), 'w', encoding='utf-8') as svg_out:
        dump(svg_data, svg_out, indent=4)
    return svg_data

def prepare_estimation(pdf: str, type: str = "Revision", ret: bool = False):
    json_dict = {}

    ml_df = DataFrame()

    if not ret:
        result_excel = join(pdf.replace(".pdf", ".xlsx"))
        workbook = xlsxwriter.Workbook(result_excel)
        worksheet1 = workbook.add_worksheet("CMM")
        worksheet2 = workbook.add_worksheet("Illustrations")

        header_format = workbook.add_format(
            {'bold': True, 'border': 1, 'align': 'center', 'font_color': 'white', 'bg_color': '#0070C0'})
        normal_body_format = workbook.add_format(
            {'bold': False, 'border': 1, 'font_color': 'black'})
        normal_body_centered_format = workbook.add_format(
            {'bold': False, 'border': 1, 'font_color': 'black', 'align': 'center'})
        yellow_body_format = workbook.add_format(
            {'bold': False, 'border': 1, 'font_color': 'black', 'bg_color': 'yellow', 'align': 'center'})

        worksheet1.write(0, 0, "Section", header_format)
        worksheet1.write(0, 1, "Page Number", header_format)
        worksheet1.write(0, 2, "Estimation", header_format)

        worksheet2.write(0, 0, "Illustration", header_format)
        worksheet2.write(0, 1, "Estimation", header_format)

        # set width of the columns
        worksheet1.set_column(0, 0, 60)
        worksheet1.set_column(1, 1, 20)
        worksheet1.set_column(2, 2, 20)

        worksheet2.set_column(0, 0, 60)
        worksheet2.set_column(1, 1, 20)

        sheet1_row = 1
        sheet2_row = 1
    figures = []
    icns_list = []
    for ind, page in enumerate(pdf_open(pdf)):
        page_text = page.get_text()
        content_length = str(len(page_text))
        words = findall(r"[\w']+", page_text)
        word_count = str(len(words) - 13)  # 13 for header and footer
        average_word_length = str(round(sum(len(word) for word in words) / len(words), 2))
        page_match = search(r"(_)([A-Z\s]*)(\n)(Page\s\d+)", page_text)

        icn_match = search(r"ICN\-[A-Z]", page_text)
        icn_matches = len(findall(r"ICN\-[A-Z]", page_text))
        local_figures = []
        icns = []
        
        word_freq = word_frequency(page_text, False, False)

        if page_match:
            section = page_match.group(2).strip().replace(",", "")
            page_number = page_match.group(4).strip().replace("Page ", "")

            fullpage = "Text"

            if icn_match:
                figure_match = search(r"Figure\s([a-zA-Z]+\-)?\d+", page_text)
                # print(f"Found {icn_matches} ICNs on page {page_number}.")

                if figure_match:
                    for figure in findall(r"Figure\s(?:[a-zA-Z]+\-)?\d+", page_text):
                        if figure not in local_figures:
                            local_figures.append(figure)
                # for icn in findall(r"[I][C][N]\-[0-9A-Z]{4}\-\d\d\-\d\d\-\d\d(?:[R][M])?\-[0-9A-Z]{5}\-\d{5}\-[A]\d\d\_\d\d\d\.[C][G][M]", page_text):
                for icn in findall(r"[I][C][N]\-[0-9A-Z]{4}\-\d\d\-\d\d\-\d\d(?:[R][M])?\-[0-9A-Z]{5}\-\d{5}", page_text):
                    print(f"Found ICN: {icn}")
                    # ICN-CO91-32-61-11-D9893-01227-A01_000.CGM
                    if icn not in icns:
                        icns.append(icn)
                    if icn not in icns_list:
                        icns_list.append(icn)
                        if not ret:
                            worksheet2.write(sheet2_row, 0, icn, normal_body_format)
                            worksheet2.write(sheet2_row, 1, "", yellow_body_format)
                            sheet2_row += 1
                fullpage = is_fullpage_illu(page_text)

            local_figures = "; ".join(local_figures)
            icns = "; ".join(icns)

            if section != "":
                json_dict[section + "_" + page_number] = {
                    "Section": section,
                    "Page Number": page_number,
                    "Fullpage Illu?": fullpage,
                }
                if not ret:
                    worksheet1.write(sheet1_row, 0, section, normal_body_format)
                    worksheet1.write(sheet1_row, 1, page_number, normal_body_centered_format)
                    worksheet1.write(sheet1_row, 2, "", yellow_body_format)
                    sheet1_row += 1

            elif len(page_number) < 4:
                text_estimation = "Simple"
                illu_estimation = "N/A"
                if fullpage == "Illustration":
                    if int(word_count) > 100:
                        illu_estimation = "Complex"
                    else:
                        illu_estimation = "Simple"
                elif fullpage == "Text + Illustration":
                    if int(word_count) / 2 > 75:
                        illu_estimation = "Complex"
                    else:
                        illu_estimation = "Simple"
                    if int(word_count) / 2 > 125:
                        text_estimation = "Complex"
                    else:
                        text_estimation = "Simple"
                else:
                    if int(word_count) > 150:
                        text_estimation = "Complex"
                    else:
                        text_estimation = "Simple"
                json_dict[section + "_" + page_number] = {
                    "Section": "DESCRIPTION AND OPERATION",
                    "Page Number": page_number,
                    "Fullpage Illu?": fullpage,
                }
                if not ret:
                    worksheet1.write(sheet1_row, 0, "DESCRIPTION AND OPERATION", normal_body_format)
                    worksheet1.write(sheet1_row, 1, page_number, normal_body_centered_format)
                    worksheet1.write(sheet1_row, 2, "", yellow_body_format)
                    sheet1_row += 1

            elif len(page_number) == 4 or page_number.startswith("15") or page_number.startswith("11"):
                text_estimation = "Simple"
                illu_estimation = "N/A"
                if fullpage == "Illustration" and len(icns) == 1:
                    if int(word_count) >= 100:
                        illu_estimation = "Complex"
                    elif int(word_count) < 100 and int(word_count) >= 50:
                        illu_estimation = "Middle"
                    else:
                        illu_estimation = "Simple"
                if fullpage == "Illustration" and len(icns) == 2:
                    if int(word_count) >= 120:
                        illu_estimation = "Complex"
                    elif int(word_count) < 120 and int(word_count) >= 60:
                        illu_estimation = "Middle"
                    else:
                        illu_estimation = "Simple"
                elif fullpage == "Text + Illustration":
                    if int(word_count) >= 180:
                        illu_estimation = "Complex"
                    elif int(word_count) < 180 and int(word_count) >= 90:
                        illu_estimation = "Middle"
                    else:
                        illu_estimation = "Simple"

                    if int(word_count) / 2 > 125:
                        text_estimation = "Complex"
                    else:
                        text_estimation = "Simple"
                else:
                    if int(word_count) > 150:
                        text_estimation = "Complex"
                    else:
                        text_estimation = "Simple"
                json_dict[section + "_" + page_number] = {
                    "Section": PAGEBLOCKS[int(page_number[:len(page_number) - 3])],
                    "Page Number": page_number,
                    "Fullpage Illu?": fullpage,
                }
                if not ret:
                    worksheet1.write(sheet1_row, 0, PAGEBLOCKS[int(page_number[:len(page_number) - 3])], normal_body_format)
                    worksheet1.write(sheet1_row, 1, page_number, normal_body_centered_format)
                    worksheet1.write(sheet1_row, 2, text_estimation, yellow_body_format)
                    sheet1_row += 1
        elif "TP1" in page_text:
            json_dict["TP" + "_" + "1"] = {
                "Section": "TP",
                "Page Number": "1",
                "Fullpage Illu?": '',
            }
            if not ret:
                worksheet1.write(sheet1_row, 0, "TP", normal_body_format)
                worksheet1.write(sheet1_row, 1, "1", normal_body_centered_format)
                worksheet1.write(sheet1_row, 2, "", yellow_body_format)
                sheet1_row += 1

        elif "TP2" in page_text:
            json_dict["TP" + "_" + "2"] = {
                "Section": "TP",
                "Page Number": "2",
                "Fullpage Illu?": '',
            }
            if not ret:
                worksheet1.write(sheet1_row, 0, "TP", normal_body_format)
                worksheet1.write(sheet1_row, 1, "2", normal_body_centered_format)
                worksheet1.write(sheet1_row, 2, "", yellow_body_format)
                sheet1_row += 1

        else:
            print(page_text)
            break
    if not ret:
        workbook.close()
    if ret:
        return json_dict

if __name__ == "__main__":
    from os import listdir
    for pdf in listdir(r"C:\Users\munteanu\Downloads\PDF page numbers to excel"):
        if pdf.endswith(".pdf"):
            prepare_estimation(
                join(r"C:\Users\munteanu\Downloads\PDF page numbers to excel", pdf),
                "Creation"
            )
