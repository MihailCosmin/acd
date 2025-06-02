from .ocr_pdf import get_ocr_pdf_content
from .ocr_pdf import ocr_pdf
from .xml_validation import XmlSchemaValidator
from .xml_validation import Punctuation
from .stp import add_iplnom_to_stp
from .smg import add_iplnom_to_smg
from .estimation import is_fullpage_illu
from .estimation import estimate_illustration
from .estimation import prepare_estimation
from .pdf import pdf_page_count
from .pdf import get_pdf_content
from .pdf import get_pdf_metadata
from .pdf import merge_pdfs
from .txt import get_textfile_content
from .txt import validate_word
from .txt import word_frequency
from .txt import add_leading
from .txt import find_characters
from .txt import string_similarity
from .extract_rows import copy_pdf_column
from .extract_rows import extract_rows_from_page
from .extract_rows import ste_dict_rows
from .extract_rows import pdf_to_dictionary
from .extract_rows import pdf_page_to_img
from .filelist import list_files3
from .filelist import list_files2
from .filelist import list_files
from .filelist import get_extensions
from .docx_ import read_word_footers
from .docx_ import get_regex_string
from .docx_ import replace_copyright
from .docx_ import get_template_version
from .docx_ import docx_footer_replace
from .docx_ import docx_header_replace
from .docx_ import docx_content_replace
from .docx_ import replace_media
from .docx_ import word2pdf
from .docx_ import get_table_column_widths
from .docx_ import adjust_column_widths
from .docx_ import get_footer_type
from .docx_ import update_footer_table_widths
from .graphics import to_256
from .graphics import compress_img
from .graphics import resize_img
from .graphics import negative
from .graphics import blueprint
from .graphics import crop_image
from .graphics import get_average_color
from .glb2dracoglb import glb2dracoglb
from .filename_version import add_filename_version
from .filename_version import update_filename_version
from .filename_version import delete_filename_version
from .filename_version import increase_filename_version
from .batcher import batch
from .batcher import rename_illustrations
from .utils import UiLoader
from .utils import load_ui
from .multi import WorkerSignals
from .multi import Worker
from .svg_checks import check_line_widths
from .svg_checks import __check_line_widths
from .svg_checks import batch_check_line_widths
from .svg_checks import check_icn
from .svg_checks import check_icns
from .svg_checks import check_text_format
from .svg_checks import batch_check_text_format
from .svg_checks import check_illu_text
from .illustration_checks import illu_date_check
from .illustration_checks import check_cgm_details
from .illustration_checks import check_tif_details
from .vendor_list import get_chrome_driver_version
from .vendor_list import update_chrome_driver
from .vendor_list import VendorList
from .unit_table import UnrecognizedUnit
from .unit_table import UnitTable
from .repair_steps import clean_xml_tags
from .repair_steps import RepairSteps
from .consTableValidator import NoExcelSet
from .consTableValidator import NoOriginalTableFound
from .consTableValidator import DictError
from .consTableValidator import ConsumablesList
from .ataispec2200 import NoXmlSet
from .ataispec2200 import ConsumablesValidator
from .ataispec2200 import TorqueValuesValidator
from .ataispec2200 import cons_and_teds_checker
from .ataispec2200 import ipl_to_dict
from .ataispec2200 import _ipl_to_dict_excel
from .ataispec2200 import pgblk_9000_ted_checker
from .ataispec2200 import AtaNumbering
from .excel_ import download_excel
from .excel_ import get_excel_sheet_names
from .excel_ import colum_number_to_letter
from .excel_ import format_excel
from .archive import unarchive_file
from .archive import zip_word_folder
from .archive import zip_excel_folder
from .archive import zip_folder
from .archive import zipdir
from .archive import seven_unzip
from .make_library import get_manual_series
from .make_library import make_library
from .cgm2clearcgm import cgm2svgclear
from .cgm2svg import cgm2svg
from .automonkey_solution import process_file
from .automonkey_solution import calculate_version
from .automonkey_solution import calculate_image_similarity2
from .brex_checker import BrexNotFound
from .brex_checker import NoBrexDefined
from .brex_checker import clean_xpath
from .brex_checker import BrexChecker
from .clearcgm2svg import get_font_types
from .clearcgm2svg import linePrepend
from .clearcgm2svg import prettyPrint
from .clearcgm2svg import getContent
from .clearcgm2svg import oneDigit
from .clearcgm2svg import twoDigits
from .clearcgm2svg import threeDigits
from .clearcgm2svg import fourDigits
from .clearcgm2svg import strokeWidth
from .clearcgm2svg import stroke
from .clearcgm2svg import fill
from .clearcgm2svg import rbg2hex
from .clearcgm2svg import svgCircle
from .clearcgm2svg import calculate_points
from .clearcgm2svg import svgPolyline
from .clearcgm2svg import svgText
from .clearcgm2svg import preprocess_svg
from .clearcgm2svg import clearCGM2SVG
from .compare_raster import calculate_image_similarity
from .copying import copy_files
from .data_extraction import clean_word
from .file_info import get_file_size
from .filepath import clean_path
from .fits_and_clearences_checker import FCChecker
from .illustrations_checker import illustrationChecker
from .illustrations_checker import baselineReportFilter
from .pdf2raster import pdf2raster
from .procedure_checker import IPLChecker
from .raster2pdf import convert_image_to_pdf
from .raster2pdf import raster2pdf
from .reference_checker import RefChecker
from .reference_checker import CsnChecker
from .reference_checker import GraphicRefChecker
from .s1000d import get_s1000d_version
from .s1000d import get_references
from .s1000d import get_s1000d_refs
from .s1000d import get_4plus_refs
from .s1000d import get_2and3_refs
from .s1000d import get_brex_ref
from .s1000d import ref_dict_to_str
from .s1000d import find_document_by_reference
from .s1000d import get_dm_codes_from_dir
from .s1000d import get_dm_code_from_filename
from .s1000d import get_dm_code_from_xml
from .s1000d import validate_references
from .s1000d import ref_dict_to_dm_code_dict
from .s1000d import ref_dict_to_xml
from .s1000d import get_ddn
from .s1000d import validate_ddn
from .s1000d import increase_issue_number
from .s1000d import set_inwork
from .s1000d import read_dmodule
from .s1000d import read_procedure
from .s1000d import read_description
from .search_bar import include_search_bar
from .search_bar import _filter_widgets
from .svg_data import get_svg_data
from .svg2jpg import svg2jpg
from .svg2pdf import svg2pdf
from .time import pdf_date_to_format
from .xml_processing import delete_first_line
from .xml_processing import linearize_xml
from .xml_processing import get_schema_from_xml
from .xml_processing import get_xml_attribute
from .xml_processing import set_xml_attribute
from .xml_processing import get_xml_tag_content
from .xml_processing import set_xml_tag_content
from .xml_processing import replace_special_characters
from .python_func import get_object_attributes
from .python_func import get_object_methods
from .python_func import count_lines_of_functions
from .python_func import simple_pretty_print
from .python_func import check_brackets


import importlib
__all__ = [
    'AtaNumbering',
    'BrexChecker',
    'BrexNotFound',
    'ConsumablesList',
    'ConsumablesValidator',
    'CsnChecker',
    'DictError',
    'FCChecker',
    'GraphicRefChecker',
    'IPLChecker',
    'NoBrexDefined',
    'NoExcelSet',
    'NoOriginalTableFound',
    'NoXmlSet',
    'Punctuation',
    'RefChecker',
    'RepairSteps',
    'TorqueValuesValidator',
    'UiLoader',
    'UnitTable',
    'UnrecognizedUnit',
    'VendorList',
    'Worker',
    'WorkerSignals',
    'XmlSchemaValidator',
    '__check_line_widths',
    '_filter_widgets',
    '_ipl_to_dict_excel',
    'add_filename_version',
    'add_iplnom_to_smg',
    'add_iplnom_to_stp',
    'add_leading',
    'adjust_column_widths',
    'baselineReportFilter',
    'batch',
    'batch_check_line_widths',
    'batch_check_text_format',
    'blueprint',
    'calculate_image_similarity',
    'calculate_image_similarity2',
    'calculate_points',
    'calculate_version',
    'cgm2svg',
    'cgm2svgclear',
    'check_brackets',
    'check_cgm_details',
    'check_icn',
    'check_icns',
    'check_illu_text',
    'check_line_widths',
    'check_text_format',
    'check_tif_details',
    'clean_path',
    'clean_word',
    'clean_xml_tags',
    'clean_xpath',
    'clearCGM2SVG',
    'colum_number_to_letter',
    'compress_img',
    'cons_and_teds_checker',
    'convert_image_to_pdf',
    'copy_files',
    'copy_pdf_column',
    'count_lines_of_functions',
    'crop_image',
    'delete_filename_version',
    'delete_first_line',
    'docx_content_replace',
    'docx_footer_replace',
    'docx_header_replace',
    'download_excel',
    'estimate_illustration',
    'extract_rows_from_page',
    'fill',
    'find_characters',
    'find_document_by_reference',
    'format_excel',
    'fourDigits',
    'getContent',
    'get_2and3_refs',
    'get_4plus_refs',
    'get_average_color',
    'get_brex_ref',
    'get_chrome_driver_version',
    'get_ddn',
    'get_dm_code_from_filename',
    'get_dm_code_from_xml',
    'get_dm_codes_from_dir',
    'get_excel_sheet_names',
    'get_extensions',
    'get_file_size',
    'get_font_types',
    'get_footer_type',
    'get_manual_series',
    'get_object_attributes',
    'get_object_methods',
    'get_ocr_pdf_content',
    'get_pdf_content',
    'get_pdf_metadata',
    'get_references',
    'get_regex_string',
    'get_s1000d_refs',
    'get_s1000d_version',
    'get_schema_from_xml',
    'get_svg_data',
    'get_table_column_widths',
    'get_template_version',
    'get_textfile_content',
    'get_xml_attribute',
    'get_xml_tag_content',
    'glb2dracoglb',
    'illu_date_check',
    'illustrationChecker',
    'include_search_bar',
    'increase_filename_version',
    'increase_issue_number',
    'ipl_to_dict',
    'is_fullpage_illu',
    'linePrepend',
    'linearize_xml',
    'list_files',
    'list_files2',
    'list_files3',
    'load_ui',
    'make_library',
    'merge_pdfs',
    'negative',
    'ocr_pdf',
    'oneDigit',
    'pdf2raster',
    'pdf_date_to_format',
    'pdf_page_count',
    'pdf_page_to_img',
    'pdf_to_dictionary',
    'pgblk_9000_ted_checker',
    'prepare_estimation',
    'preprocess_svg',
    'prettyPrint',
    'process_file',
    'raster2pdf',
    'rbg2hex',
    'read_description',
    'read_dmodule',
    'read_procedure',
    'read_word_footers',
    'ref_dict_to_dm_code_dict',
    'ref_dict_to_str',
    'ref_dict_to_xml',
    'rename_illustrations',
    'replace_copyright',
    'replace_media',
    'replace_special_characters',
    'resize_img',
    'set_inwork',
    'set_xml_attribute',
    'set_xml_tag_content',
    'seven_unzip',
    'simple_pretty_print',
    'ste_dict_rows',
    'string_similarity',
    'stroke',
    'strokeWidth',
    'svg2jpg',
    'svg2pdf',
    'svgCircle',
    'svgPolyline',
    'svgText',
    'threeDigits',
    'to_256',
    'twoDigits',
    'unarchive_file',
    'update_chrome_driver',
    'update_filename_version',
    'update_footer_table_widths',
    'validate_ddn',
    'validate_references',
    'validate_word',
    'word2pdf',
    'word_frequency',
    'zip_excel_folder',
    'zip_folder',
    'zip_word_folder',
    'zipdir',
]

def __getattr__(name):
    modules = {
        'AtaNumbering': 'ataispec2200',
        'BrexChecker': 'brex_checker',
        'BrexNotFound': 'brex_checker',
        'ConsumablesList': 'consTableValidator',
        'ConsumablesValidator': 'ataispec2200',
        'CsnChecker': 'reference_checker',
        'DictError': 'consTableValidator',
        'FCChecker': 'fits_and_clearences_checker',
        'GraphicRefChecker': 'reference_checker',
        'IPLChecker': 'procedure_checker',
        'NoBrexDefined': 'brex_checker',
        'NoExcelSet': 'consTableValidator',
        'NoOriginalTableFound': 'consTableValidator',
        'NoXmlSet': 'ataispec2200',
        'Punctuation': 'xml_validation',
        'RefChecker': 'reference_checker',
        'RepairSteps': 'repair_steps',
        'TorqueValuesValidator': 'ataispec2200',
        'UiLoader': 'utils',
        'UnitTable': 'unit_table',
        'UnrecognizedUnit': 'unit_table',
        'VendorList': 'vendor_list',
        'Worker': 'multi',
        'WorkerSignals': 'multi',
        'XmlSchemaValidator': 'xml_validation',
        '__check_line_widths': 'svg_checks',
        '_filter_widgets': 'search_bar',
        '_ipl_to_dict_excel': 'ataispec2200',
        'add_filename_version': 'filename_version',
        'add_iplnom_to_smg': 'smg',
        'add_iplnom_to_stp': 'stp',
        'add_leading': 'txt',
        'adjust_column_widths': 'docx_',
        'baselineReportFilter': 'illustrations_checker',
        'batch': 'batcher',
        'batch_check_line_widths': 'svg_checks',
        'batch_check_text_format': 'svg_checks',
        'blueprint': 'graphics',
        'calculate_image_similarity': 'compare_raster',
        'calculate_image_similarity2': 'automonkey_solution',
        'calculate_points': 'clearcgm2svg',
        'calculate_version': 'automonkey_solution',
        'cgm2svg': 'cgm2svg',
        'cgm2svgclear': 'cgm2clearcgm',
        'check_brackets': 'python_func',
        'check_cgm_details': 'illustration_checks',
        'check_icn': 'svg_checks',
        'check_icns': 'svg_checks',
        'check_illu_text': 'svg_checks',
        'check_line_widths': 'svg_checks',
        'check_text_format': 'svg_checks',
        'check_tif_details': 'illustration_checks',
        'clean_path': 'filepath',
        'clean_word': 'data_extraction',
        'clean_xml_tags': 'repair_steps',
        'clean_xpath': 'brex_checker',
        'clearCGM2SVG': 'clearcgm2svg',
        'colum_number_to_letter': 'excel_',
        'compress_img': 'graphics',
        'cons_and_teds_checker': 'ataispec2200',
        'convert_image_to_pdf': 'raster2pdf',
        'copy_files': 'copying',
        'copy_pdf_column': 'extract_rows',
        'count_lines_of_functions': 'python_func',
        'crop_image': 'graphics',
        'delete_filename_version': 'filename_version',
        'delete_first_line': 'xml_processing',
        'docx_content_replace': 'docx_',
        'docx_footer_replace': 'docx_',
        'docx_header_replace': 'docx_',
        'download_excel': 'excel_',
        'estimate_illustration': 'estimation',
        'extract_rows_from_page': 'extract_rows',
        'fill': 'clearcgm2svg',
        'find_characters': 'txt',
        'find_document_by_reference': 's1000d',
        'format_excel': 'excel_',
        'fourDigits': 'clearcgm2svg',
        'getContent': 'clearcgm2svg',
        'get_2and3_refs': 's1000d',
        'get_4plus_refs': 's1000d',
        'get_average_color': 'graphics',
        'get_brex_ref': 's1000d',
        'get_chrome_driver_version': 'vendor_list',
        'get_ddn': 's1000d',
        'get_dm_code_from_filename': 's1000d',
        'get_dm_code_from_xml': 's1000d',
        'get_dm_codes_from_dir': 's1000d',
        'get_excel_sheet_names': 'excel_',
        'get_extensions': 'filelist',
        'get_file_size': 'file_info',
        'get_font_types': 'clearcgm2svg',
        'get_footer_type': 'docx_',
        'get_manual_series': 'make_library',
        'get_object_attributes': 'python_func',
        'get_object_methods': 'python_func',
        'get_ocr_pdf_content': 'ocr_pdf',
        'get_pdf_content': 'pdf',
        'get_pdf_metadata': 'pdf',
        'get_references': 's1000d',
        'get_regex_string': 'docx_',
        'get_s1000d_refs': 's1000d',
        'get_s1000d_version': 's1000d',
        'get_schema_from_xml': 'xml_processing',
        'get_svg_data': 'svg_data',
        'get_table_column_widths': 'docx_',
        'get_template_version': 'docx_',
        'get_textfile_content': 'txt',
        'get_xml_attribute': 'xml_processing',
        'get_xml_tag_content': 'xml_processing',
        'glb2dracoglb': 'glb2dracoglb',
        'illu_date_check': 'illustration_checks',
        'illustrationChecker': 'illustrations_checker',
        'include_search_bar': 'search_bar',
        'increase_filename_version': 'filename_version',
        'increase_issue_number': 's1000d',
        'ipl_to_dict': 'ataispec2200',
        'is_fullpage_illu': 'estimation',
        'linePrepend': 'clearcgm2svg',
        'linearize_xml': 'xml_processing',
        'list_files': 'filelist',
        'list_files2': 'filelist',
        'list_files3': 'filelist',
        'load_ui': 'utils',
        'make_library': 'make_library',
        'merge_pdfs': 'pdf',
        'negative': 'graphics',
        'ocr_pdf': 'ocr_pdf',
        'oneDigit': 'clearcgm2svg',
        'pdf2raster': 'pdf2raster',
        'pdf_date_to_format': 'time',
        'pdf_page_count': 'pdf',
        'pdf_page_to_img': 'extract_rows',
        'pdf_to_dictionary': 'extract_rows',
        'pgblk_9000_ted_checker': 'ataispec2200',
        'prepare_estimation': 'estimation',
        'preprocess_svg': 'clearcgm2svg',
        'prettyPrint': 'clearcgm2svg',
        'process_file': 'automonkey_solution',
        'raster2pdf': 'raster2pdf',
        'rbg2hex': 'clearcgm2svg',
        'read_description': 's1000d',
        'read_dmodule': 's1000d',
        'read_procedure': 's1000d',
        'read_word_footers': 'docx_',
        'ref_dict_to_dm_code_dict': 's1000d',
        'ref_dict_to_str': 's1000d',
        'ref_dict_to_xml': 's1000d',
        'rename_illustrations': 'batcher',
        'replace_copyright': 'docx_',
        'replace_media': 'docx_',
        'replace_special_characters': 'xml_processing',
        'resize_img': 'graphics',
        'set_inwork': 's1000d',
        'set_xml_attribute': 'xml_processing',
        'set_xml_tag_content': 'xml_processing',
        'seven_unzip': 'archive',
        'simple_pretty_print': 'python_func',
        'ste_dict_rows': 'extract_rows',
        'string_similarity': 'txt',
        'stroke': 'clearcgm2svg',
        'strokeWidth': 'clearcgm2svg',
        'svg2jpg': 'svg2jpg',
        'svg2pdf': 'svg2pdf',
        'svgCircle': 'clearcgm2svg',
        'svgPolyline': 'clearcgm2svg',
        'svgText': 'clearcgm2svg',
        'threeDigits': 'clearcgm2svg',
        'to_256': 'graphics',
        'twoDigits': 'clearcgm2svg',
        'unarchive_file': 'archive',
        'update_chrome_driver': 'vendor_list',
        'update_filename_version': 'filename_version',
        'update_footer_table_widths': 'docx_',
        'validate_ddn': 's1000d',
        'validate_references': 's1000d',
        'validate_word': 'txt',
        'word2pdf': 'docx_',
        'word_frequency': 'txt',
        'zip_excel_folder': 'archive',
        'zip_folder': 'archive',
        'zip_word_folder': 'archive',
        'zipdir': 'archive',
    }
    if name in modules:
        module = importlib.import_module(f'.{modules[name]}', __package__)
        return getattr(module, name)
    raise AttributeError(f'module {__name__} has no attribute {name}')
