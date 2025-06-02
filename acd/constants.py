"""This module provides general purpose constants"""

# Imaage Constants
FUNC_DICT = {
    'to_256': "to_256_",
    'compress_img': "compressed_",
    'resize_img': "resized_",
    'negative': "negative_",
    'blueprint': "blueprint_",
}
"""Functions Dictionary for Image Processing
"""


IMG_EXT = ['.jpg', '.png', '.jpeg', '.bmp', '.tiff', '.gif', '.tga', '.tif', '.psd', '.ai', '.svg', '.webp']
"""Image Exsensions
"""

TIFF_COMPRESSION = ['zlib', 'jpeg', 'deflate', 'none']
"""Tiff Compression Types"""


# S1000D Constants
S1000D_VERSION_REGEX = r"(S1000D_)(\d-\d)(/)"
DM_REF_REGEX = r"<(?:dmRef|refdm)>.*?</(?:dmRef|refdm)>"  # dmRef for Version 4+ and refdm for Version 2.3

DELIVERY_LIST_ITEM_REGEX = """(<deliveryListItem>
<dispatchFileName>)(.*?)(</dispatchFileName>)(<entityControlNumber>)?(.*?)?
(</entityControlNumber>)?(<issueInfo inWork=")?(\d\d)?(" issueNumber=")?(\d\d\d)?("/>)?
(</deliveryListItem>)""".replace("\n", "")

DM_ADDRESS_REGEX = """(<dmAddress><dmIdent>
<dmCode assyCode=")(.*?)(" disassyCode=")(.*?)(" disassyCodeVariant=")(.*?)("
 infoCode=")(.*?)(" infoCodeVariant=")(.*?)(" itemLocationCode=")(.*?)("
 modelIdentCode=")(.*?)(" subSubSystemCode=")(.*?)("
 subSystemCode=")(.*?)(" systemCode=")(.*?)(" systemDiffCode=")(.*?)(" ?/>)
(<language countryIsoCode=")(.*?)(" languageIsoCode=")(.*?)(" />)
(<issueInfo inWork=")(.*?)(" issueNumber=")(.*?)(" />)(</dmIdent>)
(<dmAddressItems>)
(<issueDate day=")(.*?)(" month=")(.*?)(" year=")(.*?)(" />)(<dmTitle>)
(<techName>)(.*?)(</techName>)
(<infoName>)(.*?)(</infoName>)
(</dmTitle></dmAddressItems>
</dmAddress>)
""".replace("\n", "")
# Group 2 = assyCode, 4 = disassyCode, 6 = disassyCodeVariant
# Group 8 = infoCode, 10 = infoCodeVariant, 12 = itemLocationCode
# Group 14 = modelIdentCode, 16 = subSubSystemCode, 18 = subSystemCode
# Group 20 = systemCode, 22 = systemDiffCode, 25 = countryIsoCode
# Group 27 = languageIsoCode, 30 = inWork, 32 = issueNumber
# Group 37 = issueDate_day, 39 = issueDate_month, 41 = issueDate_year
# Group 45 = techName, 48 = infoName


OLD_TO_NEW = {  # S1000D version <= 3.0 to >= 4.0
    "dmaddres": "dmAddress",
    "dmc": "dmCode",
    "modelic": "modelIdentCode",
    "sdc": "systemDiffCode",
    "chapnum": "systemCode",
    "section": "subSystemCode",
    "subsect": "subSubSystemCode",
    "subject": "assyCode",
    "discode": "disassyCode",
    "discodev": "disassyCodeVariant",
    "incode": "infoCode",
    "incodev": "infoCodeVariant",
    "itemloc": "itemLocationCode",
    "dmtitle": "dmTitle",
    "techname": "techName",
    "infoname": "infoName",
    "issno": "issueNumber",
    "inwork": "inWork",
    "type": "",
    "issdate": "issueDate",
    "year": "year",
    "month": "month",
    "day": "day",
    "refdm": "dmRef",
    "pmissuer": "pmIssuer",
    "pmnumber": "pmNumber",
    "pmvolume": "pmVolume",
    "pmtitle": "pmTitle",
    "sendid": "senderIdent",
    "recvid": "receiverIdent",
    "diyear": "yearOfDataIssue",
    "seqnum": "seqNumber",
    "ddnc": "ddnCode",
    "pmc": "pmCode",
}

OR_ITEMNUMBER_VALUES_REGEX = r"(the |each |an )([a-z \-]+)(\(\d{1,2}[a-z]?\-?\d{0,4}[a-z]?\))( or )(\(\d{1,2}[a-z]?\-?\d{0,4}[a-z]?\))([a-z \d:\-]*?)(\d+\.?\d*)( [+-±]\d+\.?\d* nm| dnm| cnm| danm)( \()(\d+\.?\d*)( [+-±]\d+\.?\d*)( lbf\.in\.?| lbf\.ft\.?)(\))"
NO_ITEMNUMBER_VALUES_REGEX = r"(the |each |an )([a-z \-]+)( to )([a-z \d:\-]*?)(\d+\.?\d*)( [+-±]\d+\.?\d* nm| dnm| cnm| danm)( \()(\d+\.?\d*)( [+-±]\d+\.?\d*)( lbf\.in\.?| lbf\.ft\.?)(\))"
ITEMNUMBER_VALUES_REGEX = r"(the |each |an )([a-z \-]+)(\(\d{1,2}[a-z]?\-?\d{0,4}[a-z]?\))([a-z \d:\-]*?)(\d+\.?\d*)( [+-±]\d+\.?\d* nm| dnm| cnm| danm)( \()(\d+\.?\d*)( [+-±]\d+\.?\d*)( lbf\.in\.?| lbf\.ft\.?)(\))"
TORQUE_VALUES_REGEX = r"(\d+\.?\d*)( [+-±]\d+\.?\d* nm| dnm| cnm| danm)( \()(\d+\.?\d*)( [+-±]\d+\.?\d*)( lbf\.in\.?| lbf\.ft\.?)(\))"

ITEMDATA_REGEX = r"(<itemdata.*?>)(.*?)(</itemdata>)"
ITEMNBR_REGEX = r"itemnbr=\"(.*?)\""
PNR_REGEX = r"<pnr>(.*?)</pnr>"
KWD_REGEX = r"<kwd>(.*?)</kwd>"
ADT_REGEX = r"<adt>(.*?)</adt>"
MFR_REGEX = r"<mfr>(.*?)</mfr>"
IPLNOM_REGEX = r"<iplnom>(.*?)</iplnom>"
