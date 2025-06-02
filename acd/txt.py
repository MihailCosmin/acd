"""This module provides functions to read and write to text files.
"""

# English Dictionary
from difflib import SequenceMatcher
import sys
if sys.version_info < (3, 12):
    import enchant

from Levenshtein import distance
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity

ENCODINGS = [
    'utf-8', 'iso-8859-1', 'cp1252', 'utf-16', 'utf-16-be', 'utf-16-le', 'utf-32', 'utf-32-be', 'utf-32-le', 'utf-7',
    'utf-8-sig', 'utf-ebcdic', 'utf-hex', 'utf-html', 'utf-java', 'utf-js', 'utf-json', 'utf-latex', 'utf-lzw',
    'utf-mac', 'utf-marshal', 'utf-moz', 'utf-text', 'utf-xml', 'utf-yaml', 'ansi', 'ascii', 'big5', 'big5hkscs',
    'cp037', 'cp424', 'cp437', 'cp500', 'cp737', 'cp775', 'cp850', 'cp852', 'cp855', 'cp857', 'cp860', 'cp861',
    'cp862', 'cp863', 'cp864', 'cp865', 'cp866', 'cp869', 'cp874', 'cp875', 'cp932', 'cp949', 'cp950', 'cp1006',
    'cp1026', 'cp1140', 'cp1250', 'cp1251', 'cp1253', 'cp1254', 'cp1255', 'cp1256', 'cp1257', 'cp1258',
    'euc-jp', 'euc-jis-2004', 'euc-jisx0213', 'euc-kr', 'gb2312', 'gbk', 'gb18030', 'hz', 'iso2022-jp',
    'iso2022-jp-1', 'iso2022-jp-2', 'iso2022-jp-2004', 'iso2022-jp-3', 'iso2022-jp-ext', 'iso2022-kr', 'iso8859-1',
    'iso8859-2', 'iso8859-3', 'iso8859-4', 'iso8859-5', 'iso8859-6', 'iso8859-7', 'iso8859-8', 'iso8859-9',
    'iso8859-10', 'iso8859-11', 'iso8859-13', 'iso8859-14', 'iso8859-15', 'iso8859-16', 'johab', 'koi8-r',
    'koi8-t', 'koi8-u', 'mac-cyrillic', 'mac-greek', 'mac-iceland', 'mac-latin2', 'mac-roman', 'mac-turkish',
    'ptcp154', 'shift-jis', 'shift-jis-2004', 'shift-jisx0213', 'utf-16-be', 'utf-16-le', 'utf-32-be', 'utf-32-le',
    'utf-7', 'utf-8-sig', 'utf-ebcdic', 'utf-hex', 'utf-html', 'utf-java', 'utf-js', 'utf-json', 'utf-latex',
    'utf-lzw', 'utf-mac', 'utf-marshal', 'utf-moz', 'utf-text', 'utf-xml', 'utf-yaml', 'windows-1250', 'windows-1251',
    'windows-1252', 'windows-1253', 'windows-1254', 'windows-1255', 'windows-1256', 'windows-1257', 'windows-1258',
    'x-mac-cyrillic', 'x-mac-greek', 'x-mac-iceland', 'x-mac-roman', 'x-mac-turkish', 'x-user-defined', 'x-utf-16le',
    'x-utf-32be', 'x-utf-32le', 'x-utf-7', 'x-utf-8-sig', 'x-utf-ebcdic', 'x-utf-hex', 'x-utf-html', 'x-utf-java',
    'x-utf-js', 'x-utf-json', 'x-utf-latex', 'x-utf-lzw', 'x-utf-mac', 'x-utf-marshal', 'x-utf-moz', 'x-utf-text',
    'x-utf-xml', 'x-utf-yaml'
]

GREEK_CHARS = ["Α", "Β", "Γ", "Δ", "Ε", "�", "Η", "Θ", "Ι", "Κ", "Λ", "Μ", "Ν", "Ξ", "Ο", "Π", "Ρ", "Σ", "Τ", "Υ", "Φ", "Χ", "Ζ", "Ψ", "Ω"]

def get_textfile_content(file_path: str) -> str:
    """
    This function takes a text file and returns its content.
    file_path: str - path to the text file

    Returns: str - content of the text file
    """
    for encoding in ENCODINGS:
        try:
            with open(file_path, encoding=encoding) as file_:
                return file_.read()
        except UnicodeDecodeError:
            continue
        except FileNotFoundError:
            return file_path
    raise UnicodeDecodeError(f"Unable to decode file: {file_path}")


def validate_word(word: str, locale: str = 'en-US') -> bool:
    """
    This function takes a word and returns whether it is valid or not.
    word: str - word to be validated
    locale: str - locale to be used for validation
        Ex: ['en_BW', 'en_AU', 'en_BZ', 'en_GB', 'en_JM', 'en_DK', 'en_HK', 'en_GH', 'en_US', 'en_ZA', 'en_ZW',
        'en_SG', 'en_NZ', 'en_BS', 'en_AG', 'en_PH', 'en_IE', 'en_NA', 'en_TT', 'en_IN', 'en_NG', 'en_CA']
        You can also install new dictionaries by downloading them from:
            https://cgit.freedesktop.org/libreoffice/dictionaries/tree/
        Download both the .dic and .aff files. Rename them to the locale name. Ex: de-DE.dic and de-DE.aff
        And then add them to pyenchant installation folder:
            Ex: ....Python39\\Lib\\site-packages\\enchant\\data\\mingw64\\share\\enchant\\hunspell

        You can also specify "all" for the locales. Then all the dictionaries will be used.
        Or you can specify multiple locales, separated by commas.
            Ex: 'en_US,en_GB,en_AU,de-DE'

    Returns: bool - whether the word is valid or not
    """
    if sys.version_info < (3, 12):
        if locale in ('all', 'ALL', 'All'):
            return any(dictionary.check(word) for dictionary in enchant.list_languages())
        if ',' in locale:
            return any(enchant.Dict(loc.strip()).check(word) for loc in locale.split(','))
        return enchant.Dict(locale).check(word)
    # Not implemented for Python 3.12 and above
    return None


def word_frequency(text: str, word_limit: int = 0, validate_words: bool = False, sort_descending: bool = True) -> dict:
    """
    This function takes a text and returns a dictionary of words and their frequency.
    text: str - text for which the word frequency is calculated
    word_limit: int - minimmum number of letters in a word
    sort_descending: bool - sort the dictionary in descending order. Default is True

    Returns: dict - dictionary with the word frequency
    """

    words = text.split()
    words = [word for word in words if len(word) > word_limit]
    words = [word for word in words if validate_word(word)] if validate_words else words
    word_freq = {}
    for word in words:
        if word in word_freq:
            word_freq[word] += 1
        else:
            word_freq[word] = 1
    if sort_descending:
        return sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
    return word_freq

def add_leading(string: str, leading: str, max: int = 0) -> str:
    """Add leading characters to a string until it reaches a certain length.

    Args:
        string (str): string to be padded
        leading (str): leading character to be added
        max (int, optional): Maximum length of the string. Defaults to 0.

    Returns:
        str: _description_
    """
    while len(string) < max:
        string = leading + string
    return string

def find_characters(txt: str, chars: list) -> dict:
    """Find if a list of characters is present in a text.

    Args:
        txt (str): text to be searched
        chars (list): list of characters to be searched for

    Returns:
        dict: dictionary with the characters and their positions
    """
    txt = get_textfile_content(txt)
    result = {}
    for char in chars:
        result[char] = False
        if char in txt:
            result[char] = True
    for char in txt:
        if not char.isascii() or not char.isprintable():
            result[char] = True
    return result

def string_similarity(str1: str, str2: str) -> float:
    """Calculate the similarity between two strings.

    Args:
        str1 (str): First string
        str2 (str): Second string

    Returns:
        float: similarity between the two strings
    """
    # Levenshtein Distance (Edit Distance)
    # The Levenshtein distance measures the minimum number of single-character edits
    # (insertions, deletions, or substitutions) required to change one string into another.
    # You can use the python-Levenshtein library to calculate it.
    dist = distance(str1, str2)
    similarity1 = 1 - (dist / max(len(str1), len(str2)))

    # SequenceMatcher (from the difflib module):
    # The SequenceMatcher class in the difflib module can
    # be used to find the similarity ratio between two strings.
    similarity2 = SequenceMatcher(None, str1, str2).ratio()

    # Jaccard Similarity:
    # Jaccard similarity measures the similarity between two sets.
    # You can calculate it by dividing the size of the intersection
    # of the sets by the size of the union of the sets.
    set1 = set(str1.split())
    set2 = set(str2.split())
    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))

    similarity3 = intersection / union

    # Cosine Similarity:
    # Cosine similarity is often used for comparing documents or text data
    # by treating the strings as vectors in a high-dimensional space.
    try:
        vectorizer = CountVectorizer().fit_transform([str1, str2])
        vectors = vectorizer.toarray()
        similarity4 = cosine_similarity([vectors[0]], [vectors[1]])[0][0]
    except ValueError:
        similarity4 = (similarity1 + similarity2 + similarity3) / 3
    return (similarity1 + similarity2) / 2

if __name__ == "__main__":
    # print(find_characters(
    #     r"C:\Users\munteanu\Downloads\PMC-LIAEREP30-D9893-05448-01_000-01_SX-US(4).XML",
    #     GREEK_CHARS
    # ))
    print(find_characters(
        r"PMC-LIAEREP30-D9893-05448-01_000-01_SX-US.XML",
        GREEK_CHARS
    ))
