import pandas as pd
from json import dump
from json import load

from re import sub
from re import search
from re import findall

from fitz import open as pdf_open
from .txt import word_frequency

from .estimation import PAGEBLOCKS

def clean_word(word: str) -> str:
    """Clean a word from special characters.

    Args:
        word (str): word to be cleaned

    Returns:
        str: cleaned word
    """
    return word.lower().replace(".", "").replace("(", "").replace(")", "").replace(":", "").replace(",", "").replace(";", "")\
        .replace("?", "").replace("!", "").replace("-", "").replace("_", "").replace("[", "").replace("]", "")\
        .replace("{", "").replace("}", "").replace("=", "").replace("+", "").replace("*", "").replace("/", "")\
        .replace("\\", "").replace("|", "").replace("<", "").replace(">", "").replace("\"", "").replace("'", "")

# # load excel file
# excel = r'C:\Users\munteanu\Downloads\Liebherr Way of Working\Revision\CMM-D9893-BD50-32-61-01_003-01_EN_IFA.xlsx'
# df = pd.read_excel(
#     excel,
#     sheet_name=0)

# #print(df.head())

# df_dict = {}

# for row in df.iterrows():
#     df_dict[str(row[1][0]) + "-" + str(row[1][1])] = row[1][2] if not pd.isna(row[1][2]) else row[1][3]
# with open(excel.split(".")[0] + ".json", 'w', encoding="utf-8") as f:
#     dump(df_dict, f, ensure_ascii=False, indent=4)
if __name__ == "__main__":
    SECTION_TO_CATEGORY = {
        "TP-1": 0,
        "TP-2": 0,
        "RTR": 0,
        "SBL": 0,
        "LEP": 0,
        "TOC": 0,
        "LOI": 0,
        "LOT": 0,
        "INTRO": 0,
        "DESCRIPTION AND OPERATION": 1,
        "TESTING AND FAULT ISOLATION": 2,
        "DISASSEMBLY": 3,
        "CLEANING": 4,
        "INSPECTION/CHECK": 5,
        "REPAIR": 6,
        "ASSEMBLY": 7,
        "SPECIAL TOOLS; FIXTURES; EQUIPMENT AND CONSUMABLES": 8,
        "STORAGE": 9,
        "IPL": 10,
        "VL": 10,
        "NI": 10,
    }


    pdf = r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Revision\CMM-D9893-C091-32-31-11_000-01_EN.pdf"

    df = pd.DataFrame(columns=["Section", "Page", "ICN", "ICN Count", "Content Length", "Word Count", "Average Word Length", "Simple Words", "Middle Words", "Complex Words", "Result"])

    with open(pdf.replace(".pdf", ".json"), 'r', encoding="utf-8") as f:
        result = load(f)

    important_words = []
    important_dict = {
        "Complex": [],
        "Middle": [],
        "Simple": []
    }

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

        word_freq = word_frequency(page_text.lower(), False, False)

        if page_match:
            section = page_match.group(2).strip().replace(",", "")
            page_number = page_match.group(4).strip().replace("Page ", "")
        elif "TP1" in page_text:
            section = "TP"
            page_number = "1"
        elif "TP2" in page_text:
            section = "TP"
            page_number = "2"

        if section != "":
            pass
        if len(page_number) < 4:
            section = "DESCRIPTION AND OPERATION"
        elif len(page_number) == 4 or page_number.startswith("15") or page_number.startswith("11"):
            section = PAGEBLOCKS[int(page_number[:len(page_number) - 3])]

        for pair in word_freq:
            word = pair[0].lower().replace(".", "").replace("(", "").replace(")", "").replace(":", "").replace(",", "").replace(";", "")\
                .replace("?", "").replace("!", "").replace("-", "").replace("_", "").replace("[", "").replace("]", "")\
                .replace("{", "").replace("}", "").replace("=", "").replace("+", "").replace("*", "").replace("/", "")\
                .replace("\\", "").replace("|", "").replace("<", "").replace(">", "").replace("\"", "").replace("'", "")
            digits = len(sub(r"[^0-9]", "", word))

            if ((pair[1] > 2 and len(word) > 4) or len(word) > 6) and digits == 0:
                important_words.append(word)
                if section + "-" + page_number in result:
                    if result[section + "-" + page_number] == "Complex":
                        important_dict["Complex"].append(word)
                    elif result[section + "-" + page_number] == "Middle":
                        important_dict["Middle"].append(word)
                    else:
                        important_dict["Simple"].append(word)

        if section in SECTION_TO_CATEGORY:
            new_section = SECTION_TO_CATEGORY[section]
        else:
            new_section = 0
        page = int(page_number)

        icn = 1 if icn_match else 0

        with open(r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Creation\complexity.json", 'r', encoding="utf-8") as f:
            important_dict_updated = load(f)

        complex_words = len([word for word in word_freq if clean_word(word[0]) in important_dict_updated["Complex"]])
        middle_words = len([word for word in word_freq if clean_word(word[0]) in important_dict_updated["Middle"]])
        simple_words = len([word for word in word_freq if clean_word(word[0]) in important_dict_updated["Simple"]])

        if section + "-" + str(page_number) in result:
            res = 3 if result[section + "-" + str(page_number)] == "Complex" else 2 if result[section + "-" + str(page_number)] == "Middle" else 1
            df.loc[ind] = [new_section, page_number, icn, icn_matches, content_length, word_count, average_word_length, simple_words, middle_words, complex_words, res]
        else:
            df.loc[ind] = [new_section, page_number, icn, icn_matches, content_length, word_count, average_word_length, simple_words, middle_words, complex_words, 1]
        # Simple = 1
        # Middle = 2
        # Complex = 3

    df.to_excel(pdf.split(".")[0] + ".xlsx", index=False)

    # with open(pdf.replace(".pdf", ".txt"), 'w', encoding="utf-8") as f:
    #     for word in important_words:
    #         f.write(word + "\n")



    # for word in important_dict["Complex"]:
    #     if word not in important_dict["Middle"] and word not in important_dict["Simple"] and word not in important_dict_updated["Complex"]:
    #         important_dict_updated["Complex"].append(word)
    # for word in important_dict["Middle"]:
    #     if word not in important_dict["Complex"] and word not in important_dict["Simple"] and word not in important_dict_updated["Middle"]:
    #         important_dict_updated["Middle"].append(word)
    # for word in important_dict["Simple"]:
    #     if word not in important_dict["Complex"] and word not in important_dict["Middle"] and word not in important_dict_updated["Simple"]:
    #         important_dict_updated["Simple"].append(word)


    # with open(r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Creation\complexity.json", 'w', encoding="utf-8") as f:
    #     dump(important_dict_updated, f, ensure_ascii=False, indent=4)