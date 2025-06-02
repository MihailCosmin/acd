
from os.path import join
from os.path import isfile
from shutil import copyfile

from pandas import DataFrame
import pandas as pd

from json import load
from json import dump

from .filelist import list_files


def get_svg_data(svg: str) -> dict:
    result = {
        "g": 0,
        "a": 0,
        "path": 0,
        # "rect": 0,  # was always the same
        "circle": 0,
        # "ellipse": 0,
        # "line": 0,
        "polyline": 0,
        "polygon": 0,
        "text": 0,
        # "marker": 0,
        # "clipPath": 0,
        "points": 0,
        "style": 0,
        "d": 0,
    }

    with open(svg, 'r', encoding='utf-8') as f:
        svg_data = f.read()
        for key in result.keys():
            result[key] = svg_data.count(f"<{key} ") + svg_data.count(f"<{key}>") + svg_data.count(f"{key}=\"")
    result["character_count"] = len(svg_data)
    return result


if __name__ == "__main__":
    with open(r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Creation\illus.json", "r", encoding="utf-8") as f:
        illus = load(f)

    svg_dir = r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Creation\SVG"

    df = DataFrame(columns=["name", "g", "a", "path", "circle", "polyline", "polygon", "text", "points", "style", "d", "character_count", "result"])

    for cgm, res in illus.items():
        svg = cgm.replace(".cgm", ".svg")
        if isfile(join(svg_dir, svg)):
            svg_data = get_svg_data(join(svg_dir, svg))
            svg_data["name"] = cgm
            svg_data["result"] = 1 if res == "Simple" else 2 if res == "Middle" else 3
            # df = df.append(svg_data, ignore_index=True)
            # use pd.concat instead of df.append
            df = pd.concat([df, pd.DataFrame(svg_data, index=[0])], ignore_index=True)

    # save df to excel
    df.to_excel(r"C:\Users\munteanu\Downloads\Liebherr Way of Working\Creation\svg_data.xlsx", index=False)