from os.path import join
from os.path import isdir
from os.path import isfile
from os.path import splitext
from os.path import basename
from os.path import dirname
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF


from subprocess import check_output

from .filelist import list_files

INKSCAPE = '\"C:\\Program Files\\Inkscape\\bin\\inkscape\"'

def svg2pdf(item: str, overwrite: bool = False, inkscape: bool = False) -> None:
    """Converts a svg file to pdf using inkscape

    Args:
        item (str): path to a svg file or a directory containing cgm files
    """
    if isfile(item):
        ext = splitext(item)[1]
        pdf = join(dirname(item), basename(item).replace(ext, ".pdf"))

        if inkscape:
            check_output(f'{INKSCAPE} --export-filename="{pdf}" "{item}"', shell=True)
        else:
            drawing = svg2rlg(item)
            renderPDF.drawToFile(drawing, pdf)

    elif isdir(item):
        for svg in list_files(item, True, ["svg"]):
            ext = splitext(svg)[1]
            pdf = join(dirname(svg), basename(svg).replace(ext, ".pdf"))
            if isfile(pdf) and not overwrite:
                continue
            # check_output(f"{INKSCAPE} {svg} --export-type=pdf --export-filename={out_fname}", shell=True)
            # check_output(f'{INKSCAPE} "{svg}" --export-type=pdf', shell=True)
            if inkscape:
                check_output(f'{INKSCAPE} "{svg}" --export-filename="{pdf}" 2> nul', shell=True)
                # Cosmin 2> nul is used to hide the output of the command
            else:
                drawing = svg2rlg(svg)
                renderPDF.drawToFile(drawing, pdf)

if __name__ == "__main__":
    svg2pdf(r"D:\TD\LLI\Liebherr S1000D\Test.the.dot\DDN-LIAERBA77-D9893-D9893-2018-02387")
