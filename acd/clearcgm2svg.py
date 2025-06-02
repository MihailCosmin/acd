from lxml import etree
from yattag import Doc, indent
from fileinput import FileInput
from re import search
from re import sub
from re import DOTALL
from re import findall

doc, tag, text = Doc().tagtext()

LINE_TYPES = {
    1: ("solid", None),
    2: ("dash", "5,5"),
    3: ("dot", "40,80"),  # this to be checked manually, for a stroke-width of 8, the correct value is 40,80 - Maybe x5, x10 ???
    4: ("dash-dot", "5,2,1,2"),
    5: ("dash-dot-dot", "5,2,1,2,1,2")
}

def get_font_types(content: str) -> dict:
    FONT_REGEX = r"fontlist.*?';\n"
    font_types = {}
    fontlist = findall(FONT_REGEX, content, DOTALL)[0]
    for ind, font in enumerate(fontlist.split("', '"), 1):
        font_types[ind] = font.replace("'", "")
    return font_types

def linePrepend(filename, line):
    with open(filename, 'r+') as f:
        content = f.read()
        f.seek(0, 0)
        f.write(line.rstrip('\r\n') + '\n' + content)


def prettyPrint(outputFile):
    x = etree.parse(outputFile)
    with open(outputFile, 'w', encoding="utf8") as f:
        f.write(str(etree.tostring(x, pretty_print=False)))

    with FileInput(outputFile, inplace=True) as file:
        for line in file:
            correctedLine = line.replace(">\\n", ">")\
                .replace("\\n", "")\
                .replace("b'<", "<")\
                .replace(">'", ">")\
                .replace("    ", "")

            print(correctedLine, end='')

    x = etree.parse(outputFile)
    with open(outputFile, 'w', encoding="utf8") as f:
        f.write(etree.tostring(x, pretty_print=True).decode())


def getContent(string, start, end):
    try:
        result = search(start + '(.*?)' + end, string)
        return result.group(1)
    except:
        return string


def oneDigit(number):
    if int(number) == 0:
        return str(0)
    else:
        return format(float(number), '.1f')


def twoDigits(number):
    if int(number) == 0:
        return str(0)
    else:
        return format(float(number), '.2f')


def threeDigits(number):
    if int(number) == 0:
        return str(0)
    else:
        return format(float(number), '.3f')

def fourDigits(number):
    if int(number) == 0:
        return str(0)
    else:
        return format(float(number), '.4f')

def strokeWidth(cgmLine, cgmLines, strokeWidVar):
    if cgmLine.strip()[0:10] == "linewidth ":
        return getContent(cgmLine, "linewidth ", ";")
    elif strokeWidVar is not False:
        return strokeWidVar
    elif strokeWidVar is False:
        picBody = getContent(cgmLines, "BEGPICBODY;", "ENDPIC;")
        for line in picBody:
            if line.strip()[0:9] == "EDGEWIDTH":
                strokeWidVar = getContent(line, "EDGEWIDTH ", ";")
                return strokeWidVar
            else:
                return False


def stroke(cgmLine, cgmLines, strokeVar):
    if strokeVar is False:
        picBody = getContent(cgmLines, "BEGPICBODY;", "ENDPIC;")
        for line in picBody:
            if line.strip()[0:8] == "EDGECOLR":
                strokeVar = str(rbg2hex(getContent(line, "EDGECOLR ", ";")))
                return strokeVar
    else:
        return "#000000"


def fill(cgmLine, cgmLines, fillVar):
    if fillVar is False:
        picBody = getContent(cgmLines, "BEGPICBODY;", "ENDPIC;")
        for line in picBody:
            if line.strip()[0:8] == "fillcolr":
                fillVar = str(rbg2hex(getContent(line, "fillcolr ", ";")))
                return fillVar
            else:
                return False  # "#ffffff"
    else:
        return False  # "#ffffff"


def rbg2hex(string):  # Ex: 255 255 255
    rgb = "(" + string.replace(" ", ",") + ")"
    rgb = eval(rgb)
    return '#%02x%02x%02x' % rgb


def svgCircle(cgmLine, strokeWidVar, strokeVar, fillVar):
    if cgmLine.strip()[0:6] == "CIRCLE":
        circleAtt = {}
        circleAtt["style"] = "stroke-width: " + str(strokeWidVar) + "; "
        circleAtt["style"] = circleAtt["style"] + \
            " stroke: " + str(strokeVar) + "; "
        if fillVar is not False:
            circleAtt["style"] = circleAtt["style"] + \
                " fill: " + str(fillVar) + "; "
        circleAtt["cx"] = getContent(cgmLine, "CIRCLE \(", ",")
        circleAtt["cy"] = getContent(cgmLine, ",", "\)")
        circleAtt["r"] = getContent(cgmLine, "\) ", ";")
        with tag("circle", **circleAtt):
            pass

def calculate_points(points_str: str, height: int) -> str:
    """This function calculates the points for the polyline
    This basically inverts the y-axis. You can test how it would work without

    Args:
        points_str (str): Points string
        height (int): Height of the image

    Returns:
        str: The new points string
    """
    _ = points_str.split(" ")
    first = [part.split(",")[0] for part in _]
    second = [part.split(",")[1] for part in _]

    second = [str(height - int(float(part))) for part in second]

    # add 4 digits to the second part
    second = [fourDigits(part) for part in second]

    # join the two lists
    re_str = " ".join([f"{first[i]},{second[i]}" for i in range(len(first))])

    # print(f"First: {first},\nSecond: {second}")
    # print(f"Result: {re_str}")
    return re_str


def svgPolyline(cgmLine, strokeWidVar, strokeVar, fillVar, linetype, height):
    if cgmLine.strip()[0:4] == "LINE":
        polylineAtt = {}
        polylineAtt["style"] = "stroke-width: " + strokeWidVar + "; "

        if linetype is not None:
            polylineAtt["style"] = polylineAtt["style"] + \
                " stroke-dasharray: " + linetype + "; "

        polylineAtt["style"] = polylineAtt["style"] + \
            " stroke: " + strokeVar + "; "
        if fillVar is not False:
            polylineAtt["style"] = polylineAtt["style"] + \
                " fill: " + fillVar + "; "
        polylineAtt["points"] = calculate_points(getContent(
            cgmLine, "LINE ", ";").replace("(", "").replace(")", ""), int(float(height)))
        with tag("polyline", **polylineAtt):
            pass


def svgText(cgmLine, textfont: str = "Arial", textsize: str = "140", svh_height: float = None):
    """_summary_

    Args:
        cgmLine (_type_): _description_
        textfont (str, optional): _description_. Defaults to "Arial".
        textsize (str, optional): _description_. Defaults to "140".
        svh_height (float, optional): _description_. Defaults to None.
    """
    # TODO: Find how to convert dimension from CGM to SVG
    # 97 = 140 = 0.0135124
    # 69 = 100 = 0.00961194

    textsize = str(round(float(textsize) * 1.4463, 0))

    if cgmLine.strip()[0:5] == "TEXT ":
        textAtt = {}
        textAtt["x"] = getContent(cgmLine, "TEXT \(", ",")
        textAtt["y"] = str(svh_height - float(getContent(cgmLine, ",", "\) ")))
        textAtt["font-size"] = textsize
        if "Oblique" in textfont:
            textfont = textfont.replace("Oblique", "")
            textAtt["font-style"] = "oblique"
        textAtt["font-family"] = textfont
        with tag("text", **textAtt):
            text(getContent(cgmLine, " '", "';"))


def preprocess_svg(content: str) -> str:
    """DISJTLINE == DISJOINT POLYLINE

    Args:
        content (str): The content of the cgm file

    Returns:
        str: The content of the cgm file with the disjoint lines replaced
    """
    DISJTLINE_REGEX = r"DISJTLINE[0-9,\. \(\)\n]*?;"
    # replace all \n with empty string in all entries in the content
    for disjtline in findall(DISJTLINE_REGEX, content, DOTALL):
        disjtline_new = disjtline.replace("\n", " ")
        disjtline_new = disjtline_new.replace("  ", " ")
        disjtline_new = disjtline_new.replace("DISJTLINE", " LINE")
        content = content.replace(disjtline, disjtline_new)
    return content


def clearCGM2SVG(file):
    svgAtt = {}

    with open(file, mode='r') as f_in:
        cgm = f_in.read()

        content = preprocess_svg(cgm)
        with open(file.replace(".cgm", "_1.cgm"), "w") as f:
            f.write(content)

        # cgmLines = open(file, mode='r').readlines()
        cgmLines = content.split("\n")
        content = getContent(cgm, "vdcext ", ";")

        firstPart = getContent(content, "\(", "\) \(")
        secondPart = getContent(content, "\) \(", "\)")

        firstX = getContent(firstPart, "^", ",")
        firstY = getContent(firstPart, ",", "$")

        secondX = getContent(secondPart, "^", ",")
        secondY = getContent(secondPart, ",", "$")

        svgAtt["xmlns"] = "http://www.w3.org/2000/svg"
        svgAtt["xmlns:xlink"] = "http://www.w3.org/1999/xlink"
        svgAtt["xmlns:xlink"] = "http://www.w3.org/1999/xlink"
        svgAtt["xmlns:ev"] = "http://www.w3.org/2001/xml-events"
        svgAtt["xmlns:webcgm"] = "http://www.w3.org/Graphics/WebCGM"

        svgAtt["width"] = threeDigits(float(secondX) / 10) + "px"  # conversion factor might be / 10.05005 or * 0.099502
        svgAtt["height"] = threeDigits(float(secondY) / 10) + "px"  # conversion factor might be / 10.05005 or * 0.099502
        svgAtt["viewBox"] = twoDigits(int(float(firstX))) + " " + twoDigits(int(float(
            firstY))) + " " + twoDigits(int(float(secondX))) + " " + twoDigits(int(float(secondY)))

        with tag('svg', **svgAtt):
            with tag('title'):
                text(getContent(file, "\\\\", "$"))
            with tag('desc'):
                text("ALTHOM GmbH CGM to SVG convertor")
            with tag('defs'):
                with tag('style', type="text/css"):  # CDATA part apparently is not needed
                    doc.asis(r"""

      <![CDATA[
polyline
{
  stroke-linejoin:round;
  stroke-linecap:round;
  stroke-miterlimit:32767;
  fill:none;
  stroke-dasharray:none;
  stroke:none;
}
path
{
  stroke-linejoin:round;
  stroke-linecap:round;
  stroke-miterlimit:32767;
  fill:none;
  stroke-dasharray:none;
  stroke:none;
  fill-rule:evenodd;
}
text
{
  font-family: Helvetica;
}
tspan
{
  font-family: Helvetica;
}
polygon
{
  fill-rule:evenodd;
  fill:none;
  stroke-linejoin:round;
  stroke-linecap:round;
  stroke-miterlimit:32767;
  stroke:none;
  stroke-dasharray:none;
}
circle
{
  fill-rule:evenodd;
  fill:none;
  stroke-linejoin:round;
  stroke-linecap:round;
  stroke-miterlimit:32767;
  stroke:none;
  stroke-dasharray:none;
}
]]>
                         """)
            polygonAtt = {}
            polygonAtt["style"] = "fill:#ffffff;"
            polygonAtt["points"] = f"0,{secondY} 0,0 {secondX},0 {secondX},{secondY} 0,{secondY}"
            with tag('polygon', **polygonAtt):
                pass

            strokeWidVar = False
            strokeVar = False
            fillVar = False
            textfontindex = 1
            textfonttypes = get_font_types(cgm)
            textsize = "140"

            linetype = 1

            # Check how some lines are created in different layers and on top of each other
            # This is the reason some lines are not visible
            # Cosmin: 29.02.2024 - This was happening because everything had white fill
            for ind, cgmLine in enumerate(cgmLines, 1):
                if "TEXTFONTINDEX" in cgmLine:
                    textfontindex = int(cgmLine.split("TEXTFONTINDEX ")[1].replace(";", ""))
                if "charheight" in cgmLine:
                    textsize = str(int(float(cgmLine.split("charheight ")[1].replace(";", ""))))
                textfont = textfonttypes[textfontindex]
                previous_line = ""
                if ind > 1:
                    previous_line = cgmLines[ind - 2]
                if "linetype" in previous_line:
                    linetype = int(previous_line.split("linetype ")[1].replace(";", ""))
                linetype_name = LINE_TYPES[linetype][1]

                strokeWidVar = strokeWidth(cgmLine, cgmLines, strokeWidVar)
                strokeVar = stroke(cgmLine, cgmLines, strokeVar)
                fillVar = fill(cgmLine, cgmLines, fillVar)

                svgCircle(cgmLine, strokeWidVar, strokeVar, fillVar)
                svgPolyline(cgmLine, strokeWidVar, strokeVar, fillVar, linetype_name, twoDigits(int(float(secondY))))
                svgText(cgmLine, textfont, textsize, float(secondY))

        result = indent(
            doc.getvalue(),
            indentation='    ',
            indent_text=True
        )

        outputFile = file.replace(".cgm", "_2.svg").replace(".CGM", "_2.svg")

        with open(outputFile, "w") as svg:
            svg.write(result)

        prettyPrint(outputFile)
        linePrepend(outputFile, "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>")
        print('Done!')

