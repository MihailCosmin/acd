from ast import literal_eval


def get_object_attributes(obj, show: bool = False) -> list:
    """Returns a list of attributes of an object

    Args:
        obj (_type_): The object to get attributes from
        show (bool, optional): If it should print the attributes to console . Defaults to False.

    Returns:
        list: The list of attributes
    """
    lst = []
    for attribute_name in dir(obj):
        try:
            if not callable(getattr(obj, attribute_name)):
                lst.append(attribute_name)
                if show:
                    print(attribute_name)
        except IndexError:
            lst.append(attribute_name)
            if show:
                print(attribute_name)
    if len(lst) == 0:
        return None


def get_object_methods(obj, show: bool = False) -> list:
    """Returns a list of methods of an object

    Args:
        obj (_type_): The object to get methods from
        show (bool, optional): If it should print the methods to console . Defaults to False.

    Returns:
        list: The list of methods
    """
    lst = []
    for method_name in dir(obj):
        try:
            if callable(getattr(obj, method_name)):
                lst.append(method_name)
                if show:
                    print(method_name)
        except IndexError:
            lst.append(method_name)
            if show:
                print(method_name)
    if len(lst) == 0:
        return None
    return lst


def count_lines_of_functions(file_path):
    """Gets the lines of code for each function and displays them in the terminal

    Args:
        file_path (_type_): Path to the python file to get the numer of lines from
    """
    with open(file_path, 'r') as f:
        code = f.readlines()
    functions = {}
    in_function = False
    current_function = ""
    for line in code:
        if "def" in line:
            in_function = True
            current_function = line.split("def")[1].split("(")[0].strip()
            functions[current_function] = 0
        if in_function:
            functions[current_function] += 1
        if "return" in line or "raise" in line:
            in_function = False
    for function_name, lines_of_code in functions.items():
        print(f"{function_name}: {lines_of_code} lines of code")


def simple_pretty_print(input_string: str) -> str:
    """Takes a string as input and returns a pseudo pretty-printed version 
    of it if it has the structure of either a list or a dictionary

    Args:
        input_string (str): string to be pseudo pretty-printed

    Returns:
        str: pseudo pretty-printed string
    """
    try:
        parsed_data = literal_eval(input_string)
        if isinstance(parsed_data, dict):
            pretty_dict = ""
            for key, value in parsed_data.items():
                pretty_dict += "  " + str(key) + ": " + str(value) + "," + "\n"
            return "{\n" + pretty_dict + "}"
        elif isinstance(parsed_data, list):
            pretty_list = ""
            for item in parsed_data:
                pretty_list += "  " + str(item) + "," + "\n"
            return "[\n" + pretty_list + "]"
        return input_string
    except (ValueError, SyntaxError):
        return input_string

def check_brackets(file_path: str, linearized=False):
    """Checks if all brackets in a file are matched or not

    Args:
        file_path (str): _description_
        linearized (bool): if set to "True" the position is given by it's col number instead of line number
    Returns:
        _type_: _description_
    """
    with open(file_path, "r", encoding="utf-8") as _:
        file_content = _.read()

    mismatch_list = []
    stack = []
    mismatches = []
    positions = []
    lines = file_content.split("\n")

    for ind, line in enumerate(lines, start=1):              # ind: line_number
        for ind2, char in enumerate(line, start=1):          # ind2: char_number
            if char == "(":
                stack.append((char, ind, ind2))
                if stack[-1:] == [("(-", ind, ind2)]:
                    break
            elif char == ")":
                if stack and stack[-1][0] == "(":
                    stack.pop()
                else:
                    mismatches.append((char, ind, ind2))

    for bracket, ind, ind2 in stack:
        mismatches.append((bracket, ind, ind2))

    if mismatches:
        for bracket, ind, ind2 in mismatches:
            if linearized:
                position = ind2
            else:
                position = f"line {ind}"
            print(f"Bracket '{bracket}' at position {position}")
            mismatch_list.append(f"Bracket '{bracket}' at position {position}")
        return mismatch_list

    return mismatch_list

