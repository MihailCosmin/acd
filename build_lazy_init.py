import os
import ast
from collections import defaultdict

BASE_DIR = 'acd'

imports = defaultdict(list)
module_map = {}

for root, _, files in os.walk(BASE_DIR):
    for file in files:
        if file.endswith(".py") and file != "__init__.py":
            module_path = os.path.splitext(file)[0]
            full_path = os.path.join(root, file)
            rel_module = os.path.relpath(full_path, BASE_DIR).replace(os.sep, ".").replace(".py", "")
            with open(full_path, "r", encoding="utf-8") as f:
                tree = ast.parse(f.read(), filename=file)

                for node in tree.body:
                    if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
                        name = node.name
                        imports[rel_module].append(name)
                        module_map[name] = rel_module
                    elif isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name):
                                name = target.id
                                if not name.startswith('_'):
                                    imports[rel_module].append(name)
                                    module_map[name] = rel_module

# -------------------------------
# Generate __init__.py (lazy only)
# -------------------------------
init_lines = [
    "import importlib\n",
    "import sys\n\n",
    "_cache = {}\n\n",
    "__all__ = [\n"
]
init_lines += [f"    '{name}',\n" for name in sorted(module_map.keys())]
init_lines += ["]\n\n"]

init_lines += [
    "def __getattr__(name):\n",
    "    if name in _cache:\n",
    "        return _cache[name]\n",
    "    modules = {\n"
]
init_lines += [f"        '{name}': '{module}',\n" for name, module in sorted(module_map.items())]
init_lines += [
    "    }\n",
    "    if name in modules:\n",
    "        module = importlib.import_module(f'.{modules[name]}', __package__)\n",
    "        value = getattr(module, name)\n",
    "        _cache[name] = value\n",
    "        return value\n",
    "    raise AttributeError(f'module {__name__} has no attribute {name}')\n"
]

with open(os.path.join(BASE_DIR, "__init__.py"), "w", encoding="utf-8") as f:
    f.writelines(init_lines)

# -------------------------------
# Generate __init__.pyi (static stubs for IDEs)
# -------------------------------
pyi_lines = []

for module, names in sorted(imports.items()):
    for name in sorted(names):
        pyi_lines.append(f"from .{module} import {name}\n")

with open(os.path.join(BASE_DIR, "__init__.pyi"), "w", encoding="utf-8") as f:
    f.writelines(pyi_lines)
