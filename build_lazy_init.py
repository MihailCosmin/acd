# build_lazy_init.py
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

# Start writing __init__.py
lines = []

# 1. Static stub imports
for module, names in imports.items():
    for name in names:
        lines.append(f"from .{module} import {name}\n")

# 2. Add dynamic __getattr__ fallback
lines += [
    "\n\nimport importlib\n",
    "__all__ = [\n"
]
lines += [f"    '{name}',\n" for name in sorted(module_map.keys())]
lines += ["]\n\n"]

lines += [
    "def __getattr__(name):\n",
    "    modules = {\n"
]
lines += [f"        '{name}': '{module}',\n" for name, module in sorted(module_map.items())]
lines += [
    "    }\n",
    "    if name in modules:\n",
    "        module = importlib.import_module(f'.{modules[name]}', __package__)\n",
    "        return getattr(module, name)\n",
    "    raise AttributeError(f'module {__name__} has no attribute {name}')\n"
]

# Write to __init__.py
with open(os.path.join(BASE_DIR, "__init__.py"), "w", encoding="utf-8") as f:
    f.writelines(lines)
