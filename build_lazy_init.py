# build_lazy_init.py
import os
import ast

BASE_DIR = 'acd'
init_lines = [
    "import importlib\n",
    "__all__ = []\n",
    "def __getattr__(name):\n",
    "    modules = {\n"
]

for root, _, files in os.walk(BASE_DIR):
    for file in files:
        if file.endswith(".py") and file != "__init__.py":
            module_path = os.path.splitext(file)[0]
            full_path = os.path.join(root, file)
            with open(full_path, "r", encoding="utf-8") as f:
                tree = ast.parse(f.read(), filename=file)

                for node in tree.body:
                    if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
                        init_lines.append(f"        '{node.name}': '{module_path}',\n")

init_lines += [
    "    }\n",
    "    if name in modules:\n",
    "        module = importlib.import_module(f'.{modules[name]}', __package__)\n",
    "        return getattr(module, name)\n",
    "    raise AttributeError(f'module {__name__} has no attribute {name}')\n"
]

with open(os.path.join(BASE_DIR, "__init__.py"), "w", encoding="utf-8") as f:
    f.writelines(init_lines)
