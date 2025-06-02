import re

# read pyproject.toml
with open('pyproject.toml', 'r', encoding='utf-8') as f:
    pyproject_contents = f.read()

# find version number pyproject
version_match_pyproject = re.search(r'version ?= ?"(\d.\d.\d.\d)"', pyproject_contents)
old_version_pyproject = version_match_pyproject.group(1)

new_version_pyproject = str(int(old_version_pyproject.replace('.', '')) + 1).zfill(4)
new_version_pyproject = '.'.join(
    [
        new_version_pyproject[:1],
        new_version_pyproject[1:2],
        new_version_pyproject[2:3],
        new_version_pyproject[3:]
    ]
)

# write back to setup.py
with open('pyproject.toml', 'w', encoding='utf-8') as f:
    f.write(updated_pyproject_contents)
