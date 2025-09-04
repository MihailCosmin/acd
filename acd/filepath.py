from os.path import sep
from sys import platform


def clean_path(path: str) -> str:
    """Convert path to use OS-specific separators and handle long paths on Windows."""
    path = path.replace("/", sep).replace("\\", sep)
    if platform == "win32" and len(path) > 259 and "\\\\?\\" not in path:
        path = "\\\\?\\" + path
    return path
