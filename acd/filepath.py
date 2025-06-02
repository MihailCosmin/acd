from os.path import sep

def clean_path(path: str) -> str:
    return path.replace("/", sep).replace("\\", sep)
