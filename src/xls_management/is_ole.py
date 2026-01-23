import olefile
from pathlib import Path

def is_ole(file_path: Path|str) -> str:

    if olefile.isOleFile(file_path):
        return f"{file_path} is ole"
    return f"{file_path} is ole"
    