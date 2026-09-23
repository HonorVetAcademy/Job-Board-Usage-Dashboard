"""Dispatch table: which parser handles which raw filename pattern."""
from . import doccafe, indeed, linkedin, monster, resume_library, signalhire, vivian

# (substring to match in filename, case-insensitive) -> parser module
FILENAME_MAP = [
    ("doccafe", doccafe),
    ("indeed", indeed),
    ("linkedin", linkedin),
    ("monster", monster),
    ("resume-library", resume_library),
    ("resume_library", resume_library),
    ("signalhire", signalhire),
    ("vivian", vivian),
]


def parser_for_filename(filename):
    lower = filename.lower()
    for token, module in FILENAME_MAP:
        if token in lower:
            return module
    return None
