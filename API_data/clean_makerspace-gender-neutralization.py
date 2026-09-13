"""
Script to clean and gender-neutralize text in makerspace study Excel files
Usage:
  python clean_makerspace-gender-neutralization.py                         # uses defaults below
Run  python clean_makerspace-gender-neutralization.py --help  for all options.
Author: Mariam Sulleiman 23/04/2026
"""


import re
import pandas as pd
from openpyxl import load_workbook

INPUT = "LLMData.xlsx"
OUTPUT = "processed_output.xlsx"

# Replacement rules – ordered from most-specific to least-specific so that
# longer patterns are matched before their substrings can interfere.
# Each tuple: (compiled regex, replacement string)
GENDER_RULES = [
    # Possessive + noun  to  "their <noun>"
    (re.compile(r'\b[Hh]er\s+role\b'), "their role"),
    (re.compile(r'\b[Hh]er\s+free\s+time\b'), "their free time"),
    (re.compile(r'\b[Hh]er\s+making\s+abilities\b'), "their making abilities"),
    (re.compile(r'\b[Hh]er\s+qualifications\b'), "their qualifications"),

    # "she is" / "She is" at start-of-sentence or after punctuation
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+is\b'), "they are"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+was\b'), "they were"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+has\b'), "they have"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+does\b'), "they do"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+serves\b'), "they serve"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+demonstrates\b'), "they demonstrate"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+initiated\b'), "they initiated"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+completed\b'), "they completed"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+mentors\b'), "they mentor"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+adapted\b'), "they adapted"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+completes\b'), "they complete"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+persists\b'), "they persist"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+consistently\b'), "they consistently"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+regularly\b'), "they regularly"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+successfully\b'), "they successfully"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+appears\b'), "they appear"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+navigates\b'), "they navigate"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+drives\b'), "they drive"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+sourced\b'), "they sourced"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+selected\b'), "they selected"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+designed\b'), "they designed"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+loves\b'), "they love"),
    (re.compile(r'(?<![a-zA-Z])[Ss]he\s+is\b'), "they are"),   # catch-all 'she is'
    (re.compile(r'\b[Ss]he\b'), "they"),                         # remaining bare 'she'

    # "her" as object  to  "them"
    (re.compile(r'\bhelping\s+her\b'), "helping them"),
    (re.compile(r'\bguiding\s+her\b'), "guiding them"),
    (re.compile(r'\babout\s+her\b'), "about them"),
    (re.compile(r'\bfor\s+her\b'), "for them"),
    (re.compile(r'\bwith\s+her\b'), "with them"),
    (re.compile(r'\bof\s+her\b'), "of them"),

    # "her" as possessive  to  "their"
    (re.compile(r'\b[Hh]er\b'), "their"),

    # "his" to  "their"
    (re.compile(r'\b[Hh]is\b'), "their"),

    # "he" (subject)
    (re.compile(r'(?<![a-zA-Z])[Hh]e\s+is\b'), "they are"),
    (re.compile(r'(?<![a-zA-Z])[Hh]e\s+was\b'), "they were"),
    (re.compile(r'(?<![a-zA-Z])[Hh]e\s+has\b'), "they have"),
    (re.compile(r'\b[Hh]e\b'), "they"),

    # "him" to  "them"
    (re.compile(r'\b[Hh]im\b'), "them"),

    # Fix double-spaces that may arise
    (re.compile(r'  +'), " "),

    # Grammar fixes after pronoun substitution
    # "they's" to "they're"
    (re.compile(r"\bthey's\b"), "they're"),
    # "they also persists/plans/adapts/etc." – singular verb after plural pronoun
    (re.compile(r'\bthey also persists\b'), "they also persist"),
    (re.compile(r'\bthey plans\b'), "they plan"),
    (re.compile(r'\bthey adapts\b'), "they adapt"),
    (re.compile(r'\bthey troubleshoots\b'), "they troubleshoot"),
    (re.compile(r'\bthey completes\b'), "they complete"),
    (re.compile(r'\bthey mentors\b'), "they mentor"),
    (re.compile(r'\bthey continues\b'), "they continue"),
    (re.compile(r'\bthey engages\b'), "they engage"),
    (re.compile(r'\bthey derives\b'), "they derive"),
    (re.compile(r'\bthey maintains\b'), "they maintain"),
    (re.compile(r'\bthey demonstrates\b'), "they demonstrate"),
    (re.compile(r'\bthey handles\b'), "they handle"),
    (re.compile(r'\bthey serves\b'), "they serve"),
    (re.compile(r'\bthey shows\b'), "they show"),
    (re.compile(r'\bthey works\b'), "they work"),
    (re.compile(r'\bthey navigates\b'), "they navigate"),
    (re.compile(r'\bthey persists\b'), "they persist"),
    (re.compile(r'\bthey adapts\b'), "they adapt"),
    (re.compile(r'\bthey plans\b'), "they plan"),
    (re.compile(r'\bthey suggests\b'), "they suggest"),
    (re.compile(r'\bthey indicates\b'), "they indicate"),
    (re.compile(r'\bthey reflects\b'), "they reflect"),
    (re.compile(r'\bthey appears\b'), "they appear"),
    # "suggesting they's" to "suggesting they're"
    (re.compile(r"\bthey's\b"), "they're"),
]

# Patterns to strip from the START of a cell value
# 1. Leading standalone number rating, e.g. "4\n", "2-3\n"
# 2. Leading "Justification:" / "Justification :" (case-insensitive, optional newline)
STRIP_RATING   = re.compile(r'^\d[\d\-\.]*\s*\n')
STRIP_JUSTIF   = re.compile(r'^Justification\s*:\s*', re.IGNORECASE)


def clean_cell(value):
    if not isinstance(value, str):
        return value

    # 1. Remove leading standalone number rating line
    value = STRIP_RATING.sub('', value)

    # 2. Remove "Justification:" wherever it appears as a label on its own line or at start
    value = re.sub(r'(?im)^Justification\s*:\s*', '', value)

    # 3. Apply gender-neutralisation rules
    for pattern, replacement in GENDER_RULES:
        value = pattern.sub(replacement, value)

    return value.strip()


def process():
    wb = load_workbook(INPUT)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cleaned = clean_cell(cell.value)
                    if cleaned != cell.value:
                        cell.value = cleaned

    wb.save(OUTPUT)
    print(f"Saved cleaned file to: {OUTPUT}")


if __name__ == "__main__":
    process()
