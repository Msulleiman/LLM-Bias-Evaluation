"""
Extract Q1_Obstacles from the All_Data sheet and split each obstacle into its own cell,
for EVERY .xlsx file in a given folder.

Each obstacle is placed in a new row below the original Q1_Obstacles row,
in the same column as the persona it belongs to.
Existing data is preserved.

Handles obstacle lists written in any of these observed styles:
  1) Plain numbered:        "1) Obstacle text\n- Severity: ...\n- Solution: ..."
  2) Markdown numbered:     "1.  **Obstacle:** Text\n    *   **Severity:** ..."
  3) Bold-lead numbered:    "1. **Title**\n   Severity: 2 (Minor)\n   Address: **Yes**"
  4) Plain dashed:          "- Obstacle title\n  - Severity: ...\n  - Solution: ..."

Processed files are saved into an output subfolder with "_extracted" appended
to the original filename.
"""

import os
import re
import glob
import openpyxl
from openpyxl.utils import get_column_letter

# CONFIGURE THESE PATHS
INPUT_DIR = "C:\\Users\\Mariam Sulleiman\\LLM\\processed_output"
OUTPUT_DIR = "C:\\Users\\Mariam Sulleiman\\LLM\\processed_output\\Result"


# Labels that mark a line as obstacle metadata (severity/solution/etc.), not a new obstacle title.
METADATA_PREFIXES = re.compile(
    r'^(severity|needs?\s*addressing\??|address|needs?|solution|suggested\s*solution)\s*:',
    re.IGNORECASE
)


def find_row(sheet, label):
    for row in sheet.iter_rows():
        if row[0].value and str(row[0].value).strip() == label:
            return row[0].row
    return None


def get_persona_columns(sheet):
    """Return dict of {col_index: persona_name} for all persona columns."""
    persona_row = None
    for row in sheet.iter_rows(min_row=1, max_row=3):
        if row[0].value and str(row[0].value).strip() == "Persona":
            persona_row = row[0].row
            break
    if not persona_row:
        raise ValueError("Could not find 'Persona' row")
    cols = {}
    for cell in sheet[persona_row]:
        if cell.column > 1 and cell.value:
            cols[cell.column] = str(cell.value).strip()
    return cols


def clean_title(raw_title):
    """Strip markdown bold markers, leading labels like 'Obstacle:', and trim."""
    t = raw_title.strip()
    # Remove leading bold markers around the whole title: **Title** to Title
    t = re.sub(r'^\*+\s*', '', t)
    t = re.sub(r'\*+\s*$', '', t)
    # Remove a leading "Obstacle:" or "Obstacle X:" label (with optional bold wrapping already stripped)
    t = re.sub(r'^\*{0,2}Obstacle\s*\d*\*{0,2}\s*:\s*', '', t, flags=re.IGNORECASE)
    # Collapse internal bold markers used for emphasis within the title
    t = t.replace('**', '')
    # Collapse whitespace/newlines into single spaces
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def is_metadata_line(line):
    """True if a (stripped, bold-stripped) line is a Severity/Address/Solution/etc. field."""
    candidate = line.strip()
    # Strip leading bullet markers: -, *, bullet dashes, and bold markers
    candidate = re.sub(r'^[-*]\s*', '', candidate)
    candidate = candidate.replace('**', '').strip()
    candidate = re.sub(r'^\*+\s*', '', candidate)
    return bool(METADATA_PREFIXES.match(candidate))


def split_top_level_items(text):
    """
    Split obstacle text into top-level items using two strategies, in order:
      A) Numbered items: lines that start with optional bold + digits + '.' or ')'.
      B) Dashed bullets at column 0 (not indented) when no numbered items exist.
    Returns a list of raw item blocks (each block = title line + its sub-lines).
    """
    lines = text.split('\n')

    # Identify candidate top-level boundaries: numbered markers like "1." "1)" "**1.**"
    numbered_idx = []
    for i, line in enumerate(lines):
        if re.match(r'^\s*\*{0,2}\d+[\.\)]\s*\**\s*\S', line):
            numbered_idx.append(i)

    if numbered_idx:
        blocks = []
        for n, start in enumerate(numbered_idx):
            end = numbered_idx[n + 1] if n + 1 < len(numbered_idx) else len(lines)
            blocks.append('\n'.join(lines[start:end]))
        return blocks, "numbered"

    # Fallback: top-level dash bullets (not indented with leading spaces)
    dash_idx = [i for i, line in enumerate(lines) if re.match(r'^-\s+\S', line)]
    if dash_idx:
        blocks = []
        for n, start in enumerate(dash_idx):
            end = dash_idx[n + 1] if n + 1 < len(dash_idx) else len(lines)
            blocks.append('\n'.join(lines[start:end]))
        return blocks, "dashed"

    return [], "none"


def extract_title_from_block(block, mode):
    """Given a raw block of lines (title + metadata sub-lines), return the cleaned title."""
    lines = block.split('\n')
    first_line = lines[0]

    if mode == "numbered":
        # Strip leading "1.", "1)", optional bold markers
        first_line = re.sub(r'^\s*\*{0,2}\d+[\.\)]\s*', '', first_line)
    elif mode == "dashed":
        first_line = re.sub(r'^-\s+', '', first_line)

    title_parts = [first_line.strip()]

    # If the title line itself is just a bold label with the real title on the same line
    # after a colon (e.g. "**Miter saw height:** The miter saw was..."), keep as-is;
    # clean_title will tidy markdown. But if first line ends up being only metadata-like
    # (rare), or title is empty, pull in continuation lines that aren't metadata.
    for line in lines[1:]:
        if not line.strip():
            continue
        if is_metadata_line(line):
            break  # stop once we hit the first metadata field (Severity/Address/Solution/etc.)
        title_parts.append(line.strip())

    raw_title = ' '.join(title_parts)
    return clean_title(raw_title)


def parse_obstacles(text):
    """Detect format and parse obstacles into a clean list of title strings."""
    if not text or not text.strip():
        return []
    blocks, mode = split_top_level_items(text)
    obstacles = []
    for block in blocks:
        title = extract_title_from_block(block, mode)
        if title:
            obstacles.append(title)
    return obstacles


def process_file(input_path, output_path):
    print(f"\n Processing: {os.path.basename(input_path)}")

    wb = openpyxl.load_workbook(input_path)

    if "All_Data" not in wb.sheetnames:
        print("  Skipped: no 'All_Data' sheet found")
        return False

    ws = wb["All_Data"]

    q1_row = find_row(ws, "Q1_Obstacles")
    if not q1_row:
        print("  Skipped: no 'Q1_Obstacles' row found")
        return False

    persona_cols = get_persona_columns(ws)
    print(f"  Q1_Obstacles at row {q1_row}; Personas: {persona_cols}")

    all_obstacles = {}
    for col_idx, persona in persona_cols.items():
        cell = ws.cell(row=q1_row, column=col_idx)
        raw = cell.value or ""
        obstacles = parse_obstacles(str(raw))
        all_obstacles[col_idx] = (persona, obstacles)
        print(f"    {persona} (col {get_column_letter(col_idx)}): {len(obstacles)} obstacles")
        if len(obstacles) == 0 and str(raw).strip():
            print(f"      WARNING: text present but 0 obstacles parsed -- check format manually")

    max_obstacles = max((len(obs) for _, obs in all_obstacles.values()), default=0)
    if max_obstacles == 0:
        print("  Skipped: no obstacles parsed for any persona")
        return False

    ws.insert_rows(q1_row + 1, max_obstacles)

    for i in range(1, max_obstacles + 1):
        target_row = q1_row + i
        ws.cell(row=target_row, column=1).value = f"Q1_Obstacle_{i}"

    for col_idx, (persona, obstacles) in all_obstacles.items():
        for i, obstacle in enumerate(obstacles):
            target_row = q1_row + 1 + i
            ws.cell(row=target_row, column=col_idx).value = obstacle

    wb.save(output_path)
    print(f"  Saved to: {output_path}")
    return True


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    xlsx_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]

    if not xlsx_files:
        print(f"No .xlsx files found in {INPUT_DIR}")
        return

    print(f"Found {len(xlsx_files)} Excel file(s) in {INPUT_DIR}")

    processed, skipped = 0, 0
    for input_path in xlsx_files:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(OUTPUT_DIR, f"{base_name}_extracted.xlsx")

        try:
            success = process_file(input_path, output_path)
            if success:
                processed += 1
            else:
                skipped += 1
        except Exception as e:
            print(f"  ERROR processing {input_path}: {e}")
            skipped += 1

    print(f"\nDone. Processed: {processed}, Skipped: {skipped}")


if __name__ == "__main__":
    main()
