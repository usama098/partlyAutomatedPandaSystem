"""
Section 3 (Support Session Notes) generator.

Reads a tenant JSON file (see ``sample_input.json`` for the expected shape)
and fills in the "Support Session Template.docx" template for every support
session, saving one Word document per session into:

    output/<tenant_name>/Section 3/<Month Year>/<DD-MM-YYYY>.docx

This module intentionally does NOT embed any evidence pictures. The JSON's
``evidence_table`` holds a markdown table (produced by a separate tool); it
is parsed and rendered as a real Word table inside the Support Notes section
(see ``fill_support_notes`` / ``add_markdown_table_to_cell`` for exactly how).

Usage:
    py -3.11 generate_section_three.py <input_json_path> [--output OUTPUT_DIR] [--template TEMPLATE_PATH]
"""

import argparse
import json
import random
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# Header labels used to locate cells/rows inside the template.
LABEL_DATE = "Date"
LABEL_CLIENT_NAME = "Client Name"
LABEL_ADDRESS = "Address"
LABEL_CONTACT_TYPES = ["In Person", "Phone Call", "Text"]
LABEL_SUPPORT_NOTES = "Support Notes"
LABEL_AREAS_COVERED = "Areas covered from support plan"
LABEL_SERVICE_USER_ACTIONS = "Any actions to be completed by service user"
LABEL_SUPPORT_WORKER_ACTIONS = "Any actions to be taken by support worker"

# Text inserted between the support notes and the evidence table, so a
# human reviewer can clearly see where the notes end and the evidence
# table begins.
EVIDENCE_PROMPT_LABEL = "Make image from the following info"

# Probability that the sub-area is also printed on its own line below the
# main area, so generated documents don't all look identically formatted.
SUB_AREA_INCLUSION_PROBABILITY = 1.0

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')

THIS_DIR = Path(__file__).parent
DEFAULT_TEMPLATE_PATH = THIS_DIR.parent / "Templates" / "Support Session Template.docx"
DEFAULT_OUTPUT_DIR = THIS_DIR.parent / "output"

# Section 3 documents always live under a "Section 3" folder within the
# tenant's folder, alongside sibling folders for other sections (e.g.
# "Section 4"), each populated by their own generator.
SECTION_FOLDER_NAME = "Section 3"


def parse_session_date(date_str):
    """Parse a DD:MM:YYYY date string into (day, month, year) ints."""
    day, month, year = (int(part) for part in date_str.split(":"))
    return day, month, year


def format_filename_date(day, month, year):
    return f"{day:02d}-{month:02d}-{year:04d}"


def format_display_date(day, month, year):
    return f"{day:02d}/{month:02d}/{year:04d}"


def sanitize_folder_name(name):
    return INVALID_FILENAME_CHARS.sub("_", name).strip()


def clear_paragraph(paragraph):
    """Remove all runs from a paragraph, leaving it empty but still present."""
    for run in list(paragraph.runs):
        run.text = ""


# Matches inline markdown-style emphasis used in support notes:
#   **<u>text</u>**  or  <u>**text**</u>  -> bold + underline
#   **text**                             -> bold only
#   <u>text</u>                          -> underline only
INLINE_FORMATTING_RE = re.compile(
    r"\*\*<u>(?P<bu1>.*?)</u>\*\*"
    r"|<u>\*\*(?P<bu2>.*?)\*\*</u>"
    r"|\*\*(?P<bold>.*?)\*\*"
    r"|<u>(?P<underline>.*?)</u>",
    re.DOTALL,
)


def add_formatted_runs(paragraph, text):
    """
    Add `text` to `paragraph` as one or more runs, translating the inline
    **bold**/<u>underline</u> markdown-style markup described by
    `INLINE_FORMATTING_RE` into actual bold/underline run formatting.
    Embedded newlines within a run's text are rendered as line breaks by
    python-docx automatically.
    """
    position = 0
    for match in INLINE_FORMATTING_RE.finditer(text):
        if match.start() > position:
            paragraph.add_run(text[position:match.start()])

        if match.group("bu1") is not None or match.group("bu2") is not None:
            content = match.group("bu1") if match.group("bu1") is not None else match.group("bu2")
            bold, underline = True, True
        elif match.group("bold") is not None:
            content, bold, underline = match.group("bold"), True, False
        else:
            content, bold, underline = match.group("underline"), False, True

        run = paragraph.add_run(content)
        run.bold = bold
        run.underline = underline
        position = match.end()

    if position < len(text):
        paragraph.add_run(text[position:])


# Marks the start of the standard end-of-session questions block that
# appears at the end of ``support_notes`` (wellbeing/safety/goal-progress
# questions). Everything from this point onward is moved to print after the
# evidence table instead of staying embedded mid-notes.
QA_SECTION_MARKER_RE = re.compile(
    r"do\s+you\s+have\s+any\s+concerns\s+with\s+your\s+wellbeing", re.IGNORECASE
)


def split_notes_and_qa_section(support_notes):
    """
    Split `support_notes` into (narrative, qa_section), where `qa_section`
    is the trailing block of standard end-of-session questions (starting
    from "Do you have any concerns with your wellbeing...") and `narrative`
    is everything before it. If the marker isn't found, `qa_section` is None
    and the full text is returned as `narrative`.
    """
    match = QA_SECTION_MARKER_RE.search(support_notes)
    if match is None:
        return support_notes, None

    boundary = support_notes.rfind("\n\n", 0, match.start())
    if boundary == -1:
        return "", support_notes.strip()

    narrative = support_notes[:boundary].rstrip()
    qa_section = support_notes[boundary:].strip()
    return narrative, qa_section


def set_cell_text(cell, text, bold=False):
    """Replace a table cell's content with a single run of text."""
    paragraph = cell.paragraphs[0]
    clear_paragraph(paragraph)
    run = paragraph.add_run(text)
    run.bold = bold


def find_label_value_row(document, label):
    """
    Find a 2-column row where the first cell's text matches `label` and
    return the second cell (the value cell to fill in).
    """
    for table in document.tables:
        for row in table.rows:
            if len(row.cells) >= 2 and row.cells[0].text.strip() == label:
                return row.cells[1]
    raise ValueError(f'Could not find a "{label}" row in the template.')


def find_section_row(document, header_label):
    """
    Find the header row whose single cell's text matches `header_label`,
    then return the content row (not just its cell) immediately below it.
    """
    for table in document.tables:
        for row_index, row in enumerate(table.rows):
            if row.cells[0].text.strip() == header_label:
                if row_index + 1 < len(table.rows):
                    return table.rows[row_index + 1]
    raise ValueError(f'Could not find a "{header_label}" section in the template.')


def find_section_content_cell(document, header_label):
    """
    Find the header row whose single cell's text matches `header_label`,
    then return the content cell in the row immediately below it.
    """
    return find_section_row(document, header_label).cells[0]


def set_row_cant_split(row, cant_split=True):
    """
    Set (or clear) the "Allow row to break across pages" option for `row`.
    When `cant_split` is True, Word will keep the whole row together on one
    page instead of splitting its content across a page break.
    """
    tr = row._tr
    tr_pr = tr.find(qn("w:trPr"))
    if tr_pr is None:
        tr_pr = OxmlElement("w:trPr")
        tr.insert(0, tr_pr)

    cant_split_el = tr_pr.find(qn("w:cantSplit"))
    if cant_split:
        if cant_split_el is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
    elif cant_split_el is not None:
        tr_pr.remove(cant_split_el)


def fill_date_client_address(document, tenant_name, session):
    #day, month, year = parse_session_date(session["session_date"])
    date_cell = find_label_value_row(document, LABEL_DATE)
    #set_cell_text(date_cell, format_display_date(day, month, year))

    date_with_time = session.get("date_with_time")
    if date_with_time:
        date_cell.add_paragraph().add_run(date_with_time)

    set_cell_text(find_label_value_row(document, LABEL_CLIENT_NAME), tenant_name)
    set_cell_text(find_label_value_row(document, LABEL_ADDRESS), session["property_address"])


def mark_contact_type(document, contact_type):
    for table in document.tables:
        for row in table.rows:
            cell_texts = [c.text.strip() for c in row.cells]
            if any(label in cell_texts for label in LABEL_CONTACT_TYPES):
                for index, text in enumerate(cell_texts):
                    if text == contact_type and index + 1 < len(row.cells):
                        set_cell_text(row.cells[index + 1], "X", bold=True)
                        return
                raise ValueError(
                    f'Contact type "{contact_type}" not recognised. '
                    f"Expected one of {LABEL_CONTACT_TYPES}."
                )
    raise ValueError("Could not find the Type of Contact table in the template.")


def fill_areas_covered(document, area_covered, sub_area_covered, include_sub_area):
    """
    Write `area_covered` into the "Areas covered from support plan" cell.
    When `include_sub_area` is True, `sub_area_covered` is added on the next
    line within the same cell.
    """
    cell = find_section_content_cell(document, LABEL_AREAS_COVERED)
    first_paragraph = cell.paragraphs[0]
    clear_paragraph(first_paragraph)
    first_paragraph.add_run(area_covered)

    if include_sub_area:
        sub_area_paragraph = cell.add_paragraph()
        sub_area_paragraph.add_run(sub_area_covered)


def parse_markdown_table(markdown_text):
    """
    Parse a GitHub-style markdown table into a list of rows, each a list of
    cell strings. The header-separator row (e.g. ``|---|---|``) is skipped.
    Rows are padded to a common column count using empty strings.
    """
    separator_re = re.compile(r'^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$')

    rows = []
    for line in markdown_text.strip().splitlines():
        line = line.strip()
        if not line or separator_re.match(line):
            continue
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]
        rows.append([cell.strip() for cell in line.split("|")])

    if not rows:
        return rows

    num_cols = max(len(row) for row in rows)
    return [row + [""] * (num_cols - len(row)) for row in rows]


def add_markdown_table_to_cell(cell, markdown_text):
    """
    Parse `markdown_text` as a markdown table and add it as a native Word
    table nested inside `cell`, with the header row shown in bold.
    """
    rows = parse_markdown_table(markdown_text)
    if not rows:
        return

    num_rows = len(rows)
    num_cols = len(rows[0])
    table = cell.add_table(rows=num_rows, cols=num_cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass

    for row_index, row_values in enumerate(rows):
        for col_index, value in enumerate(row_values):
            table_cell = table.cell(row_index, col_index)
            paragraph = table_cell.paragraphs[0]
            run = paragraph.add_run(value)
            run.bold = row_index == 0


def fill_support_notes(document, support_notes, support_note_summary, evidence_table):
    """
    Populate the Support Notes cell with, in order:
      1. The support-notes narrative (everything before the standard
         end-of-session questions block).
      2. A blank line.
      3. The support note summary.
      4. The evidence table, rendered as a native Word table (parsed from
         the `evidence_table` markdown).
      5. A blank line, then the standard end-of-session questions block
         (wellbeing/safety/goal-progress questions), moved here from the
         end of `support_notes` so it prints after the evidence table.
    """
    row = find_section_row(document, LABEL_SUPPORT_NOTES)
    cell = row.cells[0]

    narrative, qa_section = split_notes_and_qa_section(support_notes)

    first_paragraph = cell.paragraphs[0]
    clear_paragraph(first_paragraph)
    add_formatted_runs(first_paragraph, narrative)

    cell.add_paragraph()  # Blank separator line.

    cell.add_paragraph().add_run(support_note_summary)
    cell.add_paragraph()
    add_markdown_table_to_cell(cell, evidence_table)

    if qa_section:
        cell.add_paragraph()  # Blank separator line.
        add_formatted_runs(cell.add_paragraph(), qa_section)

    # Keep the whole Support Notes row together on one page instead of
    # letting Word split its (often lengthy) content across a page break.
    set_row_cant_split(row, cant_split=True)


def fill_bullet_list(document, header_label, items):
    cell = find_section_content_cell(document, header_label)
    paragraphs = cell.paragraphs
    for index, item in enumerate(items):
        if index < len(paragraphs):
            paragraph = paragraphs[index]
            clear_paragraph(paragraph)
        else:
            paragraph = cell.add_paragraph()
        paragraph.add_run(f"\u2022 {item}")


def build_session_document(template_path, tenant_name, session_entry):
    document = Document(template_path)
    session = session_entry["session"]

    fill_date_client_address(document, tenant_name, session)
    mark_contact_type(document, session["contact_type"])
    fill_areas_covered(
        document,
        session["area_covered"],
        session["sub_area_covered"],
        include_sub_area=random.random() < SUB_AREA_INCLUSION_PROBABILITY,
    )
    fill_support_notes(
        document,
        session["support_notes"],
        session["support_note_summary"],
        session_entry["evidence_table"],
    )
    fill_bullet_list(
        document,
        LABEL_SERVICE_USER_ACTIONS,
        session.get("any_actions_to_be_completed_by_service_user", []),
    )
    fill_bullet_list(
        document,
        LABEL_SUPPORT_WORKER_ACTIONS,
        session.get("any_action_to_be_taken_by_support_worker", []),
    )
    return document


def generate_documents(input_json_path, output_dir, template_path):
    input_json_path = Path(input_json_path)
    output_dir = Path(output_dir)
    template_path = Path(template_path)

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tenant_name = data["tenant_name"]
    section_folder = output_dir / sanitize_folder_name(tenant_name) / SECTION_FOLDER_NAME

    created_files = []
    for session_entry in data["sessions"]:
        session = session_entry["session"]
        day, month, year = parse_session_date(session["session_date"])
        month_folder_name = f"{MONTH_NAMES[month - 1]} {year}"
        month_folder = section_folder / month_folder_name
        month_folder.mkdir(parents=True, exist_ok=True)

        document = build_session_document(template_path, tenant_name, session_entry)

        filename = format_filename_date(day, month, year) + ".docx"
        output_path = month_folder / filename
        document.save(output_path)
        created_files.append(output_path)
        print(f"Created: {output_path}")

    return created_files


def main():
    parser = argparse.ArgumentParser(description="Generate Section 3 Support Session Notes Word documents from a JSON file.")
    parser.add_argument("input_json", help="Path to the input JSON file.")
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR}).",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE_PATH),
        help=f"Path to the template .docx (default: {DEFAULT_TEMPLATE_PATH}).",
    )
    args = parser.parse_args()

    try:
        generate_documents(args.input_json, args.output, args.template)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
