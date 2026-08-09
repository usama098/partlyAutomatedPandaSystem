"""
Section 3 (Support Session Notes) generator.

Reads a tenant JSON file (see ``sample_input.json`` for the expected shape)
and fills in the "Support Session Template.docx" template for every support
session, saving one Word document per session into:

    output/<tenant_name>/Section 3/<Month Year>/<DD-MM-YYYY>.docx

This module intentionally does NOT embed any evidence pictures. The JSON's
``evidence_picture_description`` is a text description meant for a separate
image-generation tool; here it is only copied into the Support Notes section
as reference text (see ``fill_support_notes`` for exactly how).

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

# Text inserted between the support notes and the evidence-picture
# description, so a human reviewer can clearly see where the notes end and
# the picture-generation prompt begins.
EVIDENCE_PROMPT_LABEL = "Make image from the following info"

# Probability that the sub-area is also printed on its own line below the
# main area, so generated documents don't all look identically formatted.
SUB_AREA_INCLUSION_PROBABILITY = 0.5

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


def find_section_content_cell(document, header_label):
    """
    Find the header row whose single cell's text matches `header_label`,
    then return the content cell in the row immediately below it.
    """
    for table in document.tables:
        for row_index, row in enumerate(table.rows):
            if row.cells[0].text.strip() == header_label:
                if row_index + 1 < len(table.rows):
                    return table.rows[row_index + 1].cells[0]
    raise ValueError(f'Could not find a "{header_label}" section in the template.')


def fill_date_client_address(document, tenant_name, session):
    day, month, year = parse_session_date(session["session_date"])
    set_cell_text(find_label_value_row(document, LABEL_DATE), format_display_date(day, month, year))
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


def fill_support_notes(document, support_notes, support_note_summary, evidence_picture_description):
    """
    Populate the Support Notes cell with, in order:
      1. The full support notes.
      2. A blank line.
      3. A bold, yellow-highlighted "Make image from the following info" label.
      4. The support note summary.
      5. The evidence picture description.
    """
    cell = find_section_content_cell(document, LABEL_SUPPORT_NOTES)

    first_paragraph = cell.paragraphs[0]
    clear_paragraph(first_paragraph)
    first_paragraph.add_run(support_notes)

    cell.add_paragraph()  # Blank separator line.

    prompt_paragraph = cell.add_paragraph()
    prompt_run = prompt_paragraph.add_run(EVIDENCE_PROMPT_LABEL)
    prompt_run.bold = True
    prompt_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    cell.add_paragraph().add_run(support_note_summary)
    cell.add_paragraph().add_run(evidence_picture_description)


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
        session_entry["evidence_picture_description"],
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
