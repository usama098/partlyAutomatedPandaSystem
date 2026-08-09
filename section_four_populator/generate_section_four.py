"""
Section 4 (Support Plan) generator.

Reads a tenant JSON file (the same schema used by Section 3 - see
``section_three_populator/generate_section_three.py`` and
``Chatgpt_first_input.md`` for the full spec) and, for every session:

1. Validates the session's ``(area_covered, sub_area_covered)`` pair is one
   of the pairs listed in the JSON's ``list_of_support_areas_covered``
   array.
2. Uses ``support_plan_engine.SupportPlanRegistry`` to find the matching
   template and copy it (on first use only) into:

       output/<tenant_name>/Section 4/<area_covered>/<sub_area_covered>.docx

   Every later session covering the same sub-area reuses that same output
   file, so its entries accumulate rather than the file being overwritten.
3. Appends one line to that document's Comments row via
   ``support_plan_engine.append_comment_entry``:

       <session_date><separator><support_note_summary>

   where ``<separator>`` is randomly one of ``"-"``, ``":"``, ``" "`` each
   time, and consecutive entries are separated by 0 or 1 blank line
   (decided randomly on each call).

All template lookup and document-writing logic lives in
``support_plan_engine.py``; this module only wires the JSON up to it.

Usage:
    py -3.11 generate_section_four.py <input_json_path> [--output OUTPUT_DIR] [--templates SUPPORT_PLAN_DIR]
"""

import argparse
import json
import random
import re
import sys
from pathlib import Path
from typing import Dict, Set, Tuple

from docx import Document

from section_four_populator.support_plan_engine import (
    DEFAULT_SUPPORT_PLAN_DIR,
    SupportPlanRegistry,
    append_comment_entry,
)

THIS_DIR = Path(__file__).parent
DEFAULT_OUTPUT_DIR = THIS_DIR.parent / "output"
DEFAULT_TEMPLATE_PATH = DEFAULT_SUPPORT_PLAN_DIR  # kept for parity with run.py's per-section defaults

SECTION_FOLDER_NAME = "Section 4"

# Randomly chosen each time an entry is written, per the spec: date and
# summary are joined by one of these three separators.
COMMENT_SEPARATORS = ["-", ":", " "]

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def sanitize_folder_name(name: str) -> str:
    return INVALID_FILENAME_CHARS.sub("_", name).strip()


def build_allowed_area_pairs(data: dict) -> Set[Tuple[str, str]]:
    """
    Return every (area_covered, sub_area_covered) pair permitted by this
    tenant's `list_of_support_areas_covered` array.
    """
    return {
        (entry["area_covered"], entry["sub_area_covered"])
        for entry in data.get("list_of_support_areas_covered", [])
    }


def format_comment_entry(session_date: str, support_note_summary: str) -> str:
    """Build one "<session_date><separator><support_note_summary>" comment line."""
    return f"{session_date}{random.choice(COMMENT_SEPARATORS)}{support_note_summary}"


def get_or_open_output_document(
    registry: SupportPlanRegistry,
    area_covered: str,
    sub_area_covered: str,
    section_folder: Path,
    opened_documents: Dict[Path, Document],
) -> Tuple[Path, Document]:
    """
    Ensure a per-tenant copy of the (area_covered, sub_area_covered)
    template exists under `section_folder`, and return its output path
    together with an open `Document` for it. The same in-memory `Document`
    is reused (via `opened_documents`) across every session in this run
    that references the same sub-area, so their comment entries accumulate
    correctly before being saved once at the end.
    """
    output_path = (
        section_folder
        / sanitize_folder_name(area_covered)
        / f"{sanitize_folder_name(sub_area_covered)}.docx"
    )

    if output_path in opened_documents:
        return output_path, opened_documents[output_path]

    registry.find(area_covered, sub_area_covered).copy_to(output_path)

    document = Document(output_path)
    opened_documents[output_path] = document
    return output_path, document


def generate_documents(input_json_path, output_dir, support_plan_dir=DEFAULT_SUPPORT_PLAN_DIR):
    input_json_path = Path(input_json_path)
    output_dir = Path(output_dir)

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tenant_name = data["tenant_name"]
    allowed_area_pairs = build_allowed_area_pairs(data)
    registry = SupportPlanRegistry(support_plan_dir)
    section_folder = output_dir / sanitize_folder_name(tenant_name) / SECTION_FOLDER_NAME

    opened_documents: Dict[Path, Document] = {}
    updated_paths = []

    for session_entry in data["sessions"]:
        session = session_entry["session"]
        area_covered = session["area_covered"]
        sub_area_covered = session["sub_area_covered"]

        if (area_covered, sub_area_covered) not in allowed_area_pairs:
            raise ValueError(
                f'Session {session_entry.get("session_number", "?")} references '
                f'area_covered="{area_covered}" / sub_area_covered="{sub_area_covered}", '
                "which is not listed in this tenant's list_of_support_areas_covered."
            )

        output_path, document = get_or_open_output_document(
            registry, area_covered, sub_area_covered, section_folder, opened_documents
        )

        entry_text = format_comment_entry(session["session_date"], session["support_note_summary"])
        append_comment_entry(document, entry_text)

        if output_path not in updated_paths:
            updated_paths.append(output_path)

    for output_path in updated_paths:
        opened_documents[output_path].save(output_path)
        print(f"Updated: {output_path}")

    return updated_paths


def main():
    parser = argparse.ArgumentParser(description="Generate/update Section 4 Support Plan documents from a JSON file.")
    parser.add_argument("input_json", help="Path to the input JSON file.")
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR}).",
    )
    parser.add_argument(
        "--templates",
        default=str(DEFAULT_SUPPORT_PLAN_DIR),
        help=f"Path to the Support Plan template folder (default: {DEFAULT_SUPPORT_PLAN_DIR}).",
    )
    args = parser.parse_args()

    try:
        generate_documents(args.input_json, args.output, args.templates)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
