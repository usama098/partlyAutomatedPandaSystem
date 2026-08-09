"""
Root entry point for the paperwork automation app.

Dispatches to the generator for a given paperwork section. Section 3
(Support Session Notes) and Section 4 (Support Plan) are implemented; new
sections can be added by writing a populator package (e.g.
`section_five_populator/`) with its own
`generate_documents(input_json_path, output_dir, template_path)` function and
registering it in `SECTION_GENERATORS` below.

Usage:
    py -3.11 run.py <input_json_path> [--section 3] [--output OUTPUT_DIR] [--template TEMPLATE_PATH]
"""

import argparse
import sys

from section_three_populator.generate_section_three import (
    DEFAULT_OUTPUT_DIR as SECTION_THREE_DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE_PATH as SECTION_THREE_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_section_three_documents,
)
from section_four_populator.generate_section_four import (
    DEFAULT_OUTPUT_DIR as SECTION_FOUR_DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE_PATH as SECTION_FOUR_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_section_four_documents,
)

# Maps a section identifier to its (generate_documents, default_output_dir,
# default_template_path). Add an entry here for each new section populator.
SECTION_GENERATORS = {
    "3": (
        generate_section_three_documents,
        SECTION_THREE_DEFAULT_OUTPUT_DIR,
        SECTION_THREE_DEFAULT_TEMPLATE_PATH,
    ),
    "4": (
        generate_section_four_documents,
        SECTION_FOUR_DEFAULT_OUTPUT_DIR,
        SECTION_FOUR_DEFAULT_TEMPLATE_PATH,
    ),
}


def main():
    parser = argparse.ArgumentParser(description="Generate paperwork documents from a JSON file.")
    parser.add_argument("input_json", help="Path to the input JSON file.")
    parser.add_argument(
        "--section",
        default="3",
        choices=sorted(SECTION_GENERATORS.keys()),
        help="Which paperwork section to generate: 3 (Support Session Notes) or 4 (Support Plan). Default: 3.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output directory (defaults to the selected section's own output folder).",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Path to the template .docx (defaults to the selected section's own template).",
    )
    args = parser.parse_args()

    generate_documents, default_output_dir, default_template_path = SECTION_GENERATORS[args.section]
    output_dir = args.output or default_output_dir
    template_path = args.template or default_template_path

    try:
        generate_documents(args.input_json, output_dir, template_path)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
