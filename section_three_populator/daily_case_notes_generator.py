"""
Daily Case/Contact Notes generator.

Reads the same tenant JSON used by ``generate_section_three.py`` (see
``sample_input.json``) and fills in the "daily_case_note_template.docx"
template with, for every calendar month that contains at least one support
session:

    1. One row per support session in the JSON that falls in that month
       (taken straight from each JSON entry).
    2. One "visited" (in-person, no full session) row and one "called"
       (phone) row for every Mon-Fri calendar week of that month, dated on
       a working day (skipping England bank holidays) that isn't already
       used by a support session that week.

Calendar weeks/months are computed with Python's standard ``calendar``
module (``calendar.monthcalendar``), so every week and month boundary is
exact - no manual date-arithmetic "gaps" between sessions, and no
possibility of a row leaking into a neighbouring month. Only one Word
document is generated per calendar month (named "<Month> daily logs.docx"),
saved alongside the Section 3 documents (same tenant/month folder produced
by ``generate_section_three.py``) into:

    output/<tenant_name>/Section 3/<Month Year>/<Month> daily logs.docx

Usage:
    py -3.11 daily_case_notes_generator.py <input_json_path> [--output OUTPUT_DIR] [--template TEMPLATE_PATH]
"""

import argparse
import calendar
import random
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

import holidays
from docx import Document

from .contact_notes import get_gendered_notes
from .generate_section_three import (
    MONTH_NAMES,
    SECTION_FOLDER_NAME,
    parse_session_date,
    sanitize_folder_name,
)

THIS_DIR = Path(__file__).parent
DEFAULT_TEMPLATE_PATH = THIS_DIR.parent / "Templates" / "daily_case_note_template.docx"
DEFAULT_OUTPUT_DIR = THIS_DIR.parent / "output"

# Known support-area full names mapped to the abbreviations printed on the
# template header (AEW/BH/EA/SS/MPC/MO). Falls back to a generic "first
# letter of each significant word" derivation for any area not listed here.
AREA_ABBREVIATIONS = {
    "achieve economic wellbeing": "AEW",
    "be healthy": "BH",
    "enjoy & achieve": "EA",
    "enjoy and achieve": "EA",
    "stay safe": "SS",
    "make a positive contribution": "MPC",
    "move on": "MO",
}

# Words ignored by the generic abbreviation fallback (articles/conjunctions
# that aren't normally represented in these support-area codes).
ABBREVIATION_STOPWORDS = {"a", "an", "the", "of", "&", "and"}

CALLED_DURATIONS = ["5 mins", "10 mins"]
VISITED_DURATIONS = ["10 mins", "15 mins", "20 mins"]
SUPPORT_SESSION_DURATION = "1 hour"

TENANT_NAME_PLACEHOLDER_RE = re.compile(r"\btenant_name('s)?\b", re.IGNORECASE)

# A first name is every leading run of letters, stopping at the first
# non-letter character (space, hyphen, comma, apostrophe, etc.).
FIRST_NAME_RE = re.compile(r"[A-Za-z]+")

# calendar.monthcalendar() rows are Mon..Sun (index 0 = Monday); only the
# first 5 entries of each week are working days (Mon-Fri).
WORKING_DAY_INDICES = slice(0, 5)


def extract_first_name(tenant_name):
    """Return just `tenant_name`'s first name (its leading run of letters)."""
    match = FIRST_NAME_RE.match(tenant_name.strip())
    return match.group(0) if match else tenant_name


def replace_tenant_name_placeholder(text, tenant_name):
    """Replace every "tenant_name" (or "tenant_name's") placeholder in `text` with `tenant_name`."""
    def repl(match):
        suffix = match.group(1) or ""
        return tenant_name + suffix

    return TENANT_NAME_PLACEHOLDER_RE.sub(repl, text)


def abbreviate_support_area(area_covered):
    """Abbreviate `area_covered` to its support-area code (e.g. "Stay Safe" -> "SS")."""
    known = AREA_ABBREVIATIONS.get(area_covered.strip().lower())
    if known:
        return known

    words = [w for w in re.split(r"\s+", area_covered.strip()) if w.lower() not in ABBREVIATION_STOPWORDS]
    return "".join(word[0].upper() for word in words if word)


def format_date_ddmmyy(d):
    return d.strftime("%d/%m/%y")


def build_uk_holidays(years):
    padded_years = set()
    for year in years:
        padded_years.update((year - 1, year, year + 1))
    return holidays.UK(subdiv="England", years=sorted(padded_years))


def build_support_session_event(session_entry, tenant_name, rng, support_session_contact_note):
    session = session_entry["session"]
    day, month, year = parse_session_date(session["session_date"])
    note = replace_tenant_name_placeholder(rng.choice(support_session_contact_note), tenant_name)
    return {
        "date": date(year, month, day),
        "duration": SUPPORT_SESSION_DURATION,
        "note": note,
        "support_area": abbreviate_support_area(session["area_covered"]),
        "property_address": session["property_address"],
    }


def build_call_and_visit_event_pair(visit_date, call_date, tenant_name, rng, called, visited):
    visited_note = replace_tenant_name_placeholder(rng.choice(visited), tenant_name)
    called_note = replace_tenant_name_placeholder(rng.choice(called), tenant_name)
    return [
        {
            "date": visit_date,
            "duration": rng.choice(VISITED_DURATIONS),
            "note": visited_note,
            "support_area": "WB",
        },
        {
            "date": call_date,
            "duration": rng.choice(CALLED_DURATIONS),
            "note": called_note,
            "support_area": "WB",
        },
    ]


def pick_two_dates_from_pool(pool, rng):
    """Pick two dates (possibly the same, if the pool only has one) from `pool`."""
    if len(pool) >= 2:
        return tuple(rng.sample(pool, 2))
    if len(pool) == 1:
        return pool[0], pool[0]
    return None


def build_month_events(year, month, support_events_this_month, uk_holidays, tenant_name, rng, called, visited):
    """
    Build every event (support session + one call + one visit per Mon-Fri
    calendar week) for a single calendar month, using ``calendar.monthcalendar``
    so weeks/month boundaries are exact and no event can leak into another
    month.
    """
    support_by_date = {e["date"]: e for e in support_events_this_month}
    events = []

    for week in calendar.monthcalendar(year, month):
        working_day_numbers = week[WORKING_DAY_INDICES]
        week_days = [date(year, month, day) for day in working_day_numbers if day != 0]
        if not week_days:
            continue

        sessions_this_week = [support_by_date[d] for d in week_days if d in support_by_date]
        events.extend(sessions_this_week)

        excluded = {e["date"] for e in sessions_this_week}
        candidates = [d for d in week_days if d not in excluded and d not in uk_holidays]
        picked = pick_two_dates_from_pool(candidates, rng)
        if picked:
            visit_date, call_date = picked
            events.extend(build_call_and_visit_event_pair(visit_date, call_date, tenant_name, rng, called, visited))

    return events


def clear_paragraph(paragraph):
    for run in list(paragraph.runs):
        run.text = ""


def fill_header(document, tenant_name, property_address):
    for paragraph in document.paragraphs:
        if "Service User" in paragraph.text and "House/Room Number" in paragraph.text:
            clear_paragraph(paragraph)
            paragraph.add_run(f"Service User: {tenant_name}\tHouse/Room Number: {property_address}")
            return
    raise ValueError('Could not find the "Service User" / "House/Room Number" line in the template.')


def set_cell_text(cell, lines):
    """Replace a table cell's content with one paragraph per entry in `lines`."""
    paragraph = cell.paragraphs[0]
    clear_paragraph(paragraph)
    paragraph.add_run(lines[0])
    for line in lines[1:]:
        cell.add_paragraph().add_run(line)


def fill_contact_notes_table(document, events):
    """
    Fill in the contact-notes table, reusing the template's existing blank
    rows (below the header) in place before adding any new ones, and
    removing one existing blank row for every row that's filled or added -
    so the table never ends up with unused blank padding rows left over
    from the template.
    """
    table = document.tables[0]
    existing_blank_rows = list(table.rows[1:])
    next_existing_index = 0

    def next_row():
        nonlocal next_existing_index
        if next_existing_index < len(existing_blank_rows):
            row = existing_blank_rows[next_existing_index]
            next_existing_index += 1
            return row
        return table.add_row()

    for index, event in enumerate(sorted(events, key=lambda e: e["date"])):
        if index > 0:
            next_row()  # Blank spacer row between entries.

        row = next_row()
        set_cell_text(row.cells[0], [format_date_ddmmyy(event["date"]), event["duration"]])
        set_cell_text(row.cells[1], [event["note"]])
        set_cell_text(row.cells[2], [event["support_area"]])

    # Remove any leftover, never-used blank rows from the template.
    for row in existing_blank_rows[next_existing_index:]:
        row._tr.getparent().remove(row._tr)


def build_month_document(template_path, tenant_name, property_address, events):
    document = Document(template_path)
    fill_header(document, tenant_name, property_address)
    fill_contact_notes_table(document, events)
    return document


def generate_documents(input_json_path, output_dir, template_path=DEFAULT_TEMPLATE_PATH, seed=None):
    """
    Generate one "<Month> daily logs.docx" document per calendar month that
    contains at least one support session, each containing every Mon-Fri
    calendar week of that month with a support session (if any that week)
    plus one call and one visit.
    """
    import json

    input_json_path = Path(input_json_path)
    output_dir = Path(output_dir)
    template_path = Path(template_path)
    rng = random.Random(seed)

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tenant_name = data["tenant_name"]
    first_name = extract_first_name(tenant_name)
    called, visited, support_session_contact_note = get_gendered_notes(data.get("gender"))
    sessions = sorted(
        data["sessions"],
        key=lambda entry: parse_session_date(entry["session"]["session_date"])[::-1],
    )

    years = {parse_session_date(entry["session"]["session_date"])[2] for entry in sessions}
    uk_holidays = build_uk_holidays(years)

    support_events = [
        build_support_session_event(entry, first_name, rng, support_session_contact_note) for entry in sessions
    ]

    support_events_by_month = defaultdict(list)
    for event in support_events:
        key = (event["date"].year, event["date"].month)
        support_events_by_month[key].append(event)

    section_folder = output_dir / sanitize_folder_name(tenant_name) / SECTION_FOLDER_NAME

    created_files = []
    for (year, month), month_support_events in sorted(support_events_by_month.items()):
        events = build_month_events(year, month, month_support_events, uk_holidays, first_name, rng, called, visited)

        month_name = MONTH_NAMES[month - 1]
        month_folder = section_folder / f"{month_name} {year}"
        month_folder.mkdir(parents=True, exist_ok=True)

        property_address = month_support_events[0]["property_address"]
        document = build_month_document(template_path, tenant_name, property_address, events)

        output_path = month_folder / f"{month_name} daily logs.docx"
        document.save(output_path)
        created_files.append(output_path)
        print(f"Created: {output_path}")

    return created_files


def main():
    parser = argparse.ArgumentParser(description="Generate Daily Case/Contact Notes Word documents from a JSON file.")
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
    parser.add_argument("--seed", type=int, default=None, help="Optional random seed for reproducible output.")
    args = parser.parse_args()

    try:
        generate_documents(args.input_json, args.output, args.template, seed=args.seed)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
