"""
Daily Case/Contact Notes generator.

Reads the same tenant JSON used by ``generate_section_three.py`` (see
``sample_input.json``) and fills in the "daily_case_note_template.docx"
template with:

    1. One row per support session in the JSON (taken straight from each
       JSON entry).
    2. One "visited" (in-person, no full session) row and one "called"
       (phone) row for every gap between two consecutive support
       sessions, dated somewhere in that gap.

The visited/called dates are randomly chosen from the working days (Mon-Fri,
skipping England bank holidays) strictly between each pair of consecutive
session dates. Documents are split by calendar month based on each row's own
actual date (so a gap spanning a month boundary can contribute rows to both
months' documents). Only one Word document is generated per calendar month
(named "<Month> daily logs.docx"), saved alongside the Section 3 documents
(same tenant/month folder produced by ``generate_section_three.py``) into:

    output/<tenant_name>/Section 3/<Month Year>/<Month> daily logs.docx

Usage:
    py -3.11 daily_case_notes_generator.py <input_json_path> [--output OUTPUT_DIR] [--template TEMPLATE_PATH]
"""

import argparse
import random
import re
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import holidays
from docx import Document

from .contact_notes import called, support_session_contact_note, visited
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


def daterange(start, end):
    """Yield every `date` from `start` to `end`, inclusive."""
    for offset in range((end - start).days + 1):
        yield start + timedelta(days=offset)


def working_days_strictly_between(start, end, uk_holidays):
    """Weekdays (Mon-Fri), skipping England bank holidays, strictly between `start` and `end`."""
    if end - start < timedelta(days=2):
        return []
    return [
        d for d in daterange(start + timedelta(days=1), end - timedelta(days=1))
        if d.weekday() < 5 and d not in uk_holidays
    ]


def week_bounds(d):
    """Return (monday, friday) of the Mon-Fri working week containing date `d`."""
    monday = d - timedelta(days=d.weekday())
    friday = monday + timedelta(days=4)
    return monday, friday


def month_start(d):
    return date(d.year, d.month, 1)


def month_end(d):
    if d.month == 12:
        return date(d.year, 12, 31)
    return date(d.year, d.month + 1, 1) - timedelta(days=1)


def split_into_weeks(start, end):
    """Yield (chunk_start, chunk_end) pairs splitting [start, end] (inclusive) into Mon-Fri week chunks."""
    if start > end:
        return
    current = start
    while current <= end:
        monday, friday = week_bounds(current)
        yield current, min(friday, end)
        current = monday + timedelta(days=7)


def pick_two_dates_from_pool(pool, rng):
    """Pick two dates (possibly the same, if the pool only has one) from `pool`."""
    if len(pool) >= 2:
        return tuple(rng.sample(pool, 2))
    if len(pool) == 1:
        return pool[0], pool[0]
    return None


def build_call_and_visit_event_pair(visit_date, call_date, tenant_name, rng):
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


def build_period_call_visit_events(start, end, tenant_name, uk_holidays, rng):
    """
    Build one "visited" and one "called" event for every Mon-Fri week
    (chunked via `split_into_weeks`) that has at least one working,
    non-bank-holiday day somewhere in [start, end] (inclusive). Used to
    cover leftover working days before the first support session's month,
    or after the last support session's month, so those trailing/leading
    weeks still get a call and a visit even though there's no session
    (or session-to-session gap) to anchor them to.
    """
    events = []
    for chunk_start, chunk_end in split_into_weeks(start, end):
        pool = [d for d in daterange(chunk_start, chunk_end) if d.weekday() < 5 and d not in uk_holidays]
        picked = pick_two_dates_from_pool(pool, rng)
        if not picked:
            continue
        visit_date, call_date = picked
        events.extend(build_call_and_visit_event_pair(visit_date, call_date, tenant_name, rng))
    return events


def pick_gap_call_and_visit_dates(prev_date, next_date, uk_holidays, rng):
    """
    Pick two dates (one for "visited", one for "called") somewhere in the
    working days (Mon-Fri, skipping England bank holidays) strictly between
    `prev_date` and `next_date`. Falls back to reusing/expanding the pool if
    the gap is too short of valid days to give two distinct dates.
    """
    candidates = working_days_strictly_between(prev_date, next_date, uk_holidays)
    if len(candidates) >= 2:
        return tuple(rng.sample(candidates, 2))
    if len(candidates) == 1:
        return candidates[0], candidates[0]

    # Gap too short (e.g. sessions on consecutive days) - fall back to any
    # weekday, non-holiday day in the inclusive range, excluding the
    # session dates themselves if possible.
    inclusive_pool = [
        d for d in daterange(prev_date, next_date)
        if d.weekday() < 5 and d not in uk_holidays and d not in (prev_date, next_date)
    ]
    if len(inclusive_pool) >= 2:
        return tuple(rng.sample(inclusive_pool, 2))
    if len(inclusive_pool) == 1:
        return inclusive_pool[0], inclusive_pool[0]

    fallback_pool = [
        d for d in daterange(prev_date, next_date) if d.weekday() < 5 and d not in uk_holidays
    ] or [prev_date, next_date]
    return rng.choice(fallback_pool), rng.choice(fallback_pool)


def build_uk_holidays(years):
    padded_years = set()
    for year in years:
        padded_years.update((year - 1, year, year + 1))
    return holidays.UK(subdiv="England", years=sorted(padded_years))


def build_support_session_event(session_entry, tenant_name, rng):
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


def build_gap_events(prev_date, next_date, tenant_name, uk_holidays, rng):
    """
    Build the one "visited" and one "called" event dated somewhere in the
    working days strictly between `prev_date` and `next_date`.
    """
    visit_date, call_date = pick_gap_call_and_visit_dates(prev_date, next_date, uk_holidays, rng)
    return build_call_and_visit_event_pair(visit_date, call_date, tenant_name, rng)



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
    Generate one "<Month> daily logs.docx" document per calendar month
    touched by any support/visited/called row, each containing every row
    whose own date falls in that month.
    """
    import json

    input_json_path = Path(input_json_path)
    output_dir = Path(output_dir)
    template_path = Path(template_path)
    rng = random.Random(seed)

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tenant_name = data["tenant_name"]
    sessions = sorted(
        data["sessions"],
        key=lambda entry: parse_session_date(entry["session"]["session_date"])[::-1],
    )

    years = {parse_session_date(entry["session"]["session_date"])[2] for entry in sessions}
    uk_holidays = build_uk_holidays(years)

    support_events = [build_support_session_event(entry, tenant_name, rng) for entry in sessions]

    all_events = list(support_events)
    for prev_event, next_event in zip(support_events, support_events[1:]):
        all_events.extend(build_gap_events(prev_event["date"], next_event["date"], tenant_name, uk_holidays, rng))

    if support_events:
        first_date = support_events[0]["date"]
        last_date = support_events[-1]["date"]
        # Cover leftover working days before the first session's month (from
        # the 1st of that month up to, but not including, the session date)
        # and after the last session's month (from the day after the session
        # date through to the end of that month), chunked per Mon-Fri week,
        # so trailing/leading weeks with no session still get a call+visit.
        all_events.extend(
            build_period_call_visit_events(
                month_start(first_date), first_date - timedelta(days=1), tenant_name, uk_holidays, rng
            )
        )
        all_events.extend(
            build_period_call_visit_events(
                last_date + timedelta(days=1), month_end(last_date), tenant_name, uk_holidays, rng
            )
        )

    events_by_month = defaultdict(list)
    for event in all_events:
        key = (event["date"].year, event["date"].month)
        events_by_month[key].append(event)

    # Pick the property address to print in each month's header: the
    # address of whichever support session in that month is used, falling
    # back to the closest support session by date for months that only
    # contain visited/called rows (e.g. a gap spanning a month boundary).
    address_by_month = {}
    for (year, month) in events_by_month:
        month_support_events = [e for e in support_events if (e["date"].year, e["date"].month) == (year, month)]
        if month_support_events:
            address_by_month[(year, month)] = month_support_events[0]["property_address"]
        else:
            reference_date = date(year, month, 1)
            closest = min(support_events, key=lambda e: abs((e["date"] - reference_date).days))
            address_by_month[(year, month)] = closest["property_address"]

    section_folder = output_dir / sanitize_folder_name(tenant_name) / SECTION_FOLDER_NAME

    created_files = []
    for (year, month), events in sorted(events_by_month.items()):
        month_name = MONTH_NAMES[month - 1]
        month_folder = section_folder / f"{month_name} {year}"
        month_folder.mkdir(parents=True, exist_ok=True)

        document = build_month_document(template_path, tenant_name, address_by_month[(year, month)], events)

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

