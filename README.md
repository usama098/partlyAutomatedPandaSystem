# Supported Housing Paperwork Automation

Generates Word documents for supported-housing paperwork from a JSON file.
Each paperwork section (e.g. Section 3 - Support Session Notes) has its own
"populator" package with its own template and generator, and is invoked
through the root `run.py` dispatcher.

```
partly_automated_work/
  run.py                              <- single entry point for all sections
  output/                             <- generated documents (shared across sections)
  Templates/
    Support Session Template.docx     <- Section 3 template
    Section_four/
      Support Plan/                   <- one .docx per (area, sub-area) - see below
  section_three_populator/
    generate_section_three.py         <- Section 3 generator logic
    sample_input.json                 <- example input for Section 3
  section_four_populator/
    support_plan_engine.py            <- Section 4 template lookup + write engine (see below)
    generate_section_four.py          <- Section 4 generator logic
  Chatgpt_first_input.md              <- spec given to ChatGPT to produce the input JSON
```

## Setup (one-time)

```powershell
py -3.11 -m pip install python-docx
```

## Usage

```powershell
py -3.11 run.py <path-to-input.json> [--section 3] [--output OUTPUT_DIR] [--template TEMPLATE_PATH]
```

- `--section` selects which populator to run: `3` (Support Session Notes,
  the default) or `4` (Support Plan).
- `--output` defaults to the shared, project-root `output/` folder.
- `--template` defaults to that section's own template
  (e.g. `Templates\Support Session Template.docx` for Section 3, or
  `Templates\Section_four\Support Plan` for Section 4).

Both sections read the same input JSON file (see the schema below), so a
single JSON file drives both:

```powershell
py -3.11 run.py section_three_populator\sample_input.json --section 3
py -3.11 run.py section_three_populator\sample_input.json --section 4
```

## Input JSON format

Both Section 3 and Section 4 read the same input JSON:

```json
{
  "tenant_name": "string",
  "gender": "male | female (optional, defaults to \"male\")",
  "move_in_date": "DD:MM:YYYY",
  "move_out_date": "DD:MM:YYYY",
  "list_of_support_areas_covered": [
    {
      "area_covered": "Main support area",
      "sub_area_covered": "Exact sub-area discussed in the support notes"
    }
  ],
  "sessions": [
    {
      "session_number": 1,
      "session": {
        "session_date": "DD:MM:YYYY",
        "date_with_time": "Start Date/Time -- End Date/Time\n\nDD/MM/YYYY, H:MM AM/PM -- H:MM AM/PM",
        "property_address": "string",
        "contact_type": "In Person",
        "area_covered": "Main support area",
        "sub_area_covered": "Exact sub-area discussed in the support notes",
        "support_notes": "Detailed one-hour first-person support notes",
        "support_note_summary": "Approximately 100-word summary matching the support notes",
        "any_actions_to_be_completed_by_service_user": [
          "Action 1",
          "Action 2"
        ],
        "any_action_to_be_taken_by_support_worker": [
          "Action 1",
          "Action 2"
        ]
      },
      "evidence_table": "markdown table which is produced by chatgpt"
    }
  ]
}
```

See `section_three_populator/sample_input.json` for a full working example,
and `Chatgpt_first_input.md` for the full spec (including the fixed list of
valid `area_covered` / `sub_area_covered` values) used to prompt ChatGPT
into producing this JSON from a tenant's Initial Assessment.

Key points:

- `tenant_name` is used as the "Client Name" on every session document and as
  the top-level output folder name.
- `gender` (`"male"` or `"female"`, case-insensitive, optional) selects which
  pronoun-matched phrase bank the Daily Case/Contact Notes generator picks
  its "called"/"visited"/support-session note text from (see
  `section_three_populator/contact_notes.py`). Defaults to `"male"` when
  omitted.
- `move_in_date` / `move_out_date` describe the tenancy period covered by the
  whole set of sessions (`DD:MM:YYYY`). They aren't written into the
  documents themselves; they're only used to generate the session dates.
- `list_of_support_areas_covered` lists every `(area_covered, sub_area_covered)`
  pair this tenant's plan actually covers. Section 4 validates every
  session's `area_covered` / `sub_area_covered` against this list and stops
  with an error if a session references a pair that isn't listed.
- `session_date` must be `DD:MM:YYYY` (e.g. `"06:08:2026"`).
- `date_with_time` is optional. When present, it's printed as an extra line
  directly beneath the formatted date in the same "Date" field (e.g.
  `"Start Date/Time -- End Date/Time\n\n14/09/2026, 2:00 PM -- 3:00 PM"`).
- `contact_type` must be one of `"In Person"`, `"Phone Call"`, `"Text"`.
- `area_covered` / `sub_area_covered` are written into Section 3's "Areas
  covered from support plan" section. `area_covered` always appears;
  `sub_area_covered` is added on the line below it only some of the time
  (randomised per session), so generated documents don't all look
  identically formatted. Section 4 uses the same two fields to pick which
  support-plan template a session's comment belongs on.
- In the Support Notes section, after `support_notes` the generator adds a
  bold, yellow-highlighted line reading "Make image from the following
  info", followed by `support_note_summary` and then a native Word table
  parsed from `evidence_table`'s markdown. No image is embedded or
  generated by this tool — `evidence_table` is expected to contain a
  markdown table (produced by a separate step), which is parsed and
  rendered as a real table in the document. Section 4 also uses
  `support_note_summary` (see below) but ignores `evidence_table`.

## Output structure

Every tenant gets one folder (their name) directly under the shared `output/`
folder, containing a subfolder per paperwork section:

```
output/
  <tenant_name>/
    Section 3/
      <Month Year>/          e.g. "August 2026"
        <DD-MM-YYYY>.docx    e.g. "06-08-2026.docx"
    Section 4/
      <area_covered>/
        <sub_area_covered>.docx   e.g. "Enjoy and Achieve/Move on.docx"
```

Sessions from the same calendar month in different years get separate
Section 3 folders (e.g. "August 2025" vs "August 2026"). Section 4 instead
creates one document per `(area_covered, sub_area_covered)` pair actually
referenced by a session, copied from the matching template the first time
it's needed; every later session for that same pair reuses and appends to
that same file rather than overwriting it.

## Section 4 - Support Plan generator

`Templates/Section_four/Support Plan/` holds one Word document per support
sub-area, filed as `<Area covered>/<Sub-area covered>.docx` (25 files across
the 5 areas listed in `Chatgpt_first_input.md`). Every file shares the same
fundamental layout: an area heading, a sub-area heading, and one table with
columns `Objective Steps | Date Completed | Client Signature | Support
Worker Signature`, a variable number of "Step ...:" rows (a few files split
a step into lettered sub-steps, e.g. "Step 1a:"/"Step 1b:"), and a trailing
"Comments:" row.

`section_four_populator/support_plan_engine.py` is the reusable engine
behind this, independent of any specific JSON schema:

- `SupportPlanRegistry` indexes every template file and looks one up by
  exact `(area_covered, sub_area_covered)` string match — via
  `.find(area, sub_area)` or `.match(obj)`, where `obj` can be a dict (e.g.
  one entry from a JSON `list_of_support_areas_covered` array) or any object
  exposing those two fields as attributes. Unmatched values raise a
  `ValueError` listing the closest known values.
- `SupportAreaTemplate.open()` loads the matched file as an editable
  `python-docx` `Document`; `SupportAreaTemplate.copy_to(path)` copies the
  raw template file to an output path (skipping the copy if a file is
  already there, so repeat calls never clobber an in-progress output file).
- `list_steps(document)`, `complete_step(document, step_identifier, ...)`,
  and `set_comments(document, ...)` read/write the Objective Steps table
  generically, regardless of how many steps a particular sub-area has.
- `append_comment_entry(document, entry_text)` appends one line to the
  Comments row without disturbing what's already there — used to build up
  a running log of entries (one per session) in a single cell.

`section_four_populator/generate_section_four.py` is a thin layer on top of
that engine: for each session in the input JSON, it resolves the session's
`(area_covered, sub_area_covered)` template, copies it into the output
location on first use, and appends one comment entry per session, formatted
as:

```
<session_date><separator><support_note_summary>
```

where `<separator>` is randomly one of `"-"`, `":"`, `" "` on each entry,
and consecutive entries are separated by either 0 or 1 blank line (also
decided randomly). Sessions referencing an `(area_covered, sub_area_covered)`
pair not present in `list_of_support_areas_covered` raise an error and stop
generation.

## Adding a new section

1. Create a new `section_<n>_populator/` package with its own
   `generate_documents(input_json_path, output_dir, template_path)` function,
   `DEFAULT_OUTPUT_DIR`, and `DEFAULT_TEMPLATE_PATH` (following
   `section_three_populator/generate_section_three.py` as a reference).
   It should write into `output/<tenant_name>/Section <n>/...`, so every
   tenant folder ends up with one subfolder per section.
2. Register it in `SECTION_GENERATORS` in `run.py`.
