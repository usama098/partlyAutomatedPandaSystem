"""
Minimal web front-end for the paperwork automation tool.

Serves a single page with a textarea to paste the input JSON and a button
to generate documents. On submit, both Section 3 and Section 4 generators
are run into an isolated output folder (kept on disk under the system temp
directory, one subfolder per run), and the user is shown a results page
with a single zip download containing every generated Word document,
named after the month/year of the sessions just generated.

Usage:
    py -3.11 -m pip install flask python-docx
    py -3.11 webapp.py
Then open http://127.0.0.1:5000 in a browser.
"""

import json
import tempfile
import uuid
import zipfile
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for

from section_three_populator.daily_case_notes_generator import (
    DEFAULT_TEMPLATE_PATH as DAILY_LOGS_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_daily_logs_documents,
)
from section_three_populator.generate_section_three import (
    DEFAULT_TEMPLATE_PATH as SECTION_THREE_DEFAULT_TEMPLATE_PATH,
    SECTION_FOLDER_NAME as SECTION_THREE_FOLDER_NAME,
    generate_documents as generate_section_three_documents,
    sanitize_folder_name,
)
from section_four_populator.generate_section_four import (
    DEFAULT_SUPPORT_PLAN_DIR as SECTION_FOUR_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_section_four_documents,
)

app = Flask(__name__, template_folder="Templates", static_folder="static")
app.secret_key = "paperwork-automation-dev-key"  # only used to flash form errors

# Every generation run gets its own folder here, named by a random run id, so
# concurrent/successive runs never clash and files stick around long enough
# for the user to click each download link on the results page.
RUNS_DIR = Path(tempfile.gettempdir()) / "panda_paperwork_runs"
RUNS_DIR.mkdir(parents=True, exist_ok=True)

# The upload form has 5 "boxes" (images_1..images_5). Box N's pictures are
# embedded at the end of the Nth Section 3 document generated (in the same
# order as the JSON's "sessions" list) - see `session_images` below.
IMAGE_FIELD_NAMES = [f"images_{i}" for i in range(1, 6)]


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


def _zip_archive_name(relative_paths, tenant_name):
    """
    Build a zip archive base name (no extension) from the Section 3
    month/year folder(s) referenced by `relative_paths` (e.g.
    "August 2026"). If sessions span multiple months, the distinct
    month/year names are joined with " & ". Falls back to `tenant_name`
    when no Section 3 documents were generated.
    """
    month_years = []
    for rel in relative_paths:
        parts = Path(rel).parts
        if len(parts) >= 3 and parts[1] == SECTION_THREE_FOLDER_NAME and parts[2] not in month_years:
            month_years.append(parts[2])

    name = " & ".join(month_years) if month_years else tenant_name
    return sanitize_folder_name(name)


@app.route("/generate", methods=["POST"])
def generate():
    raw_json = request.form.get("input_json", "").strip()
    if not raw_json:
        flash("Please paste some JSON before generating.")
        return redirect(url_for("index"))

    try:
        data = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        flash(f"Invalid JSON: {exc}")
        return redirect(url_for("index"))

    generate_s3 = "section3" in request.form
    generate_s4 = "section4" in request.form
    generate_daily_logs = "daily_logs" in request.form
    if not generate_s3 and not generate_s4 and not generate_daily_logs:
        flash("Select at least one section to generate.")
        return redirect(url_for("index"))

    run_id = uuid.uuid4().hex
    run_dir = RUNS_DIR / run_id
    output_dir = run_dir / "output"
    output_dir.mkdir(parents=True)

    input_json_path = run_dir / "input.json"
    input_json_path.write_text(json.dumps(data), encoding="utf-8")

    # Save each box's uploaded images to disk (skipping empty file inputs),
    # keeping them in box order so box N's images land on the Nth Section 3
    # document.
    session_images = []
    for field_name in IMAGE_FIELD_NAMES:
        uploaded = [f for f in request.files.getlist(field_name) if f and f.filename]
        if not uploaded:
            session_images.append([])
            continue

        box_dir = run_dir / "uploads" / field_name
        box_dir.mkdir(parents=True, exist_ok=True)
        saved_paths = []
        for file in uploaded:
            dest = box_dir / file.filename
            file.save(dest)
            saved_paths.append(dest)
        session_images.append(saved_paths)

    created_files = []
    try:
        if generate_s3:
            created_files += generate_section_three_documents(
                input_json_path, output_dir, SECTION_THREE_DEFAULT_TEMPLATE_PATH,
                session_images=session_images,
            )
        if generate_s4:
            created_files += generate_section_four_documents(
                input_json_path, output_dir, SECTION_FOUR_DEFAULT_TEMPLATE_PATH
            )
        if generate_daily_logs:
            created_files += generate_daily_logs_documents(
                input_json_path, output_dir, DAILY_LOGS_DEFAULT_TEMPLATE_PATH
            )
    except Exception as exc:
        flash(f"Error generating documents: {exc}")
        return redirect(url_for("index"))

    if not created_files:
        flash("No documents were generated from that JSON.")
        return redirect(url_for("index"))

    relative_paths = sorted(Path(f).relative_to(output_dir).as_posix() for f in created_files)

    tenant_name = data.get("tenant_name", "Documents")
    zip_name = _zip_archive_name(relative_paths, tenant_name) + ".zip"
    zip_path = run_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for full_path, relpath in zip(created_files, relative_paths):
            zip_file.write(full_path, arcname=relpath)

    return render_template(
        "results.html",
        run_id=run_id,
        tenant_name=tenant_name,
        files=relative_paths,
        zip_name=zip_name,
    )


@app.route("/download-zip/<run_id>/<path:zip_name>", methods=["GET"])
def download_zip(run_id, zip_name):
    run_dir = (RUNS_DIR / run_id).resolve()
    zip_path = (run_dir / zip_name).resolve()

    # Guard against path traversal outside this run's folder, and only ever
    # serve the zip archive built for this run (not arbitrary files).
    if run_dir not in zip_path.parents or zip_path.suffix != ".zip" or not zip_path.is_file():
        abort(404)

    return send_file(zip_path, as_attachment=True, download_name=zip_path.name)


@app.route("/download/<run_id>/<path:relpath>", methods=["GET"])
def download(run_id, relpath):
    output_dir = (RUNS_DIR / run_id / "output").resolve()
    file_path = (output_dir / relpath).resolve()

    # Guard against path traversal outside this run's output folder.
    if output_dir not in file_path.parents or not file_path.is_file():
        abort(404)

    return send_file(file_path, as_attachment=True, download_name=file_path.name)


if __name__ == "__main__":
    app.run(debug=True)
