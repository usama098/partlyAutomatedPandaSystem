"""
Minimal web front-end for the paperwork automation tool.

Serves a single page with a textarea to paste the input JSON and a button
to generate documents. On submit, both Section 3 and Section 4 generators
are run into an isolated temporary output folder, and the results are
zipped up and returned to the browser as a download.

Usage:
    py -3.11 -m pip install flask python-docx
    py -3.11 webapp.py
Then open http://127.0.0.1:5000 in a browser.
"""

import io
import json
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

from section_three_populator.generate_section_three import (
    DEFAULT_TEMPLATE_PATH as SECTION_THREE_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_section_three_documents,
)
from section_four_populator.generate_section_four import (
    DEFAULT_SUPPORT_PLAN_DIR as SECTION_FOUR_DEFAULT_TEMPLATE_PATH,
    generate_documents as generate_section_four_documents,
)

app = Flask(__name__)
app.secret_key = "paperwork-automation-dev-key"  # only used to flash form errors


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


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
    if not generate_s3 and not generate_s4:
        flash("Select at least one section to generate.")
        return redirect(url_for("index"))

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir = Path(tmp_dir)
        input_json_path = tmp_dir / "input.json"
        input_json_path.write_text(json.dumps(data), encoding="utf-8")

        output_dir = tmp_dir / "output"
        output_dir.mkdir()

        created_files = []
        try:
            if generate_s3:
                created_files += generate_section_three_documents(
                    input_json_path, output_dir, SECTION_THREE_DEFAULT_TEMPLATE_PATH
                )
            if generate_s4:
                created_files += generate_section_four_documents(
                    input_json_path, output_dir, SECTION_FOUR_DEFAULT_TEMPLATE_PATH
                )
        except Exception as exc:
            flash(f"Error generating documents: {exc}")
            return redirect(url_for("index"))

        if not created_files:
            flash("No documents were generated from that JSON.")
            return redirect(url_for("index"))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in created_files:
                file_path = Path(file_path)
                zip_file.write(file_path, arcname=str(file_path.relative_to(output_dir)))
        zip_buffer.seek(0)

    tenant_name = data.get("tenant_name", "documents")
    download_name = f"{tenant_name}.zip".replace("/", "_").replace("\\", "_")
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=download_name,
    )


if __name__ == "__main__":
    app.run(debug=True)
