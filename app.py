#!/usr/bin/env python3
"""
SIA 451 / CRBX -> Excel Web-App
"""

import io
from flask import Flask, request, send_file, render_template, flash, redirect, url_for
from werkzeug.middleware.proxy_fix import ProxyFix

from parse_sia451 import parse_sia451
from parse_sia451_excel import build_workbook

app = Flask(__name__)
app.secret_key = "sia451-parser-secret"
app.config["APPLICATION_ROOT"] = "/ausmass"
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1, x_prefix=1)

ALLOWED_EXTENSIONS = {".01s", ".crbx"}


def _allowed(filename: str) -> bool:
    return any(filename.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/parse", methods=["POST"])
def parse():
    project_name = request.form.get("project_name", "").strip()
    uploaded_file = request.files.get("efile")

    if not project_name:
        flash("Bitte einen Projektnamen eingeben.")
        return redirect(url_for("index"))

    if not uploaded_file or not uploaded_file.filename:
        flash("Bitte eine .01S-Datei auswählen.")
        return redirect(url_for("index"))

    if not _allowed(uploaded_file.filename):
        flash("Nur .01S- oder .crbx-Dateien sind erlaubt.")
        return redirect(url_for("index"))

    # Datei in temporären Puffer lesen (kein Schreiben auf Disk nötig)
    raw_bytes = uploaded_file.read()

    # parse_sia451 erwartet einen Dateipfad – wir schreiben in einen NamedTemporaryFile
    import tempfile, os
    suffix = os.path.splitext(uploaded_file.filename)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(raw_bytes)
        tmp_path = tmp.name

    try:
        positions = parse_sia451(tmp_path)
    finally:
        os.unlink(tmp_path)

    if not positions:
        flash("Keine Positionen in der Datei gefunden.")
        return redirect(url_for("index"))

    wb = build_workbook(positions)

    # Workbook direkt in den Response-Stream schreiben
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in project_name)
    download_name = f"{safe_name}.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5002, debug=False)
