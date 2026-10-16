"""Web front end.  Run:  python app.py   then open http://127.0.0.1:5000"""
import io
import re
import tempfile
import uuid
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

import core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
STORE = Path(tempfile.gettempdir()) / "resource_sheet_generator"
STORE.mkdir(exist_ok=True)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/inspect/master")
def inspect_master():
    f = request.files.get("file")
    if not f:
        return jsonify(ok=False, error="No file received")
    try:
        return jsonify(ok=True, **core.inspect_master(f.read(), f.filename))
    except core.ProcessingError as e:
        return jsonify(ok=False, error=str(e))


@app.post("/api/inspect/resource")
def inspect_resource():
    f = request.files.get("file")
    if not f:
        return jsonify(ok=False, error="No file received")
    try:
        return jsonify(ok=True, **core.inspect_resource(f.read(), f.filename))
    except core.ProcessingError as e:
        return jsonify(ok=False, error=str(e))


@app.post("/api/generate")
def generate():
    master = request.files.get("master")
    resources = [f for f in request.files.getlist("resources") if f and f.filename]
    if not master or not master.filename:
        return jsonify(errors=["Please upload the master Sangraha attendance sheet"], warnings=[], employees=[])
    if not resources:
        return jsonify(errors=["Please upload at least one resource sheet"], warnings=[], employees=[])

    override = None
    m = re.match(r"^(\d{4})-(\d{2})$", request.form.get("month", "").strip())
    if m:
        override = (int(m.group(1)), int(m.group(2)))

    result = core.process((master.filename, master.read()), [(f.filename, f.read()) for f in resources], override)
    payload = {k: result[k] for k in ("errors", "warnings", "employees", "month")}
    if not result["errors"] and result["files"]:
        token = uuid.uuid4().hex
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for path, data in result["files"]:
                z.writestr(path, data)
        (STORE / f"{token}.zip").write_bytes(buf.getvalue())
        payload["download"] = f"/download/{token}"
        payload["file_count"] = len(result["files"])
    return jsonify(payload)


@app.get("/download/<token>")
def download(token):
    if not re.fullmatch(r"[0-9a-f]{32}", token):
        return "Not found", 404
    path = STORE / f"{token}.zip"
    if not path.exists():
        return "File expired - please generate again", 404
    return send_file(path, as_attachment=True, download_name="Resource Sheets.zip")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
