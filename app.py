import os
import uuid
import threading
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template

app = Flask(__name__)

UPLOAD_DIR = Path("/tmp/pdf2docx")
UPLOAD_DIR.mkdir(exist_ok=True)

jobs: dict[str, dict] = {}  # job_id -> {status, out_path, error}


def do_convert(job_id: str, pdf_path: Path, out_path: Path):
    try:
        from pdf2docx import Converter
        cv = Converter(str(pdf_path))
        cv.convert(str(out_path), start=0, end=None)
        cv.close()
        jobs[job_id] = {"status": "done", "out_path": str(out_path)}
    except Exception as e:
        jobs[job_id] = {"status": "error", "error": str(e)}
    finally:
        pdf_path.unlink(missing_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify(error="No file provided"), 400
    f = request.files["file"]
    if not f.filename.lower().endswith(".pdf"):
        return jsonify(error="Only PDF files are accepted"), 400

    job_id = uuid.uuid4().hex
    pdf_path = UPLOAD_DIR / f"{job_id}.pdf"
    out_path = UPLOAD_DIR / f"{job_id}.docx"
    f.save(pdf_path)

    jobs[job_id] = {"status": "processing"}
    threading.Thread(target=do_convert, args=(job_id, pdf_path, out_path), daemon=True).start()
    return jsonify(job_id=job_id)


@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="Unknown job"), 404
    if job["status"] == "done":
        return jsonify(status="done")
    if job["status"] == "error":
        return jsonify(status="error", error=job["error"])
    return jsonify(status="processing")


@app.route("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify(error="Not ready"), 404
    out_path = job["out_path"]
    original_name = request.args.get("name", "converted.docx")
    stem = Path(original_name).stem
    return send_file(
        out_path,
        as_attachment=True,
        download_name=f"{stem}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=True, port=port)
