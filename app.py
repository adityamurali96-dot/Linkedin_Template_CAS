"""
Flask web server wrapping crowe_formatter for Railway deployment.

Endpoints:
    GET  /health           → Health check (200 OK)
    POST /audit            → Upload .docx, returns audited .docx
    POST /convert          → Upload .docx (+optional title), returns converted .docx
"""

import os
import tempfile

from flask import Flask, request, send_file, jsonify

from crowe_formatter import audit_document, convert_document, TEMPLATE_PATH

app = Flask(__name__)


@app.route("/health", methods=["GET"])
def health():
    template_ok = TEMPLATE_PATH.exists()
    return jsonify({"status": "ok", "template_loaded": template_ok}), 200


@app.route("/audit", methods=["POST"])
def audit():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Send a .docx as 'file'."}), 400

    uploaded = request.files["file"]
    if not uploaded.filename.lower().endswith(".docx"):
        return jsonify({"error": "File must be a .docx document."}), 400

    tmp_in = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_out = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    try:
        uploaded.save(tmp_in.name)
        result = audit_document(tmp_in.name, tmp_out.name)
        return send_file(
            tmp_out.name,
            as_attachment=True,
            download_name="audited_output.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    finally:
        _safe_remove(tmp_in.name)
        # tmp_out is cleaned up by send_file / after response


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Send a .docx as 'file'."}), 400

    uploaded = request.files["file"]
    if not uploaded.filename.lower().endswith(".docx"):
        return jsonify({"error": "File must be a .docx document."}), 400

    title = request.form.get("title")

    tmp_in = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_out = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    try:
        uploaded.save(tmp_in.name)
        convert_document(tmp_in.name, tmp_out.name, title=title)
        return send_file(
            tmp_out.name,
            as_attachment=True,
            download_name="converted_output.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    finally:
        _safe_remove(tmp_in.name)


def _safe_remove(path):
    try:
        os.remove(path)
    except OSError:
        pass


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
