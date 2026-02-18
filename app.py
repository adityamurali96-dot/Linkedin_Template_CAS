"""
Flask web server wrapping crowe_formatter for Railway deployment.

Endpoints:
    GET  /health           → Health check (200 OK / 503 if template missing)
    POST /audit            → Upload .docx, returns audited .docx
    POST /convert          → Upload .docx (+optional title), returns converted .docx
"""

import logging
import os
import tempfile

from flask import Flask, request, send_file, jsonify, after_this_request

from crowe_formatter import audit_document, convert_document, TEMPLATE_PATH

app = Flask(__name__)

# Configure logging so output is visible in gunicorn / Railway logs
gunicorn_logger = logging.getLogger("gunicorn.error")
app.logger.handlers = gunicorn_logger.handlers or logging.getLogger().handlers
app.logger.setLevel(gunicorn_logger.level or logging.INFO)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = app.logger


@app.route("/health", methods=["GET"])
def health():
    template_ok = TEMPLATE_PATH.exists()
    if not template_ok:
        log.error("Health check FAILED: template not found at %s", TEMPLATE_PATH)
        return jsonify({"status": "error", "template_loaded": False}), 503
    return jsonify({"status": "ok", "template_loaded": True}), 200


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
        log.info("Auditing document: %s", uploaded.filename)
        result = audit_document(tmp_in.name, tmp_out.name)

        @after_this_request
        def cleanup(response):
            _safe_remove(tmp_out.name)
            return response

        return send_file(
            tmp_out.name,
            as_attachment=True,
            download_name="audited_output.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception:
        log.exception("Error auditing document")
        _safe_remove(tmp_out.name)
        return jsonify({"error": "Failed to audit document."}), 500
    finally:
        _safe_remove(tmp_in.name)


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
        log.info("Converting document: %s (title=%s)", uploaded.filename, title)
        convert_document(tmp_in.name, tmp_out.name, title=title)

        @after_this_request
        def cleanup(response):
            _safe_remove(tmp_out.name)
            return response

        return send_file(
            tmp_out.name,
            as_attachment=True,
            download_name="converted_output.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception:
        log.exception("Error converting document")
        _safe_remove(tmp_out.name)
        return jsonify({"error": "Failed to convert document."}), 500
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
