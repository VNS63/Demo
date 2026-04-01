import base64
import json
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request
from pdf2docx import Converter

app = Flask(__name__)

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _safe_docx_name(filename: str | None) -> str:
    if not filename:
        return "converted.docx"
    clean = filename.strip().replace("/", "_").replace("\\", "_")
    if not clean.lower().endswith(".docx"):
        clean += ".docx"
    return clean or "converted.docx"


def _extract_pdf_base64(payload: dict) -> str:
    # Primary contract requested by user: { "file": "<base64-pdf>" }
    base64_text = payload.get("file")
    if not isinstance(base64_text, str) or not base64_text.strip():
        raise ValueError("JSON must include non-empty 'file' field with Base64 PDF content.")

    text = base64_text.strip()
    if text.lower().startswith("data:") and "base64," in text:
        text = text.split("base64,", 1)[1].strip()

    text = "".join(text.split())
    if not text:
        raise ValueError("Base64 PDF content is empty after whitespace cleanup.")

    pad = len(text) % 4
    if pad:
        text += "=" * (4 - pad)
    return text


def _decode_pdf_bytes_from_json_body() -> tuple[bytes, str]:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        raise ValueError("Expected JSON object body.")

    output_filename = _safe_docx_name(payload.get("filename"))
    encoded_pdf = _extract_pdf_base64(payload)

    try:
        pdf_bytes = base64.b64decode(encoded_pdf, validate=True)
    except Exception:
        try:
            pdf_bytes = base64.urlsafe_b64decode(encoded_pdf)
        except Exception as exc:
            raise ValueError(f"Invalid Base64 PDF content: {exc}") from exc

    if not pdf_bytes.startswith(b"%PDF"):
        raise ValueError("Decoded content is not a valid PDF stream (missing %PDF header).")

    return pdf_bytes, output_filename


@app.get("/")
def health():
    return jsonify({"status": "ok", "service": "pdf-to-docx-flask"})


@app.post("/convert-pdf-to-docx")
def convert_pdf_to_docx():
    try:
        pdf_bytes, output_filename = _decode_pdf_bytes_from_json_body()
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        pdf_path = tmp_path / "input.pdf"
        docx_temp_path = tmp_path / output_filename
        pdf_path.write_bytes(pdf_bytes)

        converter = Converter(str(pdf_path))
        try:
            # Convert all pages; preserve as much layout as possible.
            converter.convert(str(docx_temp_path), start=0, end=None)
        except Exception as exc:
            return jsonify({"error": f"PDF to DOCX conversion failed: {exc}"}), 500
        finally:
            converter.close()

        if not docx_temp_path.exists():
            return jsonify({"error": "Conversion failed: DOCX file was not created."}), 500

        # Save a copy in project output folder for quick local access.
        docx_bytes = docx_temp_path.read_bytes()
        saved_path = OUTPUT_DIR / output_filename
        saved_path.write_bytes(docx_bytes)

        return jsonify(
            {
                "message": "Conversion successful",
                "filename": output_filename,
                "docx_content": base64.b64encode(docx_bytes).decode("utf-8"),
            }
        )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
