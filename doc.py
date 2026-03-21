import azure.functions as func
from docx import Document
import io
import base64
import json

app = func.FunctionApp()

@app.route(route="generate_doc", methods=["POST"])
def generate_doc(req: func.HttpRequest) -> func.HttpResponse:
    try:
        data = req.get_json()
        text_content = data.get("text", "")

        # Create document in memory
        doc = Document()
        doc.add_heading("Generated Document", 0)
        doc.add_paragraph(text_content)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_bytes = buffer.read()
        encoded = base64.b64encode(file_bytes).decode("utf-8")

        return func.HttpResponse(
            json.dumps({
                "fileName": "response.docx",
                "fileContent": encoded,
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }),
            mimetype="application/json"
        )

    except Exception as e:
        return func.HttpResponse(str(e), status_code=500)
