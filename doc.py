from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from docx import Document
import io
import base64

app = FastAPI()

@app.post("/generate-doc")
async def generate_doc(request: Request):
    try:
        data = await request.json()
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

        return JSONResponse({
            "fileName": "response.docx",
            "fileContent": encoded,
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )
