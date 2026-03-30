from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from html2docx import html2docx
import io

app = FastAPI()


@app.post("/convert-html-to-docx")
async def convert_html_to_docx(file: UploadFile = File(...)):
    try:
        # Validate file type
        if not file.filename.endswith(".html"):
            raise HTTPException(status_code=400, detail="Only HTML files are allowed")

        # Read HTML content
        html_content = await file.read()
        html_string = html_content.decode("utf-8")

        # Convert HTML → DOCX
        docx_bytes = html2docx(html_string)

        # Convert to stream
        file_stream = io.BytesIO(docx_bytes)

        # Return as downloadable file
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename=converted.docx"
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
