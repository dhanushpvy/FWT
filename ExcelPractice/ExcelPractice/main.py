# main.py
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
import io
from cleanup import process_file

app = FastAPI(title="Smart Table Cleanup Tool")


@app.get("/", response_class=HTMLResponse)
async def home():
    """Simple HTML upload page."""
    return """
    <html>
        <head>
            <title>Table Cleanup Tool</title>
        </head>
        <body>
            <h2>Upload DOCX, PDF, or Excel file</h2>
            <form action="/upload" enctype="multipart/form-data" method="post">
                <input name="file" type="file" accept=".docx,.pdf,.xls,.xlsx"/>
                <input type="submit" value="Upload and Clean"/>
            </form>
        </body>
    </html>
    """


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Endpoint to upload file, clean tables, and download Excel."""
    try:
        file_bytes = await file.read()
        cleaned_bytes = process_file(file_bytes, file.filename)
        return StreamingResponse(
            io.BytesIO(cleaned_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=cleaned_{file.filename}.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error processing file: {str(e)}")
