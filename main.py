# main.py
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse, HTMLResponse
from cleanup import process_file
import io

app = FastAPI(title="Smart Table Alignment Tool")

@app.get("/", response_class=HTMLResponse)
async def home():
    return """
    <html>
        <head><title>Table Alignment Tool</title></head>
        <body style="font-family:sans-serif;text-align:center;padding:50px">
            <h2>📊 Smart Table Alignment Tool</h2>
            <form action="/upload" enctype="multipart/form-data" method="post">
                <input name="file" type="file" accept=".docx,.pdf,.xlsx" required/>
                <br><br>
                <input type="submit" value="Upload & Clean"/>
            </form>
        </body>
    </html>
    """
@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    file_bytes = await file.read()
    try:
        cleaned_bytes = process_file(file_bytes, file.filename)
        return StreamingResponse(
            io.BytesIO(cleaned_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=cleaned_{file.filename.split('.')[0]}.xlsx"}
        )
    except Exception as e:
        return {"error": str(e)}