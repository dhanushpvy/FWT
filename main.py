# main.py
from fastapi import FastAPI
from fastapi.responses import HTMLResponse

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
