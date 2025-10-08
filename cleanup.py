# cleanup.py
import pandas as pd
import pdfplumber
from docx import Document
import io

def read_docx_tables(file_bytes: bytes):
    """Extracts tables from a Word document (.docx)."""
    doc = Document(io.BytesIO(file_bytes))
    tables = []
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        if data:
            tables.append(pd.DataFrame(data))
    return tables
