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

def read_pdf_tables(file_bytes: bytes):
    """Extracts tables from a PDF file using pdfplumber."""
    tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            extracted = page.extract_table()
            if extracted:
                tables.append(pd.DataFrame(extracted))
    return tables


def read_excel(file_bytes: bytes):
    """Reads an Excel file into DataFrames."""
    excel_data = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    return list(excel_data.values())


def clean_table(df: pd.DataFrame) -> pd.DataFrame:
    """Cleans up merged, blank, or misaligned rows."""
    # Drop completely empty rows
    df = df.dropna(how="all").reset_index(drop=True)

    # Fill misaligned cells: shift non-empty cells left
    df = df.apply(lambda row: pd.Series([x for x in row if pd.notna(x)] + [None]*(len(row)-sum(pd.notna(row)))), axis=1)

    # Merge similar header rows
    if df.shape[0] > 1:
        header_candidate = df.iloc[0].astype(str).str.lower()
        next_candidate = df.iloc[1].astype(str).str.lower()
        if (header_candidate == next_candidate).sum() > len(df.columns)//2:
            df = df.drop(1).reset_index(drop=True)

    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df[1:].reset_index(drop=True)
    return df


def process_file(file_bytes: bytes, filename: str) -> bytes:
    """Main processing function: detect type, clean, and return Excel bytes."""
    if filename.endswith(".docx"):
        tables = read_docx_tables(file_bytes)
    elif filename.endswith(".pdf"):
        tables = read_pdf_tables(file_bytes)
    elif filename.endswith(".xlsx"):
        tables = read_excel(file_bytes)
    else:
        raise ValueError("Unsupported file type")

    cleaned_tables = []
    for t in tables:
        cleaned_tables.append(clean_table(t))

    # Combine all cleaned tables into one Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, table in enumerate(cleaned_tables):
            sheet_name = f"Table_{i+1}"
            table.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()
