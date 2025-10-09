#cleanup.py
import pandas as pd
import pdfplumber
from docx import Document
import io
import re


def read_docx_tables(file_bytes: bytes):
    """Extracts tables from a Word document (.docx)."""
    doc = Document(io.BytesIO(file_bytes))
    tables = []
    # --- Extract actual tables ---
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        if data:
            tables.append(pd.DataFrame(data))

    # --- Extract paragraph data (Name:..., Age:..., City:...) ---
    paragraph_data = extract_paragraph_data(doc)
    if paragraph_data is not None:
        tables.append(paragraph_data)

    return tables



def extract_paragraph_data(doc):
    """Detect paragraph data like Name:Abi, Age:20, City:Madurai."""
    # Split paragraphs and keep even empty lines to detect record breaks
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    lines = [p for p in paragraphs if p or p == ""]

    records = []
    current = {}

    # Detect available keys dynamically (Name:, Age:, etc.)
    key_pattern = re.compile(r"^(\w+)\s*:")
    all_keys = set()

    # First pass – collect all possible keys
    for line in lines:
        m = key_pattern.match(line)
        if m:
            all_keys.add(m.group(1).strip())

    # Second pass – build records
    for line in lines:
        if not line.strip():  # blank line = end of one record
            if current:
                records.append(current)
                current = {}
            continue

        if ":" in line:
            k, v = line.split(":", 1)
            k, v = k.strip(), v.strip()
            current[k] = v

    if current:
        records.append(current)

    # If no structured key-value data found
    if not records or len(all_keys) < 2:
        return None

    # Convert to DataFrame (fill missing values with None)
    df = pd.DataFrame(records)
    for k in all_keys:
        if k not in df.columns:
            df[k] = None

    # Reorder columns alphabetically
    df = df[[c for c in sorted(df.columns)]]
    return df
