# Excel - Data cleanup utility

This repository contains a small Python utility to extract tabular data from Word (`.docx`), PDF (`.pdf`), and Excel (`.xlsx`) files, clean and normalize the tables, and export the cleaned tables into an Excel workbook.

The main processing code is in `cleanup.py`.

## Features

- Extract tables from `.docx`, `.pdf`, and `.xlsx` inputs.
- Normalize misaligned cells (shifts non-empty cells left when text was split across columns).
- Drop fully-empty rows while preserving rows that contain some missing cells.
- Paragraph/key:value parsing: detects repeated left-side keys  and converts them into columns when present.
- Outputs a multi-sheet Excel workbook with one cleaned table per sheet.

## Requirements

- Python 3.8+
- See `requirements.txt` for the pinned Python packages used in the project. Minimum tools typically include:
  - pandas
  - openpyxl
  - python-docx (optional, for `.docx`)
  - pdfplumber (optional, for `.pdf`)

Install dependencies using pip:

```powershell
python -m pip install -r requirements.txt
```

If you don't need `.pdf` or `.docx` support you can skip installing `pdfplumber` and `python-docx`.

## Usage

The main entry point is `process_file(file_bytes: bytes, filename: str) -> bytes` in `cleanup.py`.

- `file_bytes` - raw bytes of the uploaded file (read in binary mode).
- `filename` - the original filename (used to detect file type by extension).

The function returns an in-memory Excel file bytes object. Example usage from a script:

```python
from cleanup import process_file

# read a local file and process
with open('SampleWord2.docx', 'rb') as f:
    data = f.read()

out_bytes = process_file(data, 'SampleWord2.docx')

# write the result to disk
with open('cleaned_output.xlsx', 'wb') as out:
    out.write(out_bytes)
```

## Common behaviors and heuristics

- Empty rows: only rows that are completely empty (all cells empty/NaN) are removed. Rows with one or more missing cells are preserved.
- Header detection: the code attempts to detect a header row automatically. For paragraph-style key:value records the left-side keys are used to build column headers dynamically.
- Paragraph parsing: the parser looks for lines containing a `:` separator and treats the left side as the column name and the right side as the cell value. Repeated keys across paragraphs form rows.

## Running tests / smoke checks

You can run a small smoke test by importing the module in a Python REPL or script and calling `clean_table` on a sample `pandas.DataFrame`.

```powershell
python -c "import pandas as pd; from cleanup import clean_table; df = pd.DataFrame([[None, 'a'], ['b', None], [None, None]]); print(clean_table(df))"
```

## Notes

- The code aims to be resilient to many input formats but may need small tweaks for edge-cases in scanned PDFs or highly irregular Word tables.
- If you want stricter rules (for example, drop rows with any missing cell), change the cleaning policy in `cleanup.py`.

## License

This project doesn't include a license file. Add one if you plan to publish or share the code.
