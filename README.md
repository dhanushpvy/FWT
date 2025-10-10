# Excel - Data Cleanup Utility (with Web Interface)

This repository contains a complete Python-based data cleanup utility with a web interface that allows users to upload Word (.docx), PDF (.pdf), and Excel (.xlsx) files, automatically extract and clean tabular data, and download the cleaned Excel output.

The backend logic is handled in cleanup.py, while the frontend web app provides an interactive and user-friendly interface for file upload and result download.

## Features

- Web Application Interface
  - Simple and intuitive upload form built with FastAPI and HTML.
  - Users can upload .docx, .pdf, or .xlsx files.
  - Instantly receive a cleaned Excel file as a downloadable output.
  
- Core Data Processing
  - Extracts tables from Word, PDF, and Excel files.
  - Normalizes misaligned or broken table cells.
  - Drops fully empty rows but keeps partially filled ones.
  - Supports paragraph-style key:value data transformation into tabular format.
  - Outputs multi-sheet Excel files (one sheet per extracted table).

## Project Structure

ExcelPractice/
│
├── cleanup.py          # Core table extraction & cleaning logic
├── main.py             # FastAPI server (handles file uploads & responses)
├── templates/          # HTML templates for frontend interface
│   └── index.html
├── static/             # CSS, JS, and assets
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation

## Requirements

- Python 3.8+
- Required packages (install using pip):

python -m pip install -r requirements.txt

Typical dependencies include:
- pandas
- openpyxl
- fastapi
- uvicorn
- python-docx (optional – for .docx support)
- pdfplumber (optional – for .pdf support)

## Usage

1. Run the Web App
Start the FastAPI server:
uvicorn main:app --reload

Then open your browser and go to:
http://127.0.0.1:8000

2. Upload & Clean
- Upload a .docx, .pdf, or .xlsx file.
- The server processes it using cleanup.py.
- Download the cleaned Excel file instantly.

## Core Function (Backend)

The main function used internally is:

process_file(file_bytes: bytes, filename: str) -> bytes

It:
- Detects file type by extension.
- Extracts and cleans tables.
- Returns an Excel file (as bytes) ready for download.

## Cleaning Rules & Heuristics

- Empty rows → removed only if all cells are empty.  
- Partial rows → retained for data integrity.  
- Header detection → auto-detects based on text density and alignment.  
- Key:Value parsing → automatically converts paragraph-style data to table format.

## Quick Local Test

Run a smoke test:

python -c "import pandas as pd; from cleanup import clean_table; df = pd.DataFrame([[None,'a'],['b',None],[None,None]]); print(clean_table(df))"

## Notes

- The web app is designed for non-technical users, enabling easy cleanup without running Python scripts manually.
- Backend functions are modular and can be reused in other automation or AI workflows.

## License

This project currently doesn’t include a license. Add one (e.g., MIT License) if you plan to make it public.
