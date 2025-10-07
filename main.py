from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
import io
from cleanup import cleanup_sheet

app = FastAPI(title="Excel Table Cleanup API")

@app.post("/cleanup")
async def cleanup(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail="Only .xls/.xlsx files supported")

    contents = await file.read()
    in_mem = io.BytesIO(contents)

    try:
        xls = pd.read_excel(in_mem, sheet_name=None, engine='openpyxl')
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Cannot read Excel: {e}")

    cleaned_sheets = {}
    for sheet_name, df in xls.items():
        cleaned_sheets[sheet_name] = cleanup_sheet(df)

    out_mem = io.BytesIO()
    with pd.ExcelWriter(out_mem, engine='openpyxl') as writer:
        for sheet_name, df in cleaned_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()

    out_mem.seek(0)
    filename = f"cleaned_{file.filename}"
    headers = {'Content-Disposition': f'attachment; filename=\"{filename}\"'}
    return StreamingResponse(out_mem, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers=headers)