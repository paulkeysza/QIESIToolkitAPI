from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json
import base64
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import cast
from datetime import datetime
from fastapi.responses import RedirectResponse

app = FastAPI(
    title="QIESI Toolkit API",
    version="1.1.0",
    description="API services for QIESI Toolkit including JSON → Excel conversion."
)


class ConvertRequest(BaseModel):
    jsonInput: str


@app.get("/", include_in_schema=False)
def root():
    return RedirectResponse(url="/docs")


@app.get("/health", tags=["System"])
def health():
    return {"status": "ok"}


@app.get("/info", tags=["System"])
def info():
    return {
        "name": "QIESI Toolkit API",
        "version": "1.1.0",
        "author": "Paul Keys",
        "routes": {
            "root": "/",
            "docs": "/docs",
            "openapi": "/openapi.json",
            "convert": "/convert",
            "health": "/health"
        }
    }


@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    # Parse JSON string
    try:
        raw = json.loads(req.jsonInput)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid JSON. Provide valid JSON in jsonInput.")

    # Normalise into a list of row dictionaries
    # Case 1: {"transactions": [ {...}, {...} ]}  -> use the list value
    if isinstance(raw, dict) and len(raw) == 1:
        sole_value = next(iter(raw.values()))
        if isinstance(sole_value, list):
            data = sole_value
        else:
            data = [raw]
    elif isinstance(raw, dict):
        data = [raw]
    elif isinstance(raw, list):
        data = raw
    else:
        raise HTTPException(status_code=400, detail="JSON must be an object or an array of objects.")

    # Ensure all items are objects
    for item in data:
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail="Each array item must be a JSON object.")

    # Build flattened rows for Excel and for the 'rows' output
    headers = []
    rows = []

    for row in data:
        flat_row = {}
        for key, value in row.items():
            # For Excel, nested structures must be stringified
            if isinstance(value, (dict, list)):
                flat_row[key] = json.dumps(value)
            else:
                flat_row[key] = value

            if key not in headers:
                headers.append(key)

        rows.append(flat_row)

    # Create Excel workbook
    try:
        wb = Workbook()
        ws = cast(Worksheet, wb.active)
        ws.title = "Data"

        # Header row
        ws.append(headers)

        # Data rows
        for row in rows:
            excel_row = [row.get(h) for h in headers]
            ws.append(excel_row)

        bio = BytesIO()
        wb.save(bio)
        excel_b64 = base64.b64encode(bio.getvalue()).decode()

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel generation failed: {str(e)}")

    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    filename = f"QIESI-{ts}.xlsx"

    # Return file info + rows collection
    return {
        "fileName": filename,
        "excelFile": excel_b64,
        "rows": rows
    }
