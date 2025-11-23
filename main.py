from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json
import base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI JSON → Excel API",
    version="1.2.2",
    description="Accepts raw JSON text or JSON object/array and converts it to Excel + returns Nintex File Object + rows."
)

class ConvertRequest(BaseModel):
    jsonInput: object  # Allow any type: string, dict, list

@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):

    raw = req.jsonInput

    # STEP 1 — Normalize into dict or list
    # -----------------------------------
    # Case A: Already an object or list
    if isinstance(raw, (dict, list)):
        data = raw

    # Case B: Comes in as a string → parse it safely
    elif isinstance(raw, str):
        try:
            data = json.loads(raw)
        except Exception as ex:
            raise HTTPException(400, f"Invalid JSON string: {str(ex)}")
    else:
        raise HTTPException(400, "jsonInput must be JSON text, object, or array")

    # STEP 2 — Ensure final structure is a list
    # ----------------------------------------
    if isinstance(data, dict):
        rows = [data]
    elif isinstance(data, list):
        rows = data
    else:
        raise HTTPException(400, "Parsed JSON must be an object or array")

    # STEP 3 — Build Excel headers
    # ----------------------------
    headers = set()
    for item in rows:
        if not isinstance(item, dict):
            raise HTTPException(400, "Array items must be JSON objects")
        headers.update(item.keys())
    headers = list(headers)

    # STEP 4 — Create Excel file
    # --------------------------
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"

        ws.append(headers)

        for item in rows:
            ws.append([item.get(h) for h in headers])

        bio = BytesIO()
        wb.save(bio)
        excel_b64 = base64.b64encode(bio.getvalue()).decode()

    except Exception as ex:
        raise HTTPException(500, f"Excel generation failed: {str(ex)}")

    # STEP 5 — Return Nintex File Object + rows
    # -----------------------------------------
    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")

    return {
        "fileName": f"QIESI-{ts}.xlsx",
        "excelFile": excel_b64,
        "rows": rows
    }

@app.get("/health")
def health():
    return {"status": "ok"}
