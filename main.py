from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI JSON → Excel API",
    version="1.2.0",
    description="Convert JSON to Excel (Base64) + return rows for NAC."
)

class ConvertRequest(BaseModel):
    jsonInput: dict | list

@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    # normalise input
    data = req.jsonInput
    if isinstance(data, dict):
        rows = [data]
    elif isinstance(data, list):
        rows = data
    else:
        raise HTTPException(400, "jsonInput must be an object or array")

    # collect headers
    headers = set()
    for item in rows:
        if not isinstance(item, dict):
            raise HTTPException(400, "Array items must be objects")
        headers.update(item.keys())
    headers = list(headers)

    # create Excel
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

    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")

    return {
        "fileName": f"QIESI-{ts}.xlsx",
        "excelFile": excel_b64,
        "rows": rows
    }

@app.get("/health")
def health():
    return {"status": "ok"}
