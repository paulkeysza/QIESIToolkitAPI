from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI Toolkit API",
    version="1.0.0",
    description="JSON → Excel converter for NAC demos"
)

class ConvertRequest(BaseModel):
    jsonInput: str


@app.get("/health", tags=["System"])
def health():
    return {"status": "ok"}


@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="Invalid JSON")

    # Normalize into list
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise HTTPException(status_code=400, detail="JSON must be list or object")

    # Flatten nested dicts
    flat_data = []
    for item in data:
        row = {}

        for key, value in item.items():
            if isinstance(value, dict):
                # flatten nested dict
                for sub_key, sub_value in value.items():
                    row[f"{key}_{sub_key}"] = sub_value
            else:
                row[key] = value

        flat_data.append(row)

    # Build Excel
    try:
        wb = Workbook()
        ws = wb.active

        headers = set()
        for row in flat_data:
            headers.update(row.keys())
        headers = list(headers)

        ws.append(headers)
        for row in flat_data:
            ws.append([row.get(h) for h in headers])

        bio = BytesIO()
        wb.save(bio)
        excel_b64 = base64.b64encode(bio.getvalue()).decode()

        ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")
        filename = f"QIESI-{ts}.xlsx"

        return {"fileName": filename, "excelFile": excel_b64}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))