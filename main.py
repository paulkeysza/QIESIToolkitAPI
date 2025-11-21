from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI Toolkit API",
    version="1.0.0",
    description="Core API for QIESI automation tools. Includes JSON→Excel conversion and future QIESI features."
)

class ConvertRequest(BaseModel):
    jsonInput: str

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/convert")
def convert(req: ConvertRequest):
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="Invalid JSON format")

    if isinstance(data, dict):
        data = [data]

    if not isinstance(data, list):
        raise HTTPException(status_code=400, detail="JSON must be object or list")

    wb = Workbook()
    ws = wb.active

    headers = set()
    for item in data:
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail="Each item must be a JSON object")
        headers.update(item.keys())

    headers = list(headers)
    ws.append(headers)

    for item in data:
        ws.append([item.get(h) for h in headers])

    buff = BytesIO()
    wb.save(buff)
    encoded = base64.b64encode(buff.getvalue()).decode()

    return {
        "fileName": f"QIESI-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.xlsx",
        "excelFile": encoded
    }
