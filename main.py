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
    version="1.0.0",
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
        "version": "1.0.0",
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
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="Invalid JSON. Provide valid JSON in jsonInput.")

    if isinstance(data, dict):
        data = [data]

    if not isinstance(data, list):
        raise HTTPException(status_code=400, detail="JSON must be an object or an array of objects.")

    for item in data:
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail="Each array item must be a JSON object.")

    headers = []
    for row in data:
        for key in row.keys():
            if key not in headers:
                headers.append(key)

    try:
        wb = Workbook()
        ws = cast(Worksheet, wb.active)
        ws.title = "Data"

        ws.append(headers)

        for row in data:
            excel_row = []
            for h in headers:
                value = row.get(h)
                if isinstance(value, (dict, list)):
                    value = json.dumps(value)
                excel_row.append(value)
            ws.append(excel_row)

        bio = BytesIO()
        wb.save(bio)

        excel_b64 = base64.b64encode(bio.getvalue()).decode()

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel generation failed: {str(e)}")

    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    filename = f"QIESI-{ts}.xlsx"

    return {
        "fileName": filename,
        "excelFile": excel_b64
    }