# QIESI JSON → Excel API v1.2.4
from fastapi import FastAPI, HTTPException, Body
from typing import Any, Dict, List
import json
import base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI JSON → Excel API",
    version="1.2.4",
    description="Accepts flexible JSON (string/object/array), auto-detects rows, converts to Excel, returns Base64 + rows."
)

def _parse_json_input(body: Any) -> Any:
    if isinstance(body, dict) and "jsonInput" in body:
        raw = body["jsonInput"]
        if isinstance(raw, str):
            try:
                return json.loads(raw)
            except Exception as ex:
                raise HTTPException(400, f"Invalid JSON string: {str(ex)}")
        return raw

    if isinstance(body, str):
        try:
            return json.loads(body)
        except Exception as ex:
            raise HTTPException(400, f"Body is not valid JSON: {str(ex)}")

    return body

def _extract_rows(data: Any) -> List[Dict[str, Any]]:
    if isinstance(data, dict):
        if isinstance(data.get("value"), dict) and isinstance(data["value"].get("transactions"), list):
            return [r for r in data["value"]["transactions"] if isinstance(r, dict)]

    if isinstance(data, dict) and isinstance(data.get("transactions"), list):
        return [r for r in data["transactions"] if isinstance(r, dict)]

    if isinstance(data, list):
        return [r for r in data if isinstance(r, dict)]

    if isinstance(data, dict):
        return [data]

    raise HTTPException(400, "Unsupported structure — expected transactions list, array, or object.")

def _rows_to_excel_b64(rows: List[Dict[str, Any]]) -> str:
    if not rows:
        raise HTTPException(400, "No rows found.")

    headers = list({k for r in rows for k in r.keys()})

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h) for h in headers])

        bio = BytesIO()
        wb.save(bio)
        return base64.b64encode(bio.getvalue()).decode("utf-8")
    except Exception as ex:
        raise HTTPException(500, f"Excel generation failed: {str(ex)}")

@app.post("/convert")
async def convert(body: Any = Body(...)):
    data = _parse_json_input(body)
    rows = _extract_rows(data)
    excel_b64 = _rows_to_excel_b64(rows)
    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")

    return {
        "fileName": f"QIESI-{ts}.xlsx",
        "excelFile": excel_b64,
        "rows": rows
    }

@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.2.4"}