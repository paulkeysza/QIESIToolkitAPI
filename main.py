from fastapi import FastAPI, HTTPException, Body
from typing import Any, Dict, List
import json
import base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI JSON → Excel API",
    version="1.2.3",
    description=(
        "Accepts flexible JSON input (including raw JSON strings), "
        "normalises it to rows, converts to Excel (.xlsx) and returns "
        "both the file (Base64) and the rows."
    ),
)


def _parse_json_input(body: Any) -> Any:
    """
    Normalise the inbound body into a Python object.

    Supported patterns:
    - { "jsonInput": "<string containing JSON>" }
    - { "jsonInput": { ... } } or [ ... ]
    - { "value": { "transactions": [ ... ] } }
    - { "transactions": [ ... ] }
    - [ { ... }, { ... } ]
    - { ...single object... }
    - "<string containing JSON>"
    """
    # 1) If body has jsonInput, prefer that
    if isinstance(body, dict) and "jsonInput" in body:
        raw = body["jsonInput"]

        # If it's a string, try to parse JSON
        if isinstance(raw, str):
            raw = raw.strip()
            if raw == "":
                raise HTTPException(400, "jsonInput cannot be empty.")
            try:
                return json.loads(raw)
            except Exception as ex:
                raise HTTPException(
                    400,
                    f"jsonInput is not valid JSON string: {str(ex)}",
                )
        # If it's already a dict/list, just use it
        return raw

    # 2) If body is a string, try to parse as JSON
    if isinstance(body, str):
        text = body.strip()
        if text == "":
            raise HTTPException(400, "Request body cannot be empty.")
        try:
            return json.loads(text)
        except Exception as ex:
            raise HTTPException(400, f"Body string is not valid JSON: {str(ex)}")

    # 3) Otherwise, body is already a parsed JSON structure (dict/list/etc.)
    return body


def _extract_rows(data: Any) -> List[Dict[str, Any]]:
    """
    Extract rows for Excel from the normalised JSON.

    Priority:
    1. data["value"]["transactions"] if present and is a list
    2. data["transactions"] if present and is a list
    3. data itself if it's a list
    4. data itself if it's a dict (wrap as single row)
    """
    # 1) Look for value.transactions (your primary pattern)
    if isinstance(data, dict):
        value = data.get("value")
        if isinstance(value, dict):
            tx = value.get("transactions")
            if isinstance(tx, list):
                # Ensure all items are dicts
                rows = []
                for item in tx:
                    if not isinstance(item, dict):
                        raise HTTPException(
                            400,
                            "Items in value.transactions must be objects.",
                        )
                    rows.append(item)
                return rows

    # 2) Look for top-level transactions
    if isinstance(data, dict) and "transactions" in data:
        tx = data["transactions"]
        if isinstance(tx, list):
            rows = []
            for item in tx:
                if not isinstance(item, dict):
                    raise HTTPException(
                        400,
                        "Items in transactions must be objects.",
                    )
                rows.append(item)
            return rows

    # 3) If data itself is a list, treat as rows
    if isinstance(data, list):
        rows = []
        for item in data:
            if not isinstance(item, dict):
                raise HTTPException(
                    400,
                    "Array items must be objects to convert to Excel.",
                )
            rows.append(item)
        return rows

    # 4) If data is a dict, treat as a single row
    if isinstance(data, dict):
        return [data]

    # Otherwise, we don't know what this is
    raise HTTPException(
        400,
        "Unsupported JSON structure. Expected value.transactions, "
        "transactions, array of objects, or a single object.",
    )


def _rows_to_excel_b64(rows: List[Dict[str, Any]]) -> str:
    """
    Convert list of dict rows to an Excel file and return as Base64 string.
    """
    if not rows:
        raise HTTPException(400, "No rows found to write to Excel.")

    # Collect all headers from all rows
    headers_set = set()
    for item in rows:
        headers_set.update(item.keys())
    headers = list(headers_set)

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"

        # Header row
        ws.append(headers)

        # Data rows
        for item in rows:
            ws.append([item.get(h) for h in headers])

        bio = BytesIO()
        wb.save(bio)
        excel_b64 = base64.b64encode(bio.getvalue()).decode("utf-8")
        return excel_b64
    except Exception as ex:
        raise HTTPException(500, f"Excel generation failed: {str(ex)}")


@app.post("/convert", tags=["Conversion"])
async def convert(body: Any = Body(...)):
    """
    Main conversion endpoint.

    - Accepts flexible JSON body (string, object, or array).
    - Normalises to rows (list of dicts).
    - Builds Excel file.
    - Returns fileName, excelFile (Base64), and rows.
    """
    # 1) Normalise inbound JSON
    data = _parse_json_input(body)

    # 2) Extract rows with value.transactions priority
    rows = _extract_rows(data)

    # 3) Convert rows to Excel
    excel_b64 = _rows_to_excel_b64(rows)

    # 4) File name with timestamp
    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    file_name = f"QIESI-{ts}.xlsx"

    return {
        "fileName": file_name,
        "excelFile": excel_b64,
        "rows": rows,
    }


@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.2.3"}
