from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

# ------------------------------------------
# FASTAPI APP
# ------------------------------------------
app = FastAPI(
    title="QIESI Toolkit API",
    version="1.0.0",
    description="Core API powering JSON-to-Excel conversion and future QIESI features."
)

# ------------------------------------------
# REQUEST MODEL
# ------------------------------------------
class ConvertRequest(BaseModel):
    jsonInput: str


# ------------------------------------------
# HEALTH CHECK
# ------------------------------------------
@app.get("/health", tags=["System"])
def health():
    return {"status": "ok"}


# ------------------------------------------
# MAIN CONVERT ENDPOINT
# ------------------------------------------
@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):

    # --------------------------------------
    # 1. LOAD JSON
    # --------------------------------------
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="Invalid JSON format")

    # --------------------------------------
    # 2. AUTO-DETECT NESTED ARRAY CASE
    # Allows inputs like:
    # { "transactions": [ {...}, {...} ] }
    # --------------------------------------
    if isinstance(data, dict) and len(data) == 1:
        only_value = list(data.values())[0]
        if isinstance(only_value, list):
            data = only_value

    # --------------------------------------
    # 3. SINGLE ROW CASE (convert to list)
    # --------------------------------------
    if isinstance(data, dict):
        data = [data]

    # --------------------------------------
    # 4. VALIDATE FINAL STRUCTURE
    # --------------------------------------
    if not isinstance(data, list):
        raise HTTPException(
            status_code=400,
            detail="JSON must be: array, object, or object containing exactly one array."
        )

    for item in data:
        if not isinstance(item, dict):
            raise HTTPException(
                status_code=400,
                detail="Each item in the list must be a JSON object."
            )

    # --------------------------------------
    # 5. BUILD EXCEL IN MEMORY
    # --------------------------------------
    try:
        wb = Workbook()
        ws = wb.active

        # Collect all unique column names
        headers = set()
        for item in data:
            headers.update(item.keys())
        headers = list(headers)

        # Header row
        ws.append(headers)

        # Data rows
        for item in data:
            row = []
            for h in headers:
                value = item.get(h)

                # Excel-safe: convert nested dicts/lists to JSON strings
                if isinstance(value, (dict, list)):
                    value = json.dumps(value)

                row.append(value)

            ws.append(row)

        # Convert workbook to Base64
        buff = BytesIO()
        wb.save(buff)
        encoded = base64.b64encode(buff.getvalue()).decode()

        timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")

        return {
            "fileName": f"QIESI-{timestamp}.xlsx",
            "excelFile": encoded
        }

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Excel generation error: {str(e)}"
        )
