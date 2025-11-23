from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Union, Any
import json
import base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="QIESI JSON → Excel API",
    version="1.2.3",  # bugfix: correctly handle {'value': {'transactions': [...]}}
    description="Convert JSON to Excel (Base64) + return rows for NAC."
)


class ConvertRequest(BaseModel):
    # Can be JSON string, object, or array (NAC may send any of these)
    jsonInput: Union[str, dict, list]


def normalise_to_rows(raw: Any) -> list[dict]:
    """
    Normalise incoming jsonInput into a list of row dicts.
    For this requirement we primarily support:
      { "value": { "transactions": [ {..}, {..}, ... ] } }

    Also handles:
      - { "transactions": [ ... ] }
      - [ {..}, {..} ]
      - { single: "object" } -> [ { single: "object" } ]
    """

    # If a string, first parse as JSON
    if isinstance(raw, str):
        try:
            data = json.loads(raw)
        except Exception as ex:
            raise HTTPException(
                status_code=400,
                detail=f"jsonInput is not valid JSON string: {str(ex)}"
            )
    else:
        data = raw

    # Now data is dict or list (hopefully)
    # 1) Preferred structure: {"value": {"transactions": [ ... ]}}
    if isinstance(data, dict):
        if (
            "value" in data
            and isinstance(data["value"], dict)
            and "transactions" in data["value"]
        ):
            tx = data["value"]["transactions"]
            if not isinstance(tx, list):
                raise HTTPException(
                    status_code=400,
                    detail="value.transactions must be an array of objects."
                )
            rows = tx

        # 2) Slight variant: {"transactions": [ ... ]}
        elif "transactions" in data and isinstance(data["transactions"], list):
            rows = data["transactions"]

        # 3) Single object -> one row
        else:
            rows = [data]

    elif isinstance(data, list):
        # Already a list of row objects
        rows = data
    else:
        raise HTTPException(
            status_code=400,
            detail="jsonInput must be a JSON object, array, or JSON string."
        )

    # Validate rows are objects
    for idx, r in enumerate(rows):
        if not isinstance(r, dict):
            raise HTTPException(
                status_code=400,
                detail=f"Each row/transaction must be an object; item at index {idx} is {type(r).__name__}."
            )

    return rows


@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    # 1) Normalise to rows based on your structure
    rows = normalise_to_rows(req.jsonInput)

    # 2) Collect headers (all keys across all rows)
    headers_set = set()
    for item in rows:
        headers_set.update(item.keys())

    # Stable header order (sorted) so Excel is deterministic
    headers = sorted(headers_set)

    # 3) Create Excel
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"

        # Header row
        ws.append(headers)

        # Data rows
        for item in rows:
            ws.append([item.get(h) for h in headers])

        # Save to memory and Base64 encode
        bio = BytesIO()
        wb.save(bio)
        excel_b64 = base64.b64encode(bio.getvalue()).decode("utf-8")

    except Exception as ex:
        raise HTTPException(
            status_code=500,
            detail=f"Excel generation failed: {str(ex)}"
        )

    # 4) Filename with UTC timestamp
    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")

    return {
        "fileName": f"QIESI-{ts}.xlsx",
        "excelFile": excel_b64,
        "rows": rows
    }


@app.get("/health")
def health():
    return {"status": "ok"}
