@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="Invalid JSON format")

    # 🔵 NEW: Auto-detect nested array case
    # e.g. { "transactions": [ {...}, {...} ] }
    if isinstance(data, dict) and len(data) == 1:
        only_value = list(data.values())[0]
        if isinstance(only_value, list):
            data = only_value

    # 🔵 If still a dict, convert to single-row list
    if isinstance(data, dict):
        data = [data]

    if not isinstance(data, list):
        raise HTTPException(status_code=400, detail="JSON must be object, list, or object-with-array")

    # 🔵 Validate every item is a dict
    for item in data:
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail="Each item in the list must be a JSON object")

    # Build Excel
    try:
        wb = Workbook()
        ws = wb.active

        # Collect all unique headers
        headers = set()
        for item in data:
            headers.update(item.keys())

        headers = list(headers)
        ws.append(headers)

        # Append rows
        for item in data:
            row = [
                item.get(h) if not isinstance(item.get(h), (list, dict)) else json.dumps(item.get(h))
                for h in headers
            ]
            ws.append(row)

        # Encode to Base64
        buff = BytesIO()
        wb.save(buff)
        encoded = base64.b64encode(buff.getvalue()).decode()

        ts = datetime.utcnow().strftime("%Y%m%d%H%M%S")
        return {
            "fileName": f"QIESI-{ts}.xlsx",
            "excelFile": encoded
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel generation error: {str(e)}")
