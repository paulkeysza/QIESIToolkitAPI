
# 📘 KeysieAPI — JSON → Excel Converter API

KeysieAPI is a lightweight **FastAPI** microservice that dynamically converts JSON input into an Excel file returned as Base64.  
It is purpose‑built for workflow engines like **Nintex Automation Cloud (NAC)** and integrates easily into automation solutions.

---

# 🚀 Features

- Accepts **raw JSON** (string)
- Automatically infers **columns** from JSON keys
- Supports:
  - Arrays of objects  
  - Single JSON objects
- Returns:
  - `fileName` — timestamped XLSX name  
  - `excelFile` — Base64 Excel file
- Zero field‑mapping required — API handles structure discovery automatically
- Ideal for automation, reconciliation workflows, and file-generation use cases
- Fully deployable on **Render.com**

---

# 📂 Project Structure

```
KeysieAPI/
  src/
    __init__.py
    main.py
  tests/
    test_convert.py
  sample/
    sample.json
  requirements.txt
  render.yaml
  .gitignore
  README.md
```

---

# 🔧 Local Installation

### 1. Clone the repository
```bash
git clone https://github.com/paulkeysza/KeysieAPI.git
cd KeysieAPI
```

### 2. Create and activate virtual environment
```bash
python -m venv .venv
.\.venv\Scriptsctivate        # Windows
source .venv/bin/activate       # Mac/Linux
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the API locally
```bash
uvicorn src.main:app --reload
```

### 5. Access the API
- Swagger UI → http://127.0.0.1:8000/docs  
- Health check → http://127.0.0.1:8000/health  

---

# 📬 API Endpoints

## **GET /health**
Checks API availability.

#### Response example:
```json
{ "status": "ok" }
```

---

## **POST /convert**

Converts JSON (string) into a Base64 Excel file.

#### Request body:
```json
{
  "jsonInput": "[{"name": "Paul", "amount": 123.45}]"
}
```

#### Response:
```json
{
  "fileName": "KeysieAPI-20250101123000.xlsx",
  "excelFile": "<BASE64_STRING>"
}
```

---

# 🧪 Testing

### Install PyTest
```bash
pip install pytest
```

### Run tests
```bash
pytest
```

Tests cover:
- Health endpoint  
- Basic JSON → Excel conversion  
- Response structure validation  

---

# ☁️ Render Deployment Guide

### Included file: `render.yaml`

Render will detect the file and auto‑configure the service.

### Build command:
```
pip install -r requirements.txt
```

### Start command:
```
uvicorn src.main:app --host 0.0.0.0 --port $PORT
```

### Setup steps:
1. Create **New Web Service**  
2. Connect GitHub  
3. Select `KeysieAPI`  
4. Runtime: **Python**  
5. Auto Deploy: **Yes**  
6. Deploy  

Your service becomes available at:

```
https://<your-service>.onrender.com
```

---

# 🤖 Nintex Xtension Integration

KeysieAPI is optimized for NAC:

### Input
| Field       | Type   | Description |
|-------------|--------|-------------|
| jsonInput   | string | Raw JSON input to convert |

### Output
| Field       | Type   | Description |
|-------------|--------|-------------|
| fileName    | string | Auto‑generated XLSX filename |
| excelFile   | string | Base64 encoded Excel file    |

### Example NAC Use Cases
- Excel file generation inside workflows  
- Bank statement reconciliation demos  
- Customer data extraction → Excel  
- Structured reporting automation  

I can generate a **ready‑to‑import Xtension JSON** on request.

---

# 🛠 Architecture Overview

```
Client / Workflow (Nintex, Postman, Custom App)
           |
           v
      KeysieAPI (FastAPI)
           |
    JSON parsing, validation
           |
  Dynamic Excel generation (OpenPyXL)
           |
      Base64 encoding
           |
           v
          Response
```

---

# 🎨 Sequence Diagram

```
Client
  │  POST /convert
  ▼
KeysieAPI
  │  Validate JSON
  │  Normalize list of dicts
  │  Create Excel workbook
  │  Encode workbook Base64
  ▼
Client
  │  Receives fileName + excelFile(Base64)
```

---

# 🔐 Security Notes

- JSON size should be monitored for abuse  
- Add API keys or OAuth if exposed publicly  
- CORS rules can be tightened if needed  
- Rate limiting recommended for production  

If you want, I can build:
- API Key middleware  
- JWT authentication  
- IP whitelisting  

---

# 📌 Future Enhancements (Optional Roadmap)

- Upload JSON file instead of string input  
- Return Excel file directly (binary MIME)  
- Field formatting support (currency/date)  
- Multiple sheet support  
- Nested JSON flattening  
- Metadata sheet generation  

---

# 👤 Author

**Paul**  
GitHub: https://github.com/paulkeysza

---

# ❤️ Contributing

Feel free to open issues or submit PRs!

---

# 📄 License

MIT License  
(Or specify another if needed)

