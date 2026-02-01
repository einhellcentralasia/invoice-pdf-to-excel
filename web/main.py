from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
import os
import json
from extractor.parser import extract_pdf_to_excel, extract_pdf_rows, is_blocked_header

app = FastAPI()

allowed_origins = []
cors_origin = os.getenv("CORS_ORIGIN")
if cors_origin:
    allowed_origins = [o.strip() for o in cors_origin.split(",") if o.strip()]
if allowed_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=allowed_origins,
        allow_credentials=True,
        allow_methods=["POST", "GET", "OPTIONS"],
        allow_headers=["*"],
    )

BASE_DIR = os.path.dirname(__file__)
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATE_DIR)

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/health")
async def health():
    return {"ok": True}

@app.post("/preview")
async def preview(file: UploadFile = File(...)):
    """Extract headers + sample rows for UI preview."""
    safe_name = os.path.basename(file.filename or "upload.pdf")
    if not safe_name.lower().endswith(".pdf"):
        safe_name = f"{safe_name}.pdf"
    pdf_path = f"/tmp/{safe_name}"

    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    try:
        rows, headers = extract_pdf_rows(pdf_path)
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

    blocked = [h for h in headers if is_blocked_header(h)]
    sample = rows[:3]
    return {"ok": True, "headers": headers, "blocked": blocked, "sample": sample, "always": ["Page"]}

@app.post("/upload")
async def upload(request: Request, file: UploadFile = File(...), columns: str = Form(None)):
    """Handles PDF upload → Excel conversion → returns file."""
    safe_name = os.path.basename(file.filename or "upload.pdf")
    if not safe_name.lower().endswith(".pdf"):
        safe_name = f"{safe_name}.pdf"
    pdf_path = f"/tmp/{safe_name}"
    out_path = pdf_path.replace(".pdf", ".xlsx")

    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    selected_columns = []
    if columns:
        try:
            selected_columns = json.loads(columns)
            if not isinstance(selected_columns, list):
                selected_columns = []
        except Exception:
            selected_columns = [c.strip() for c in columns.split(",") if c.strip()]

    # filter out blocked headers on the server side
    selected_columns = [c for c in selected_columns if not is_blocked_header(c)]
    if columns is not None and not selected_columns:
        return JSONResponse({"ok": False, "error": "NO_COLUMNS"}, status_code=400)

    try:
        extract_pdf_to_excel(pdf_path, out_path, selected_columns=selected_columns)
    except Exception as e:
        return HTMLResponse(f"<pre>Processing error:\n{e}</pre>", status_code=500)

    return FileResponse(out_path, filename=os.path.basename(out_path))
