from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
from extractor.parser import extract_pdf_to_excel

app = FastAPI()

BASE_DIR = os.path.dirname(__file__)
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATE_DIR)

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload(request: Request, file: UploadFile = File(...)):
    """Handles PDF upload → Excel conversion → returns file."""
    pdf_path = f"/tmp/{file.filename}"
    out_path = pdf_path.replace(".pdf", ".xlsx")

    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    try:
        extract_pdf_to_excel(pdf_path, out_path)
    except Exception as e:
        return HTMLResponse(f"<pre>Processing error:\n{e}</pre>", status_code=500)

    return FileResponse(out_path, filename=os.path.basename(out_path))
