# app.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
import shutil
import os
import pandas as pd
from openpyxl import load_workbook

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
TEMPLATE_PATH = "XQ-3352605.xlsx"
OUTPUT_PATH = "exported_quoteD.xlsx"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("index.html", "r") as file:
        html_content = file.read()
    return HTMLResponse(content=html_content)

@app.post("/process-quote-d")
async def process_quote_d(file: UploadFile = File(...)):
    csv_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(csv_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        df = pd.read_csv(csv_path, sep=';', encoding='latin1')  # fallback encoding
    except UnicodeDecodeError:
        return {"error": "Could not decode file. Please ensure it is a valid CSV with correct encoding."}

    if 'Quote' not in df.columns:
        return {"error": "'Quote' column not found in CSV"}

    quote_d_df = df[df['Quote'] == 'D']

    if quote_d_df.empty:
        return {"error": "No data found for Quote D"}

    print("\n===== Quote D Data Preview =====")
    print(quote_d_df.head())
    print("===== End of Preview =====\n")

    return {"message": "Quote D data loaded successfully and printed to console."}

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return {"error": "No exported file found."}


