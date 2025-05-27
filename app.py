# app.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
import shutil
import os
import pandas as pd
import numpy as np
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
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if file.filename.endswith('.xlsx') or file.filename.endswith('.xls'):
            df = pd.read_excel(file_path, header=None)
        else:
            df = pd.read_csv(file_path, sep=';', encoding='latin1', header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {str(e)}"}, status_code=400)

    found_rows = []
    for idx, row in df.iterrows():
        if row.astype(str).str.contains('Quote D For distributor to quote to the reseller only', case=False, na=False).any():
            cleaned_row = row.replace({np.nan: None, np.inf: None, -np.inf: None})
            found_rows.append(cleaned_row)

    if not found_rows:
        return JSONResponse(content={"error": "No data found containing 'Quote D For distributor to quote to the reseller only'"}, status_code=404)

    print("\n===== Quote D Rows Found =====")
    preview_data = []
    for row in found_rows:
        row_list = row.to_list()
        print(row_list)
        preview_data.append(row_list)
    print("===== End of Found Rows =====\n")

    return JSONResponse(content={"message": "Quote D rows found and printed to console.", "preview": preview_data}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)

