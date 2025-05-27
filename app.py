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

    print("\n===== Raw Data Preview =====")
    print(df.head(30))  # show more to locate useful rows
    print("===== Data Shape =====")
    print(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")
    print("===== End of Preview =====\n")

    # Find row where real table header starts (after 'Quote D' marker)
    quote_d_row = df[df.apply(lambda r: r.astype(str).str.contains('Quote D', case=False, na=False).any(), axis=1)].index

    if quote_d_row.empty:
        return JSONResponse(content={"error": "No 'Quote D' marker found."}, status_code=404)

    header_start_idx = quote_d_row[0] + 15  # skip down ~15 lines to table header (adjust if needed)
    table_df = pd.read_excel(file_path, skiprows=header_start_idx)

    print("\n===== Extracted Table Preview =====")
    print(table_df.head())
    print("===== Table Columns =====")
    print(table_df.columns.tolist())
    print("===== End of Table Preview =====\n")

    return JSONResponse(content={"message": "Extracted Quote D table, preview printed to console."}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)

