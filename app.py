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
    print(df.head(50))  # show even more rows to locate headers
    print("===== Data Shape =====")
    print(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")
    print("===== End of Preview =====\n")

    # Find where the real header row (with column names) starts
    header_row_index = df[df.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index

    if header_row_index.empty:
        return JSONResponse(content={"error": "Could not locate the actual table header (Parent Quote Name row)."}, status_code=404)

    header_idx = header_row_index[0]
    data_start_idx = header_idx + 1

    # Read again using that row as header
    df_full = pd.read_excel(file_path, skiprows=data_start_idx, header=0)

    print("\n===== Parsed Table with Correct Headers =====")
    print(df_full.head())
    print("===== Columns =====")
    print(df_full.columns.tolist())
    print("===== End of Table Preview =====\n")

    return JSONResponse(content={"message": "Parsed Quote D table with correct headers. Check console for preview."}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)
