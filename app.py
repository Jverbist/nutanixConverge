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

    header_row_index = df[df.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index

    if header_row_index.empty:
        return JSONResponse(content={"error": "Could not locate the actual table header (Parent Quote Name row)."}, status_code=404)

    header_idx = header_row_index[0]
    data_start_idx = header_idx + 1

    df_full = pd.read_excel(file_path, skiprows=data_start_idx, header=0)

    if 'Parent Quote Name' not in df_full.columns:
        return JSONResponse(content={"error": "'Parent Quote Name' column not found in data."}, status_code=400)

    combined_data = {}
    for _, row in df_full.iterrows():
        key = str(row['Parent Quote Name']).strip()
        row_data = row.drop(labels=[col for col in ['Parent Quote Name'] if col in row]).to_dict()
        if key not in combined_data:
            combined_data[key] = []
        combined_data[key].append(row_data)

    print("\n===== Combined Data Preview =====")
    for k, v in combined_data.items():
        print(f"Key: {k}")
        for item in v:
            print(item)
    print("===== End of Combined Data =====\n")

    return JSONResponse(content={"message": "Combined data by Parent Quote Name created. Check console for preview."}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)

