# app.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
import shutil
import os
import pandas as pd
import numpy as np

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
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
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {str(e)}"}, status_code=400)

    # Locate the row with 'Parent Quote Name'
    header_row_index = df[df.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index

    if header_row_index.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)

    header_idx = header_row_index[0]
    data_start_idx = header_idx + 1

    # Extract headers and data
    header_row = df.iloc[header_idx].fillna('').tolist()
    data_rows = df.iloc[data_start_idx:].reset_index(drop=True)
    data_rows.columns = header_row

    # Filter only valid quote rows (exclude TOTAL, disclaimers, footnotes)
    filtered_rows = data_rows[data_rows['Parent Quote Name'].astype(str).str.startswith('XQ-', na=False)]

    combined_data = {}
    for _, row in filtered_rows.iterrows():
        key = str(row['Parent Quote Name']).strip()
        row_data = row.drop(labels=['Parent Quote Name']).replace({np.nan: None}).to_dict()
        if key not in combined_data:
            combined_data[key] = []
        combined_data[key].append(row_data)

    print("\n===== Combined Filtered Data Preview =====")
    for k, v in combined_data.items():
        print(f"Key: {k}")
        for item in v:
            print(item)
    print("===== End of Filtered Data =====\n")

    return JSONResponse(content={"message": "Filtered and combined data by Parent Quote Name created. Check console for preview."}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)

