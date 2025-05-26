# app.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os
import pandas as pd
from openpyxl import load_workbook

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
TEMPLATE_PATH = "XQ-3352605.xlsx"
OUTPUT_PATH = "exported_quoteD.xlsx"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/process-quote-d")
async def process_quote_d(file: UploadFile = File(...)):
    csv_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(csv_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    df = pd.read_csv(csv_path, sep=';')

    if 'Quote' not in df.columns:
        return {"error": "'Quote' column not found in CSV"}

    quote_d_df = df[df['Quote'] == 'D']

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    start_row = 2
    for idx, row in quote_d_df.iterrows():
        ws.cell(row=start_row, column=1, value=row.get('Field1'))  # Replace with actual column names
        ws.cell(row=start_row, column=2, value=row.get('Field2'))
        start_row += 1

    wb.save(OUTPUT_PATH)

    return {"message": "Quote D data processed and exported.", "output_file": OUTPUT_PATH}

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return {"error": "No exported file found."}
