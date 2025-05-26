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
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if file.filename.endswith('.xlsx') or file.filename.endswith('.xls'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path, sep=';', encoding='latin1')
    except Exception as e:
        return {"error": f"Failed to read file: {str(e)}"}

    if 'Quote' not in df.columns:
        return {"error": "'Quote' column not found in the provided file"}

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
