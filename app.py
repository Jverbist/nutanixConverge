from fastapi import FastAPI, UploadFile, File
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
    # Save the uploaded CSV
    csv_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(csv_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Load CSV with pandas (semicolon-delimited)
    df = pd.read_csv(csv_path, sep=';')

    # Filter rows where 'Quote' (or equivalent field) == 'D'
    if 'Quote' not in df.columns:
        return {"error": "'Quote' column not found in CSV"}

    quote_d_df = df[df['Quote'] == 'D']

    # Load the Excel template
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active  # Or specify by name if needed

    # Write filtered data into Excel (example: starting at row 2)
    start_row = 2
    for idx, row in quote_d_df.iterrows():
        ws.cell(row=start_row, column=1, value=row.get('Field1'))  # replace 'Field1' with actual column name
        ws.cell(row=start_row, column=2, value=row.get('Field2'))
        # ... add as needed
        start_row += 1

    # Save the output Excel
    wb.save(OUTPUT_PATH)

    return {"message": "Quote D data processed and exported.", "output_file": OUTPUT_PATH}

