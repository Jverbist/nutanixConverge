from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
import shutil
import os
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import calendar

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
OUTPUT_PATH = "exported_quoteD.xlsx"
RESELLER_FILE = "Reseller.xlsx"
# Ensure upload directory exists
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("index.html", "r") as file:
        return HTMLResponse(content=file.read())

@app.get("/resellers")
async def get_resellers():
    try:
        df = pd.read_excel(RESELLER_FILE, header=None)
        resellers = []
        for _, row in df.iterrows():
            code = str(row[0]).strip()
            name = str(row[1]).strip() if len(row) > 1 else ""
            if code and name and code.lower() != 'nan' and name.lower() != 'nan':
                resellers.append(f"{code} {name}")
        return sorted(resellers)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to load resellers: {e}"}, status_code=500)

@app.post("/process-quote-d")
async def process_quote_d(
    file: UploadFile = File(...),
    reseller: str = Form(...),
    currency: str = Form(...),
    exchangeRate: float = Form(...),
    margin: float = Form(...)
):
    # Save uploaded file to disk
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Read raw Excel with unknown header row
    try:
        raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {e}"}, status_code=400)

    # Find header row containing 'Parent Quote Name'
    header_idx = raw[raw.apply(lambda row: row.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index
    if header_idx.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)
    header_idx = header_idx[0]
    data = raw.iloc[header_idx+1:].reset_index(drop=True)
    data.columns = raw.iloc[header_idx].fillna('').tolist()
    # Filter only XQ- quotes
    filtered = data[data['Parent Quote Name'].astype(str).str.startswith('XQ-', na=False)]

    # Prepare dates in DD/MM/YYYY
    today = datetime.today()
    quote_date = today.strftime('%d/%m/%Y')
    if currency.upper() == 'USD':
        last_day = calendar.monthrange(today.year, today.month)[1]
        expires = today.replace(day=last_day).strftime('%d/%m/%Y')
    else:
        d = today.day
        if d <= 10:
            expires = today.replace(day=10).strftime('%d/%m/%Y')
        elif d <= 20:
            expires = today.replace(day=20).strftime('%d/%m/%Y')
        else:
            last_day = calendar.monthrange(today.year, today.month)[1]
            expires = today.replace(day=last_day).strftime('%d/%m/%Y')

    # Create workbook and header
    wb = Workbook()
    ws = wb.active
    ws.append([
        "ExternalId","Title","Currency","Date","Reseller","ResellerContact","Expires","ExpectedClose",
        "EndUser","BusinessUnit","Item","Quantity","Salesprice","Salesdiscount","Purchaseprice",
        "PurchaseDiscount","Location","ContractStart","ContractEnd","Serial#Supported","Rebate",
        "Opportunity","Memo (Line)","Quote ID (Line)","VendorSpecialPriceApproval",
        "VendorSpecialPriceApproval (Line)","SalesCurrency","SalesExchangeRate"
    ])

    # Clean reseller for ExternalId (remove spaces)
    reseller_clean = reseller.replace(' ', '_')

    for _, row in filtered.iterrows():
        # Parse discount
        disc = row.get('Total Discount (%)')
        if pd.isna(disc):
            disc = 0
        try:
            disc = float(str(disc).replace('%','').replace(',','').strip())
        except:
            disc = 0

        # Parse list price
        lp = row.get('List Price')
        if pd.isna(lp): lp = 0
        try:
            lp = float(str(lp).replace('$','').replace(',','').strip())
        except:
            lp = 0

                # Parse vendor sale price
        vendor_price = row.get('Sale Price')
        if pd.isna(vendor_price):
            vendor_price = 0
        try:
            vendor_price = float(str(vendor_price).replace('$','').replace(',','').strip())
        except:
            vendor_price = 0

                # Calculate net price after discount
        net_price = lp * (1 - disc/100)
        # Purchase price equals vendor sale price
        purchase_price = vendor_price

        # Determine base sales price
        code = str(row.get('Product Code')).strip()
        if code.startswith('NX'):
            base = net_price * (2 * exchangeRate if currency.upper()=='EUR' else 2)
        else:
            base = net_price * (exchangeRate if currency.upper()=='EUR' else 1)

        # Apply margin
        sales_price = base * (1 + margin/100)

        # Sales discount relative to net
        sales_disc = round(1 - (net_price/sales_price),2) if sales_price>0 else 0

        # ExternalId: reseller_clean + quote + date
        ext_id = f"{reseller_clean}_{row.get('Parent Quote Name')}_{quote_date}"

        # Purchase discount string
        purchase_disc_str = f"{int(disc)}%"

        # Append row
        ws.append([
            ext_id, None, currency, quote_date, reseller, None, expires, None,
            None, 'Belgium', code, row.get('Quantity'), round(sales_price,2),
            f"{int(sales_disc*100)}%", round(purchase_price,2), purchase_disc_str,
            'Duffel : BE Sales Stock', None, None, None, None, None, None, None,
            row.get('Parent Quote Name'), None, currency, exchangeRate
        ])

    # Save workbook
    try:
        wb.save(OUTPUT_PATH)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to save output: {e}"}, status_code=500)

    return JSONResponse(content={"message":"Data exported successfully.","output_file":"/download"}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename=os.path.basename(OUTPUT_PATH))
    return JSONResponse(content={"error":"No exported file found."}, status_code=404)

