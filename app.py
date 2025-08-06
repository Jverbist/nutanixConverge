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
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("index.html", "r") as file:
        html_content = file.read()
    return HTMLResponse(content=html_content)

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
    enduser: str = Form(...),
    currency: str = Form(...),
    exchangeRate: float = Form(...),
    margin: float = Form(...)
):
    # Save uploaded file
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Read Excel
    try:
        raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {e}"}, status_code=400)

    # Locate header row
    header_idx = raw[raw.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index
    if header_idx.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)
    header_idx = header_idx[0]
    data = raw.iloc[header_idx+1:].reset_index(drop=True)
    data.columns = raw.iloc[header_idx].fillna('').tolist()
    filtered = data[data['Parent Quote Name'].astype(str).str.startswith('XQ-', na=False)]

    # Current dates
    today = datetime.today()
    today_str = today.strftime('%d/%m/%Y')  # DD/MM/YYYY
    if currency.upper() == 'USD':
        last_day = calendar.monthrange(today.year, today.month)[1]
        expires_date = today.replace(day=last_day).strftime('%d/%m/%Y')
    else:
        day = today.day
        if day <= 10:
            expires_date = today.replace(day=10).strftime('%d/%m/%Y')
        elif day <= 20:
            expires_date = today.replace(day=20).strftime('%d/%m/%Y')
        else:
            last_day = calendar.monthrange(today.year, today.month)[1]
            expires_date = today.replace(day=last_day).strftime('%d/%m/%Y')

    # Prepare workbook
    wb = Workbook()
    ws = wb.active
    ws.append([
        "ExternalId", "Title", "Currency", "Date", "Reseller", "ResellerContact", "Expires", "ExpectedClose",
        "EndUser", "BusinessUnit", "Item", "Quantity", "Salesprice", "Salesdiscount", "Purchaseprice",
        "PurchaseDiscount", "Location", "ContractStart", "ContractEnd", "Serial#Supported", "Rebate",
        "Opportunity", "Memo (Line)", "Quote ID (Line)", "VendorSpecialPriceApproval",
        "VendorSpecialPriceApproval (Line)", "SalesCurrency", "SalesExchangeRate"
    ])

    # Clean enduser for ExternalId (no spaces)
    enduser_clean = enduser.replace(' ', '')

    for _, row in filtered.iterrows():
        # Discount parsing
        discount = row.get('Total Discount (%)')
        if pd.isna(discount): discount = 0
        try:
            discount = float(str(discount).replace('%', '').replace(',', '').strip())
        except:
            discount = 0

        # List price parsing
        list_price = row.get('List Price')
        if pd.isna(list_price): list_price = 0
        try:
            list_price = float(str(list_price).replace('$', '').replace(',', '').strip())
        except:
            list_price = 0

        # Net price after discount
        net_price = list_price * (1 - discount / 100)
        # Purchase price is list price
        purchase_price = list_price

        # Base sales price calculation
        if str(row.get('Product Code')).strip().startswith("NX"):
            base_price = net_price * (2 * exchangeRate if currency.upper() == 'EUR' else 2)
        else:
            base_price = net_price * (exchangeRate if currency.upper() == 'EUR' else 1)

        # Apply margin markup
        sales_price = base_price * (1 + margin/100)

        # Sales discount relative to list price
        sales_discount = round(1 - (net_price / sales_price), 2) if sales_price > 0 else 0

        # ExternalId: EndUser + ParentQuote + date
        external_id = f"{enduser_clean}_{row.get('Parent Quote Name')}_{today_str}"

        # Format purchase discount string
        purchase_discount_str = f"{int(discount)}%"

        ws.append([
            external_id,
            None,
            currency,
            today_str,
            reseller,
            None,
            expires_date,
            None,
            enduser,  # EndUser
            "Belgium",
            str(row.get('Product Code')).strip(),
            row.get('Quantity'),
            round(sales_price, 2),
            f"{int(sales_discount * 100)}%",
            round(purchase_price, 2),
            purchase_discount_str,
            "Duffel : BE Sales Stock",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            row.get('Parent Quote Name'),
            None,
            currency,
            exchangeRate
        ])

    # Save workbook
    try:
        wb.save(OUTPUT_PATH)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to save output file: {e}"}, status_code=500)

    return JSONResponse(content={"message": "Data exported successfully.", "output_file": "/download"}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename=os.path.basename(OUTPUT_PATH))
    return JSONResponse(content={"error": "No exported file found."}, status_code=404)

