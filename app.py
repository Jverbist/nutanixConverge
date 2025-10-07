# app.py
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
import shutil
import os
import pandas as pd
import csv
from datetime import datetime
import calendar

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
OUTPUT_PATH = "exported_quoteD.csv"
# Ensure upload directory exists
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("index.html", "r") as file:
        return HTMLResponse(content=file.read())

@app.post("/process-quote-d")
async def process_quote_d(
    file: UploadFile = File(...),
    reseller: str = Form(...),
    currency: str = Form(...),  # EUR, USD, SEK, NOK, DKK
    exchangeRate: float = Form(...),
    margin: float = Form(...)
):
    """
    currency behavior:
    - EUR: treat prices as EUR; do not multiply by exchangeRate for base calc
    - USD/SEK/NOK/DKK: multiply by provided exchangeRate when computing sales price
    """
    currency = currency.upper().strip()
    supported = {"EUR", "USD", "SEK", "NOK", "DKK"}
    if currency not in supported:
        return JSONResponse(
            status_code=400,
            content={"error": f"Unsupported currency '{currency}'. Supported: {sorted(supported)}"},
        )

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
    header_idx = raw[raw.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index
    if header_idx.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)
    header_idx = header_idx[0]
    data = raw.iloc[header_idx + 1 :].reset_index(drop=True)
    data.columns = raw.iloc[header_idx].fillna("").tolist()

    # Filter only XQ- quotes
    filtered = data[data["Parent Quote Name"].astype(str).str.startswith("XQ-", na=False)]

    # Prepare dates in DD/MM/YYYY
    today = datetime.today()
    quote_date = today.strftime("%d/%m/%Y")
    if currency == "USD":
        last_day = calendar.monthrange(today.year, today.month)[1]
        expires = today.replace(day=last_day).strftime("%d/%m/%Y")
    else:
        d = today.day
        if d <= 10:
            expires = today.replace(day=10).strftime("%d/%m/%Y")
        elif d <= 20:
            expires = today.replace(day=20).strftime("%d/%m/%Y")
        else:
            last_day = calendar.monthrange(today.year, today.month)[1]
            expires = today.replace(day=last_day).strftime("%d/%m/%Y")

    # Prepare CSV header and rows
    columns = [
        "ExternalId","Title","Currency","Date","Reseller","ResellerContact","Expires","ExpectedClose",
        "EndUser","BusinessUnit","Item","Quantity","Salesprice","Salesdiscount","Purchaseprice",
        "PurchaseDiscount","Location","ContractStart","ContractEnd","Serial#Supported","Rebate",
        "Opportunity","Memo (Line)","Quote ID (Line)","VendorSpecialPriceApproval",
        "VendorSpecialPriceApproval (Line)","SalesCurrency","SalesExchangeRate"
    ]
    rows = []

    # Clean reseller for ExternalId (replace spaces with underscore)
    reseller_clean = reseller.replace(" ", "_")


    for _, row in filtered.iterrows():
        # Parse discount (Total Discount (%))
        disc = row.get("Total Discount (%)")
        if pd.isna(disc):
            disc = 0
        try:
            disc = float(str(disc).replace("%", "").replace(",", "").strip())
        except:
            disc = 0

        # Parse list price
        lp = row.get("List Price")
        if pd.isna(lp):
            lp = 0
        try:
            lp = float(str(lp).replace("$", "").replace(",", "").strip())
        except:
            lp = 0

        # Parse vendor sale price (used for Purchaseprice)
        vendor_price = row.get("Sale Price")
        if pd.isna(vendor_price):
            vendor_price = 0
        try:
            vendor_price = float(str(vendor_price).replace("$", "").replace(",", "").strip())
        except:
            vendor_price = 0

        # Net price after discount (used as our base before markup)
        net_price = lp * (1 - disc / 100)

        # Purchase price equals the vendor's Sale Price
        purchase_price = vendor_price

        # Determine base sales price (EUR/SEK/NOK/DKK multiply by exchangeRate; USD keep as-is)
        code = str(row.get("Product Code")).strip()
        if code.startswith("NX"):
            base_native = net_price * 2
        else:
            base_native = net_price

        if currency == "EUR":
            base = base_native
        else:
            base = base_native * exchangeRate

        # Apply margin markup
        sales_price = base * (1 + margin / 100)

        # Sales discount relative to net
        sales_disc = round(1 - (net_price / sales_price), 2) if sales_price > 0 else 0

        # ExternalId: reseller_clean + quote + date
        ext_id = f"{reseller_clean}_{row.get('Parent Quote Name')}_{quote_date}"

        # Purchase discount string
        purchase_disc_str = f"{int(disc)}%"

        # Append row
        rows.append([
            ext_id, None, currency, quote_date, reseller, None, expires, None,
            None, "Belgium", code, row.get("Quantity"), round(sales_price, 2),
            f"{int(sales_disc * 100)}%", round(purchase_price, 2), purchase_disc_str,
            "Duffel : BE Sales Stock", None, None, None, None, None, None, None,
            row.get("Parent Quote Name"), None, currency, exchangeRate
        ])

    # Save CSV
    try:
        pd.DataFrame(rows, columns=columns).to_csv(OUTPUT_PATH, index=False)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to save output: {e}"}, status_code=500)

    return JSONResponse(content={"message": "Data exported successfully.", "output_file": "/download"}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename=os.path.basename(OUTPUT_PATH), media_type='text/csv')
    return JSONResponse(content={"error": "No exported file found."}, status_code=404)

