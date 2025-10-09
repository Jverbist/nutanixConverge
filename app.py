# app.py
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
import shutil
import os
import pandas as pd
from datetime import datetime
import calendar

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
OUTPUT_PATH = "exported_quoteD.csv"
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
    - USD/SEK/NOK/DKK: multiply by provided exchangeRate when computing prices
    """
    currency = currency.upper().strip()
    supported = {"EUR", "USD", "SEK", "NOK", "DKK"}
    if currency not in supported:
        return JSONResponse(
            status_code=400,
            content={"error": f"Unsupported currency '{currency}'. Supported: {sorted(supported)}"},
        )

    # Save uploaded file
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Read Excel and locate header row
    try:
        raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {e}"}, status_code=400)

    header_idx = raw[raw.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index
    if header_idx.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)

    header_idx = header_idx[0]
    data = raw.iloc[header_idx + 1:].reset_index(drop=True)
    data.columns = raw.iloc[header_idx].fillna("").tolist()

    filtered = data[data["Parent Quote Name"].astype(str).str.startswith("XQ-", na=False)]

    # Dates
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

    # Prepare CSV
    columns = [
        "ExternalId","Title","Currency","Date","Reseller","ResellerContact","Expires","ExpectedClose",
        "EndUser","BusinessUnit","Item","Quantity","Salesprice","Salesdiscount","Purchaseprice",
        "PurchaseDiscount","Location","ContractStart","ContractEnd","Serial#Supported","Rebate",
        "Opportunity","Memo (Line)","Quote ID (Line)","VendorSpecialPriceApproval",
        "VendorSpecialPriceApproval (Line)","SalesCurrency","SalesExchangeRate"
    ]
    rows = []

    reseller_clean = reseller.replace(" ", "_")

    for _, row in filtered.iterrows():
        # Discount
        purchase_disc = row.get("Total Discount (%)")
        if pd.isna(purchase_disc):
            purchase_disc = 0
        try:
            purchase_disc = float(str(purchase_disc).replace("%", "").replace(",", "").strip())
        except:
            purchase_disc = 0

        # List Price
        lp = row.get("List Price")
        if pd.isna(lp): lp = 0
        try:
            lp = float(str(lp).replace("$", "").replace(",", "").strip())
        except:
            lp = 0

        # Vendor Sale Price (Purchase Price)
        vendor_price = row.get("Sale Price")
        if pd.isna(vendor_price): vendor_price = 0
        try:
            vendor_price = float(str(vendor_price).replace("$", "").replace(",", "").strip())
        except:
            vendor_price = 0

        # --- New Pricing Logic ---
        fx = 1.0 if currency == "USD" else exchangeRate
        list_price_cur = lp * fx
        purchase_price = vendor_price * fx
        rate = purchase_price * (1 + (margin / 100))  # Sales price with margin

        if list_price_cur > 0:
            sales_disc_pct = max(0.0, min(100.0, 100.0 - (rate / list_price_cur * 100.0)))
        else:
            sales_disc_pct = 0.0

        # External ID and discount formatting
        ext_id = f"{reseller_clean}_{row.get('Parent Quote Name')}_{quote_date}"
        purchase_disc_str = f"{int(purchase_disc)}%"

        rows.append([
            ext_id, None, currency, quote_date, reseller, None, expires, None,
            None, "Belgium", str(row.get("Product Code")).strip(), row.get("Quantity"),
            round(rate, 2), f"{int(round(sales_disc_pct))}%", round(purchase_price, 2),
            purchase_disc_str, "Duffel : BE Sales Stock", None, None, None, None, None, None, None,
            row.get("Parent Quote Name"), None, currency, exchangeRate
        ])

    try:
        pd.DataFrame(rows, columns=columns).to_csv(OUTPUT_PATH, index=False)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to save output: {e}"}, status_code=500)

    return JSONResponse(content={"message": "Data exported successfully.", "output_file": "/download"}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename=os.path.basename(OUTPUT_PATH), media_type="text/csv")
    return JSONResponse(content={"error": "No exported file found."}, status_code=404)
