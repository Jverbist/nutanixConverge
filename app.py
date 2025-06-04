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
            if code and name and code != 'nan' and name != 'nan':
                combined = f"{code} {name}"
                resellers.append(combined)
        return sorted(resellers)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to load resellers: {str(e)}"}, status_code=500)

@app.post("/process-quote-d")
async def process_quote_d(
    file: UploadFile = File(...),
    reseller: str = Form(...),
    currency: str = Form(...),
    exchangeRate: float = Form(...),
    margin: float = Form(...)
):
    print(f"Received currency: {currency}, exchangeRate: {exchangeRate}, margin: {margin}")

    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to read file: {str(e)}"}, status_code=400)

    header_row_index = df[df.apply(lambda r: r.astype(str).str.contains('Parent Quote Name', case=False, na=False).any(), axis=1)].index
    if header_row_index.empty:
        return JSONResponse(content={"error": "Could not locate 'Parent Quote Name' header row."}, status_code=404)

    header_idx = header_row_index[0]
    data_start_idx = header_idx + 1
    header_row = df.iloc[header_idx].fillna('').tolist()
    data_rows = df.iloc[data_start_idx:].reset_index(drop=True)
    data_rows.columns = header_row
    filtered_rows = data_rows[data_rows['Parent Quote Name'].astype(str).str.startswith('XQ-', na=False)]

    today = datetime.today()
    today_str = today.strftime('%Y-%m-%d')

    if currency.upper() == 'USD':
        last_day = calendar.monthrange(today.year, today.month)[1]
        expires_date = today.replace(day=last_day).strftime('%Y-%m-%d')
    else:  # EUR
        day = today.day
        if day <= 10:
            expires_date = today.replace(day=10).strftime('%Y-%m-%d')
        elif day <= 20:
            expires_date = today.replace(day=20).strftime('%Y-%m-%d')
        else:
            last_day = calendar.monthrange(today.year, today.month)[1]
            expires_date = today.replace(day=last_day).strftime('%Y-%m-%d')

    wb = Workbook()
    ws = wb.active
    ws.append([
        "ExternalId", "Title", "Currency", "Date", "Reseller", "ResellerContact", "Expires", "ExpectedClose",
        "EndUser", "BusinessUnit", "Item", "Quantity", "Salesprice", "Salesdiscount", "Purchaseprice",
        "PurchaseDiscount", "Location", "ContractStart", "ContractEnd", "Serial#Supported", "Rebate",
        "Opportunity", "Memo (Line)", "Quote ID (Line)", "VendorSpecialPriceApproval",
        "VendorSpecialPriceApproval (Line)", "SalesCurrency", "SalesExchangeRate"
    ])

    for _, row in filtered_rows.iterrows():
        product_code = str(row.get('Product Code')).strip()

        discount = row.get('Total Discount (%)')
        if pd.isna(discount):
            discount = 0
        try:
            discount = float(str(discount).replace('%', '').replace(',', '').strip())
        except:
            discount = 0

        list_price = row.get('List Price')
        if pd.isna(list_price):
            list_price = 0
        try:
            list_price = float(str(list_price).replace('$', '').replace(',', '').strip())
        except:
            list_price = 0

        sale_price_quote = row.get('Sale Price')
        if pd.isna(sale_price_quote):
            sale_price_quote = 0
        try:
            sale_price_quote = float(str(sale_price_quote).replace('$', '').replace(',', '').strip())
        except:
            sale_price_quote = 0

        netto = list_price * (1 - discount / 100)
        purchase_price = netto
        sales_price = sale_price_quote

        if sales_price > 0:
            sales_discount = round(1 - (netto / sales_price), 2)
        else:
            sales_discount = 0

        external_id = f"{reseller}_{row.get('Parent Quote Name')}_{today_str}"

        print(f"Processing row: Product Code: {product_code}, List Price: {list_price}, Discount: {discount}, Sale Price: {sales_price}, Netto: {netto}, Sales Discount: {sales_discount}")

        ws.append([
            external_id,  # ExternalId
            None,  # Title
            currency,  # Currency
            today_str,  # Date
            reseller,  # Reseller
            None,  # ResellerContact
            expires_date,  # Expires
            None,  # ExpectedClose
            None,  # EndUser
            "Belgium",  # BusinessUnit
            product_code,  # Item
            row.get('Quantity'),  # Quantity
            round(sales_price, 2),  # Salesprice
            f"{int(sales_discount * 100)}%",  # Salesdiscount
            round(purchase_price, 2),  # Purchaseprice
            f"{int(discount)}%",  # PurchaseDiscount
            "Duffel : BE Sales Stock",  # Location
            None,  # ContractStart
            None,  # ContractEnd
            None,  # Serial#Supported
            None,  # Rebate
            None,  # Opportunity
            None,  # Memo (Line)
            None,  # Quote ID (Line)
            row.get('Parent Quote Name'),  # VendorSpecialPriceApproval
            None,  # VendorSpecialPriceApproval (Line)
            currency,  # SalesCurrency
            exchangeRate  # SalesExchangeRate
        ])

    try:
        wb.save(OUTPUT_PATH)
    except Exception as e:
        return JSONResponse(content={"error": f"Failed to save output file: {str(e)}"}, status_code=500)

    return JSONResponse(content={"message": "Data exported successfully.", "output_file": "/download"}, status_code=200)

@app.get("/download")
async def download_file():
    if os.path.exists(OUTPUT_PATH):
        return FileResponse(OUTPUT_PATH, filename="exported_quoteD.xlsx")
    else:
        return JSONResponse(content={"error": "No exported file found."}, status_code=404)

