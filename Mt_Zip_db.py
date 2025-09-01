import requests
import pandas as pd
from datetime import date, datetime
from dotenv import load_dotenv
import os
import pytz
import logging
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe

# === Load .env ===
load_dotenv()

# ========= CONFIG ==========
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

COMPANIES = {
    1: "Zipper",
    3: "Metal Trims",
}

FROM_DATE = "2025-01-01"
TO_DATE = date.today().strftime("%Y-%m-%d")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# === Logging ===
logging.basicConfig(level=logging.INFO)
log = logging.getLogger()

# === Session ===
session = requests.Session()
USER_ID = None

# ========= LOGIN ==========
def login():
    global USER_ID
    payload = {
        "jsonrpc": "2.0",
        "params": {"db": DB, "login": USERNAME, "password": PASSWORD}
    }
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        log.info(f"✅ Logged in (uid={USER_ID})")
        return result
    raise Exception("❌ Login failed")

# ========= SWITCH COMPANY ==========
def switch_company(company_id):
    if USER_ID is None:
        raise Exception("User not logged in yet")
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "res.users",
            "method": "write",
            "args": [[USER_ID], {"company_id": company_id}],
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}},
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        log.error(f"❌ Failed to switch company {company_id}: {r.json()['error']}")
        return False
    log.info(f"🔄 Switched to company {company_id}")
    return True

# ========= CREATE FORECAST WIZARD ==========
def create_forecast_wizard(company_id):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "create",
            "args": [{"from_date": FROM_DATE, "to_date": TO_DATE}],
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}},
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    wiz_id = r.json()["result"]
    log.info(f"🪄 Created wizard {wiz_id} for company {company_id}")
    return wiz_id

# ========= COMPUTE FORECAST ==========
def compute_forecast(company_id, wizard_id):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "print_date_wise_stock_register",
            "args": [[wizard_id]],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": "Asia/Dhaka",
                    "uid": USER_ID,
                    "allowed_company_ids": [company_id],
                    "company_id": company_id,
                }
            },
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload)
    r.raise_for_status()
    log.info(f"⚡ Forecast computed for wizard {wizard_id} (company {company_id})")
    return r.json()

# ========= FETCH OPENING/CLOSING ========== 
def fetch_opening_closing(company_id, cname):
    context = {"allowed_company_ids": [company_id], "company_id": company_id}
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.opening.closing",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "product_id": {"fields": {"display_name": {}}},        # Product
                    "product_category": {"fields": {"display_name": {}}},  # Category
                    "parent_category": {"fields": {"display_name": {}}},   # Item
                    "pr_code": {},                                         # Item Code
                    "lot_id": {"fields": {"display_name": {}}},            # Invoice
                    "receive_date": {},                                     # Receive Date
                    "pur_price": {},                                        # Pur Price
                    "landed_cost": {},                                      # Landed Cost
                    "lot_price": {},                                        # Price
                    "product_uom": {"fields": {"display_name": {}}},       # Unit
                    "opening_qty": {},                                      # Opening Quantity
                    "opening_value": {},                                    # Opening Value
                    "receive_qty": {},                                      # Receive Quantity
                    "receive_value": {},                                    # Receive Value
                    "issue_qty": {},                                        # Issue Quantity
                    "issue_value": {},                                      # Issue Value
                    "cloing_qty": {},                                       # Closing Quantity
                    "cloing_value": {},                                     # Closing Value
                    "po_type": {},                                          # Po Type
                    "rejected": {},                                         # Rejected
                    "shipment_mode": {},                                    # Shipment Mode
                },
                "offset": 0,
                "limit": 5000,
                "context": {**context, "active_model": "stock.forecast.report", "active_id": 0, "active_ids": [0]},
                "count_limit": 10000,
                "domain": [["product_id.categ_id.complete_name", "ilike", "All / RM"]],
            },
        },
    }

    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()

    try:
        data = r.json()["result"]["records"]

        # Flatten nested dicts → keep only display_name
        def flatten_record(record):
            flat = {}
            for k, v in record.items():
                if isinstance(v, dict) and "display_name" in v:
                    flat[k] = v["display_name"]
                else:
                    flat[k] = v
            return flat

        flattened = [flatten_record(rec) for rec in data]
        log.info(f"📊 {cname}: {len(flattened)} rows fetched (flattened)")
        return pd.DataFrame(flattened)

    except Exception:
        log.error(f"❌ {cname}: Failed to parse report: {r.text[:200]}")
        return pd.DataFrame()



# ========= PASTE TO GOOGLE SHEETS ==========
def paste_to_google_sheet(df, sheet_key, worksheet_name):
    if df.empty:
        log.warning("DataFrame empty. Skipping Google Sheet update.")
        return
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file("service_account.json", scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_key)
    worksheet = sheet.worksheet(worksheet_name)
    worksheet.clear()
    set_with_dataframe(worksheet, df)
    tz = pytz.timezone("Asia/Dhaka")
    timestamp = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    # Write timestamp in last column safely
    last_col = chr(65 + min(25, df.shape[1]))  # A-Z, max 26 columns
    worksheet.update(f"{last_col}2", [[timestamp]])
    log.info(f"✅ Data pasted to {worksheet_name} & timestamp updated: {timestamp}")

# ========= MAIN SYNC ==========
if __name__ == "__main__":
    login()
    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            wiz_id = create_forecast_wizard(cid)
            compute_forecast(cid, wiz_id)
            df = fetch_opening_closing(cid, cname)
            if not df.empty:
                # Save locally
                local_file = os.path.join(DOWNLOAD_DIR, f"{cname.lower().replace(' ', '')}_opening_closing_{TO_DATE}.xlsx")
                df.to_excel(local_file, index=False)
                log.info(f"📂 Saved locally: {local_file}")

                # Worksheet name based on company
                worksheet_name = "Zipper" if cid == 1 else "Metal" if cid == 3 else cname

                # Paste to Google Sheets
                paste_to_google_sheet(df, sheet_key="1tSgmESOWqYRkDk_KhewnaaJmQGSUaSILzOzpade9tRc", worksheet_name=worksheet_name)
