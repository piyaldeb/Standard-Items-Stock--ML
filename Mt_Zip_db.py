import requests
import pandas as pd
from datetime import date
from dotenv import load_dotenv
import os
# Load .env
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
import datetime

today = datetime.date.today()

# Always fixed at 01/01/2025
FROM_DATE = "2025-01-01" # Changed to YYYY-MM-DD format

# Today's date in YYYY-MM-DD
TO_DATE = today.strftime("%Y-%m-%d")             # Changed to YYYY-MM-DD format

session = requests.Session()
USER_ID = None


# ========= LOGIN ==========
def login():
    global USER_ID
    payload = {
        "jsonrpc": "2.0",
        "params": {
            "db": DB,
            "login": USERNAME,
            "password": PASSWORD
        }
    }
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        print(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
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
            "kwargs": {
                "context": {
                    "allowed_company_ids": [company_id],
                    "company_id": company_id
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        print(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    else:
        print(f"🔄 Session switched to company {company_id}")
        return True


# ========= CREATE FORECAST WIZARD ==========
def create_forecast_wizard(company_id, from_date, to_date):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "create",
            "args": [{
                "from_date": from_date,
                "to_date": to_date,
            }],
            "kwargs": {
                "context": {
                    "allowed_company_ids": [company_id],
                    "company_id": company_id,
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    wizard_id = r.json()["result"]
    print(f"🪄 Created wizard {wizard_id} for company {company_id}")
    return wizard_id


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
                    "company_id": company_id
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload)
    r.raise_for_status()
    result = r.json()
    if "error" in result:
        print(f"❌ Error computing forecast for {company_id}: {result['error']}")
    else:
        print(f"⚡ Forecast computed for wizard {wizard_id} (company {company_id})")
    return result


# ========= FETCH REPORT ==========
def fetch_opening_closing(company_id, cname):
    context = {
        "allowed_company_ids": [company_id],
        "company_id": company_id,
    }

    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.opening.closing",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "product_id": {"fields": {"display_name": {}}},
                    "product_category": {"fields": {"display_name": {}}},
                    "parent_category": {"fields": {"display_name": {}}},
                    "classification_id": {"fields": {"display_name": {}}},
                    "product_type": {"fields": {"display_name": {}}},
                    "item_category": {"fields": {"display_name": {}}},
                    "pr_code": {},
                    "product_uom": {"fields": {"display_name": {}}},
                    "lot_id": {"fields": {"display_name": {}}},
                    "opening_qty": {},
                    "opening_value": {},
                    "receive_qty": {},
                    "receive_value": {},
                    "issue_qty": {},
                    "issue_value": {},
                    "cloing_qty": {},
                    "cloing_value": {},
                },
                "offset": 0,
                "limit": 5000,
                "context": {
                    **context,
                    "active_model": "stock.forecast.report",
                    "active_id": 0,
                    "active_ids": [0],
                },
                "count_limit": 10000,
                "domain": [["product_id.categ_id.complete_name", "ilike", "All / RM"]],
            },
        },
    }

    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]["records"]
        print(f"📊 {cname}: {len(data)} rows fetched")
        return data
    except Exception:
        print(f"❌ {cname}: Failed to parse report:", r.text[:200])
        return []


# ========= MAIN ==========
if __name__ == "__main__":
    userinfo = login()
    print("User info (allowed companies):", userinfo.get("user_companies", {}))

    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            # Step 1: Create wizard
            wiz_id = create_forecast_wizard(cid, FROM_DATE, TO_DATE)

            # Step 2: Compute forecast
            compute_forecast(cid, wiz_id)

            # Step 3: Fetch report
            records = fetch_opening_closing(cid, cname)

            if records:
                df = pd.DataFrame(records)
                output_file = f"{cname.lower().replace(' ', '')}_opening_closing{today.isoformat()}.xlsx"
                df.to_excel(output_file, index=False)
                print(f"📂 Saved: {output_file}")
            else:
                print(f"❌ No data fetched for {cname}")