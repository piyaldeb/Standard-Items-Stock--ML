import sys
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import re
from pathlib import Path
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from datetime import datetime
import pytz

# === Setup Logging ===
# This sets up logging to the console (GitHub Actions will capture this)
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# === Setup: Linux-compatible download directory ===
download_dir = os.path.join(os.getcwd(), "download")
os.makedirs(download_dir, exist_ok=True)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")  # Comment this line for debug
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

pattern = "Stock Opening  Closing Report (stock.opening.closing)"

def is_file_downloaded():
    return any(Path(download_dir).glob(f"*{pattern}*.xlsx"))

# === Debugging Loop ===
while True:  # Infinite loop until the file is downloaded
    try:
        log.info("Attempting to start the browser...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 20)

        log.info("Navigating to login page...")
        driver.get("https://taps.odoo.com")

        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(2)

        time.sleep(2)
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop")))
        except:
            pass

        switcher_span = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            "div.o_menu_systray div.o_switch_company_menu > button > span"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher_span)
        switcher_span.click()
        time.sleep(2)

        log.info("Selecting 'Zipper' company...")
        target_div = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//div[contains(@class, 'log_into')][span[contains(text(), 'Metal')]]"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", target_div)
        target_div.click()
        time.sleep(4)

        driver.get("https://taps.odoo.com/web#action=441&model=stock.picking.type&view_type=kanban&cids=1&menu_id=280")
        wait.until(EC.presence_of_element_located((By.XPATH, "//html")))
        time.sleep(5)

        log.info("Clicking export button...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/header/nav/div[1]/div[3]/button/span"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/header/nav/div[1]/div[3]/div/a[2]"))).click()
        time.sleep(5)
        first_day_of_month = datetime(2025, 1, 1).strftime("%d/%m/%Y")
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='from_date_0']"))).send_keys(first_day_of_month)
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]"))).click()
        time.sleep(60)

        log.info("Confirming file export...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[1]/div/div[2]/div[3]/button"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/table/thead/tr/th[1]/div"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[1]/span/a[1]"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/button"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div/span"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select/option[34]"))).click()
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/button[1]"))).click()
        time.sleep(50)

        if is_file_downloaded():
            log.info("✅ File download complete!")
            files = list(Path(download_dir).glob(f"*{pattern}*.xlsx"))
            if len(files) > 1:
                files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                for file in files[1:]:
                    file.unlink()
            driver.quit()
            break  # Exit the loop after file download is complete
        else:
            log.warning("⚠️ File not downloaded. Retrying...")

    except Exception as e:
        log.error(f"❌ Error occurred: {e}\nRetrying in 10 seconds...\n")
        try:
            driver.quit()
        except:
            pass
        time.sleep(10)

import requests
import pandas as pd
from datetime import datetime, date
from dotenv import load_dotenv
import os
import pytz
import logging
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from pathlib import Path
import time

# === Load env & config ===
load_dotenv()
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")
COMPANIES = {1: "Zipper", 3: "Metal Trims"}
FROM_DATE = "2025-01-01"
TO_DATE = date.today().strftime("%Y-%m-%d")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# === Logging ===
logging.basicConfig(level=logging.INFO)
log = logging.getLogger()

session = requests.Session()
USER_ID = None

# === Login ===
def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        log.info(f"✅ Logged in (uid={USER_ID})")
        return result
    raise Exception("❌ Login failed")

# === Switch company ===
def switch_company(company_id):
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

# === Create wizard & compute forecast ===
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

# === Fetch opening/closing ===
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
                "specification": {"product_id": {"fields": {"display_name": {}}}, "opening_qty": {}, "opening_value": {}, "receive_qty": {}, "receive_value": {}, "issue_qty": {}, "issue_value": {}, "cloing_qty": {}, "cloing_value": {}},
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
        log.info(f"📊 {cname}: {len(data)} rows fetched")
        return pd.DataFrame(data)
    except Exception:
        log.error(f"❌ Failed to fetch {cname}: {r.text[:200]}")
        return pd.DataFrame()

# === Upload to Google Sheets ===
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
    time.sleep(1)
    set_with_dataframe(worksheet, df)
    tz = pytz.timezone("Asia/Dhaka")
    timestamp = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    worksheet.update("AA2", [[timestamp]])
    log.info(f"✅ Data pasted to {worksheet_name} & timestamp updated: {timestamp}")

# === MAIN SYNC ===
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
                # Paste to Google Sheets
                paste_to_google_sheet(df, sheet_key="1tSgmESOWqYRkDk_KhewnaaJmQGSUaSILzOzpade9tRc", worksheet_name=cname)
