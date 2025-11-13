import pandas as pd
import asyncio
import json
import pyodbc
import time
import os
from urllib.parse import urlparse
from playwright.async_api import async_playwright
from datetime import datetime

# ---------------------- SQL Helper ----------------------
def connect_to_sql():
    try:
        conn = pyodbc.connect(SQL_CONNECTION_STRING)
        print("✅ Connected to SQL Server successfully.")
        return conn
    except Exception as e:
        print(f"❌ Error connecting to SQL Server: {e}")
        return None

# ==============================
#   CONFIGURATION SECTION
# ==============================
DATA_SOURCE = "sql"   # Change to "excel" when reading from SQL Server
EXCEL_FILE = "users with domain logout.xlsx"

# SQL Server connection details
SQL_CONNECTION_STRING = (
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=localhost;"
    "Database=PortalHealthCheckDB;"
    "Trusted_Connection=yes;"
)

# ==============================
#   DATA LOADING SECTION
# ==============================
if DATA_SOURCE.lower() == "excel":
    print("📗 Reading input from Excel file...")
    df = pd.read_excel(EXCEL_FILE)

elif DATA_SOURCE.lower() == "sql":
    print("🧩 Reading input from SQL Server...")
    try:
        conn_load = pyodbc.connect(SQL_CONNECTION_STRING)
        query = "SELECT Username, Password, URL, Domain FROM ExcelLoginStatus;"
        df = pd.read_sql(query, conn_load)
        conn_load.close()
    except Exception as e:
        print(f"❌ SQL Read Error: {e}")
        df = pd.DataFrame()  # fallback if SQL fails

else:
    raise ValueError("❌ Invalid DATA_SOURCE! Please use 'excel' or 'sql'.")


# ========== LOAD SELECTORS ==========
def load_selectors(file="selectors.json"):
    with open(file, "r", encoding="utf-8") as f:
        return json.load(f)


# ========== CORE PLAYWRIGHT HELPERS ==========
async def detect_and_fill(page, username, password, selectors):
    for u_sel in selectors.get("username", []):
        for p_sel in selectors.get("password", []):
            try:
                user_box = await page.query_selector(u_sel)
                pass_box = await page.query_selector(p_sel)
                if user_box and pass_box:
                    await user_box.fill(username)
                    await pass_box.fill(password)
                    print(f"✅ Filled fields: {u_sel}, {p_sel}")
                    return True
            except:
                continue
    return False


async def click_login_button(page, selectors):
    for b_sel in selectors.get("button", []):
        try:
            btn = await page.query_selector(b_sel)
            if btn:
                await btn.click()
                print(f"🖱️ Clicked login button: {b_sel}")
                return True
        except:
            continue
    return False


async def perform_logout(page, selectors):
    for l_sel in selectors.get("logout", []):
        try:
            logout_btn = await page.query_selector(l_sel)
            if logout_btn:
                await logout_btn.click()
                await asyncio.sleep(2)
                print(f"🚪 Logged out using: {l_sel}")
                return "✅ Logged Out"
        except:
            continue
    print("⚠️ Logout button not found.")
    return "⚠️ Logout Not Found"


# ========== MAIN FUNCTION ==========
async def check_logins(excel_file="users with domain logout.xlsx", selector_file="selectors.json"):
    """Main automation + SQL integration"""

   # -------------------------------
# ✅ Choose data source dynamically
# -------------------------------
    if DATA_SOURCE == "sql":
        print("📡 Reading data from SQL Server...")
        conn = pyodbc.connect(SQL_CONNECTION_STRING)
        df = pd.read_sql("SELECT * FROM ExcelLoginStatus", conn)
    else:
        print("📘 Reading data from Excel file...")
        try:
          df = pd.read_excel(EXCEL_FILE)
        except Exception as e:
          print(f"❌ Error opening Excel file: {e}\nClose it and retry.")

    selector_map = load_selectors(selector_file)

    # Ensure required columns exist
    for col in ["domain", "status", "logout_status", "last_checked"]:
        if col not in df.columns:
            df[col] = ""

    # Connect to SQL
    conn = connect_to_sql()
    if conn:
        cursor = conn.cursor()
        # Ensure table exists
        cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ExcelLoginStatus' AND xtype='U')
        CREATE TABLE ExcelLoginStatus (
            ID INT IDENTITY(1,1) PRIMARY KEY,
            Username NVARCHAR(100),
            Password NVARCHAR(100),
            URL NVARCHAR(255),
            Domain NVARCHAR(100),
            Status NVARCHAR(100),
            Logout_Status NVARCHAR(100),
            Last_Checked DATETIME
        );
        """)
        conn.commit()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=1000)
        page = await browser.new_page()

        total_rows = len(df)

        for i, row in df.iterrows():
            # ✅ Read data safely
            url = str(row.get("URL", "")).strip()
            username = str(row.get("Username", "")).strip()
            password = str(row.get("Password", "")).strip()
            domain = str(row.get("Domain", "")).strip() or urlparse(url).netloc


            # ✅ Validate URL before using it
            if url and not url.lower().startswith("http"):
                url = "https://" + url

            if not url:
                print(f"⚠️ No valid URL found for row {i+1}, skipping.")
                continue

            print(f"\n🌍 Navigating to: {url}")
            selectors = selector_map.get(domain, selector_map.get("default", {}))

            if not selectors:
                print(f"⚠️ No selectors found for {domain}")
                df.at[i, "status"] = "⚠️ No Selectors Found"
                continue

            try:
                # ✅ Navigate to the page
                await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                await asyncio.sleep(3)

                # Fill username/password
                filled = await detect_and_fill(page, username, password, selectors)
                if not filled:
                    raise Exception("Username/Password fields not found")

                # Click login button
                clicked = await click_login_button(page, selectors)
                if not clicked:
                    raise Exception("Login button not found")

                await asyncio.sleep(4)
                html = (await page.content()).lower()

                # ✅ Determine login result
                if "welcome" in html or "success" in html:
                    status = "✅ Login Successful"
                elif "invalid" in html or "error" in html or "wrong" in html:
                    status = "❌ Invalid Credentials"
                else:
                    status = "⚠️ Unknown Response"

                # ✅ Logout logic
                if i == total_rows - 1:
                    logout_status = "⏭️ Logout Skipped (Last Row)"
                elif status == "❌ Invalid Credentials":
                    logout_status = "⏭️ Logout Skipped (Invalid Login)"
                else:
                    logout_status = await perform_logout(page, selectors)

            except Exception as e:
                status = f"❌ Error: {str(e)[:60]}"
                logout_status = "—"

            # ✅ Update Excel columns
            df.at[i, "Domain"] = domain
            df.at[i, "Status"] = status
            df.at[i, "Logout_Status"] = logout_status
            df.at[i, "Last_Checked"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ✅ Update existing data in SQL
    if conn:
        try:
            now_time = datetime.now()
            cursor.execute(
                """
                UPDATE ExcelLoginStatus
                SET Last_Checked = ?;
                """,
                (now_time,)
            )
            conn.commit()
            print(f"✅ Updated Last_Checked for all URLs at {now_time}")

        except Exception as e:
            print(f"💥 SQL Update Error while updating all URLs: {e}")


        # ✅ Update Excel 'last_checked' with date and time
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.at[i, "last_Checked"] = current_time
        print(f"🕒 Updated Excel last_Checked for row {i+1}: {current_time}")

        # ✅ Save Excel safely
        for retry in range(3):
            try:
                temp_file = excel_file.replace(".xlsx", f"_tmp{retry}.xlsx")
                df.to_excel(temp_file, index=False)
                os.replace(temp_file, excel_file)
                break
            except PermissionError:
                print("⚠️ Excel file is open, retrying in 3 seconds...")
                time.sleep(3)

            print(f"➡️ {status} | {logout_status}")

        await browser.close()

    if conn:
        conn.close()
        print("🔒 SQL connection closed.")

    print("\n✅ All URLs checked, Excel & SQL updated successfully!")


# ========== RUN SCRIPT ==========
asyncio.run(check_logins())
