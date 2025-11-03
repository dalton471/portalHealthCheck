import pandas as pd
import asyncio
import json
import pyodbc
import time
import os
from urllib.parse import urlparse
from playwright.async_api import async_playwright
from datetime import datetime

# ========================
# 🔧 Data source parameter
# ========================
# Choose "excel" or "sql" depending on where you want to pull user data from.
data_source = "excel" or "sql"


# ========== SQL CONNECTION ==========
def connect_to_sql():
    """Connect to SQL Server"""
    try:
        conn = pyodbc.connect(
            "Driver={ODBC Driver 17 for SQL Server};"
            "Server=localhost;"
            "Database=PortalHealthCheckDB;"
            "Trusted_Connection=yes;"
        )
        print("✅ Connected to SQL Server successfully.")
        return conn
    except Exception as e:
        print(f"❌ SQL Connection Error: {e}")
        return None


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
async def check_logins(excel_file="users_with_domain_logout.xlsx", selector_file="selectors.json"):
    """Main automation + SQL integration"""

    # Read Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return

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
            url = str(row.get("url", "")).strip()
            username = str(row.get("username", "")).strip()
            password = str(row.get("password", "")).strip()
            domain = str(row.get("domain", "")).strip() or urlparse(url).netloc

            print(f"\n🌐 Checking {domain} for user '{username}' (Row {i+1}/{total_rows})")
            selectors = selector_map.get(domain, selector_map.get("default", {}))

            if not selectors:
                print(f"⚠️ No selectors for {domain}")
                df.at[i, "status"] = "⚠️ No Selectors Found"
                continue

            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                await asyncio.sleep(3)

                filled = await detect_and_fill(page, username, password, selectors)
                if not filled:
                    raise Exception("Username/Password fields not found")

                clicked = await click_login_button(page, selectors)
                if not clicked:
                    raise Exception("Login button not found")

                await asyncio.sleep(4)
                html = (await page.content()).lower()

                # Identify status by response text
                if "welcome" in html or "success" in html:
                    status = "✅ Login Successful"
                elif "invalid" in html or "error" in html or "wrong" in html:
                    status = "❌ Invalid Credentials"
                else:
                    status = "⚠️ Unknown Response"

                # ✅ Logout logic rules
                if i == total_rows - 1:
                    # Skip logout for last row
                    logout_status = "⏭️ Logout Skipped (Last Row)"
                elif status == "❌ Invalid Credentials":
                    # Skip logout for invalid credentials
                    logout_status = "⏭️ Logout Skipped (Invalid Login)"
                else:
                    # Perform logout normally
                    logout_status = await perform_logout(page, selectors)

            except Exception as e:
                status = f"❌ Error: {str(e)[:60]}"
                logout_status = "—"

            # Update Excel
            df.at[i, "domain"] = domain
            df.at[i, "status"] = status
            df.at[i, "logout_status"] = logout_status
            df.at[i, "last_checked"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # ✅ Insert into SQL
            if conn:
                try:
                    cursor.execute("""
                        INSERT INTO ExcelLoginStatus
                        (Username, Password, URL, Domain, Status, Logout_Status, Last_Checked)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, username, password, url, domain, status, logout_status, datetime.now())
                    conn.commit()
                    print("🗂️ Inserted into SQL successfully.")
                except Exception as e:
                    print(f"⚠️ SQL Insert Error: {e}")

            # Save Excel safely
            for retry in range(3):
                try:
                    temp = excel_file.replace(".xlsx", f"_tmp{retry}.xlsx")
                    df.to_excel(temp, index=False)
                    os.replace(temp, excel_file)
                    break
                except PermissionError:
                    print("⚠️ Excel open, retrying in 3s...")
                    time.sleep(3)

            print(f"➡️ {status} | {logout_status}")

        await browser.close()

    if conn:
        conn.close()
        print("🔒 SQL connection closed.")

    print("\n✅ All URLs checked, Excel & SQL updated successfully!")


# ========== RUN SCRIPT ==========
asyncio.run(check_logins())




