import pandas as pd
import asyncio
import json
from urllib.parse import urlparse
from playwright.async_api import async_playwright
from datetime import datetime
import time
import os

def load_selectors(file="selectors.json"):
    with open(file, "r", encoding="utf-8") as f:
        return json.load(f)

async def detect_and_fill(page, username, password, selectors):
    """Find username/password fields and fill them."""
    for u_sel in selectors.get("username", []):
        for p_sel in selectors.get("password", []):
            try:
                user_box = await page.query_selector(u_sel)
                pass_box = await page.query_selector(p_sel)
                if user_box and pass_box:
                    await user_box.fill(username)
                    await pass_box.fill(password)
                    print(f"✅ Filled fields with selectors: {u_sel}, {p_sel}")
                    return True
            except:
                continue
    return False

async def click_login_button(page, selectors):
    """Try clicking the login button using selectors."""
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
    """Click logout button if found."""
    logout_selectors = selectors.get("logout", [])
    if not logout_selectors:
        print("ℹ️ Logout skipped for this domain (no logout selectors).")
        return "Skipped (No Logout)"
    for l_sel in logout_selectors:
        try:
            logout_btn = await page.query_selector(l_sel)
            if logout_btn:
                await logout_btn.click()
                print(f"🚪 Logged out using: {l_sel}")
                await asyncio.sleep(2)
                return "✅ Logged Out"
        except:
            continue
    print("⚠️ Logout button not found.")
    return "⚠️ Logout Failed"

async def check_logins(excel_file="users_with_domain_logout.xlsx", selector_file="selectors.json"):
    """Main Playwright automation flow."""
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"❌ Error opening Excel file: {e}\nClose it and retry.")
        return

    selector_map = load_selectors(selector_file)

    # Ensure all required columns exist
    for col in ["domain", "status", "logout_status", "last_checked"]:
        if col not in df.columns:
            df[col] = ""

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=1200)
        page = await browser.new_page()

        for i, row in df.iterrows():
            url = str(row.get("url", "")).strip()
            username = str(row.get("username", "")).strip()
            password = str(row.get("password", "")).strip()
            domain = str(row.get("domain", "")).strip().lower()

            print(f"\n🌐 Checking domain '{domain}' for user '{username}'")

            # Strict domain matching with JSON
            if domain not in selector_map:
                print(f"⚠️ Domain '{domain}' not found in selectors.json — skipping...")
                df.at[i, "status"] = f"⚠️ Domain '{domain}' not found in selectors.json"
                continue

            selectors = selector_map[domain]

            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                await asyncio.sleep(3)

                filled = await detect_and_fill(page, username, password, selectors)
                if not filled:
                    raise Exception("❌ Username/Password fields not found")

                clicked = await click_login_button(page, selectors)
                if not clicked:
                    raise Exception("❌ Login button not found")

                await asyncio.sleep(4)
                html = (await page.content()).lower()

                if "logged in successfully" in html or "welcome" in html:
                    status = "✅ Login Successful"
                    logout_status = await perform_logout(page, selectors)
                elif "invalid" in html or "wrong" in html:
                    status = "❌ Invalid Credentials"
                    logout_status = "—"
                else:
                    status = "⚠️ Unknown Response"
                    logout_status = "—"

            except Exception as e:
                status = f"❌ Error: {str(e)[:60]}"
                logout_status = "—"

            df.at[i, "status"] = status
            df.at[i, "logout_status"] = logout_status
            df.at[i, "last_checked"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Safe Excel save
            for retry in range(3):
                try:
                    tmp = excel_file.replace(".xlsx", f"_tmp{retry}.xlsx")
                    df.to_excel(tmp, index=False)
                    os.replace(tmp, excel_file)
                    break
                except PermissionError:
                    print("⚠️ Excel file open, retrying in 3s...")
                    time.sleep(3)

            print(f"➡️ {status} | {logout_status}")

        await browser.close()
        print("\n✅ Finished checking all domains!")

asyncio.run(check_logins())




