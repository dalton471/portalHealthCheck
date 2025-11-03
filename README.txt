Excel Login Automation using Playwright (Multi-URL Version)
------------------------------------------------------------

✅ Features:
- Supports multiple URLs and multiple users per URL.
- Automatically fills login forms and validates responses.
- Detects successful or invalid logins.
- Updates results in Excel after every attempt.
- Automatically installs Playwright browsers if missing.

⚙️ Setup Instructions:

1️⃣ Install dependencies:
    pip install playwright pandas openpyxl
    python -m playwright install

2️⃣ Files included:
   - excel_login_playwright.py  → Main automation script
   - users.xlsx  → Sample Excel input file
   - README.txt  → Setup guide

3️⃣ Run the script:
    python excel_login_playwright.py

4️⃣ The script will:
   - Open each unique URL in your Excel file.
   - Try all username/password pairs for that URL.
   - Update 'status' in Excel after every attempt.

🧩 Excel Format:
| username | password | url | status |

Example Test Site:
https://practicetestautomation.com/practice-test-login/
(username: student, password: Password123)
