from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import os
import pandas as pd
import gspread
import warnings

warnings.filterwarnings("ignore", message="Workbook contains no default style*")

load_dotenv()

download_path = os.getenv("DOWNLOAD_PATH")
spreadsheet_id = os.getenv("SPREADSHEET_ID")

files = [
    os.path.join(download_path, f)
    for f in os.listdir(download_path)
    if f.startswith("BRANCH STOCK") and f.endswith(".xlsx")
]

if not files:
    print("File tidak ditemukan ⚠️")
else:
    latest_file = max(files, key=os.path.getctime)

    df = pd.read_excel(latest_file, sheet_name=0, header=5, usecols="B:U")

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet("Data")

    df = df.fillna("")

    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
