import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import streamlit as st

# --- Read credentials from Streamlit secrets ---
SERVICE_ACCOUNT_INFO = dict(st.secrets["service_account"])

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
SHEET_NAME = "CentralAutomationDB"  # Use your Google Sheet name

creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
client = gspread.authorize(creds)

def get_sheet(sheet_name):
    """Return the worksheet object, create it if not exists."""
    sh = client.open(SHEET_NAME)
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows="1000", cols="20")
    return ws

@st.cache_data(ttl=180)
def load_sheet_from_db(sheet_name):
    """Load data from Google Sheet worksheet into pandas DataFrame."""
    ws = get_sheet(sheet_name)
    data = ws.get_all_values()
    if not data or len(data) < 2:
        return pd.DataFrame()
    header, *values = data
    df = pd.DataFrame(values, columns=header)
    return df

def save_sheet_to_db(sheet_name, df):
    """Save pandas DataFrame to Google Sheet worksheet (replace all data)."""
    ws = get_sheet(sheet_name)
    ws.clear()
    ws.append_row(df.columns.tolist())
    rows = df.astype(str).values.tolist()
    if rows:
        ws.append_rows(rows)
    load_sheet_from_db.clear()
