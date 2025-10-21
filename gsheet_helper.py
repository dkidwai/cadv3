import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import streamlit as st

# --- Friendly check so the app doesn't crash if secrets.toml is missing ---
def _load_service_account():
    if "service_account" not in st.secrets:
        st.error(
            "Google credentials not found.\n\n"
            "Add **.streamlit/secrets.toml** with a [service_account] block "
            "(or set them in Streamlit Cloud: Settings → Secrets)."
        )
        raise SystemExit
    return dict(st.secrets["service_account"])

SERVICE_ACCOUNT_INFO = _load_service_account()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SHEET_NAME = "CentralAutomationDB"  # Use your Google Sheet name

creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
client = gspread.authorize(creds)

@st.cache_resource
def _open_spreadsheet():
    return client.open(SHEET_NAME)


def get_sheet(sheet_name):
    """Return the worksheet object, create it if not exists."""
    sh = _open_spreadsheet()  # use cached spreadsheet
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
