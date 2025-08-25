import streamlit as st
import pandas as pd
import base64
import io
import gsheet_helper
from gsheet_helper import load_sheet_from_db, save_sheet_to_db


def clean_df(df):
    df = df.loc[:, [col for col in df.columns if not str(col).lower().startswith("unnamed")]]
    df = df.astype(str)
    df = df.replace(['nan', 'NaN', 'None', 'NONE'], '')
    df = df.dropna(axis=1, how='all')
    df = df.loc[:, (df != '').any(axis=0)]
    return df

# --------- STYLES ---------
def set_bg_all():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: #fff !important;
        }}
        .block-container {{
            max-width: 1200px !important;
            width: 96vw !important;
            margin-left: auto !important;
            margin-right: auto !important;
            background: #f4e7da !important;
            border-radius: 20px !important;
            box-shadow: 0 8px 32px #00000011;
            padding-left: 2vw !important;
            padding-right: 2vw !important;
        }}
        /* All navigation and area buttons as gradient filled */
        .stButton > button, .stDownloadButton > button {{
            background: linear-gradient(90deg,#299bff 10%, #55e386 90%) !important;
            color: #000000 !important;
            font-weight: 900 !important;
            border-radius: 18px !important;
            font-size: 1.22rem !important;
            min-height: 54px !important;
            margin-bottom: 8px !important;
            box-shadow: 0 2px 12px #8fd3fe60;
            border: none !important;
            border-width: 0px !important;
            border-style: none !important;
            transition: filter 0.17s;
        }}
        .stButton > button:active, .stDownloadButton > button:active {{
            filter: brightness(0.93);
        }}
        /* Dataframe */
        .stDataFrame {{
            background: #fff !important;
            border-radius: 24px !important;
            box-shadow: 0 8px 32px #00000011;
            margin-left: 0 !important;
            margin-right: 0 !important;
            width: 100% !important;
            max-width: 100% !important;
        }}
        .stDataFrame thead tr th {{
            background: linear-gradient(90deg,#299bff 10%, #55e386 90%) !important;
            color: #273972 !important;
            font-size: 20px !important;
            font-weight: bold !important;
            border-radius: 0 !important;
        }}
        .stDataFrame tbody tr td {{
            font-size: 17px !important;
        }}
        /* Big Logo Section */
        .logo-title-box {{
            background: none !important;
            display: flex;
            flex-direction: column;
            align-items: center;
            margin-bottom: 0;
            margin-top: 0 !important;
        }}
        .logo-title-box img {{
            width: 250px;
            height: 250px;
            margin-bottom: 8px;
            border-radius: 0;
            box-shadow: 0 8px 36px #64b5ff33;
        }}
        .logo-title-box .main-title {{
            font-size: 3.1rem;
            color: #2056b5;
            font-weight: 900;
            text-align: center;
            text-shadow: 0 3px 14px #cce3ff;
            margin-bottom: 8px;
        }}
        /* Search Bar and Area Buttons */
        .stTextInput input {{
            background: linear-gradient(90deg,#e3f4ff 70%, #e9ffe4 100%);
            color: #2056b5 !important;
            border-radius: 13px !important;
            font-size: 1.22rem !important;
            font-weight: 600;
            border: 1px solid #64b5ff !important;
        }}
        label[for^="search_in_area"], label[for^="univ_search"] {{
            color: #2056b5 !important;
            font-weight: 900 !important;
            font-size: 1.18rem !important;
        }}
        /* Dropdown */
        .stSelectbox>div, div[data-baseweb="select"]>div {{
            background: linear-gradient(90deg,#299bff 10%, #55e386 90%) !important;
            color: #fff !important;
            border-radius: 13px !important;
            font-size: 1.18rem !important;
            font-weight: bold !important;
            min-height: 42px !important;
        }}
        .stSelectbox label, .stSelectbox>div>div, .stSelectbox>div>div>div {{
            color: #fff !important;
            font-size: 1.08rem !important;
        }}
        /* Expander */
        .stExpanderHeader {{
            color: #299bff !important;
            font-weight: 900 !important;
            font-size: 1.13rem !important;
        }}
        /* Action bar */
        .action-btn-container {{
            display: flex;
            flex-direction: row;
            justify-content: flex-start;
            align-items: center;
            margin-top: 20px;
            margin-bottom: 20px;
            gap: 22px;
        }}
        /* Hide hamburger menu */
        header .st-emotion-cache-1avcm0n{{
            display: none !important;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

def show_logo_and_title():
    with open("logo.png", "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode()
    st.markdown(
        f"""
        <div class="logo-title-box">
            <img src="data:image/png;base64,{encoded}">
            <div class="main-title">CHANDRAGUPTA</div>
        </div>
        """, unsafe_allow_html=True
    )


def render_mopr():
    # Title/Header
    show_logo_and_title()
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    st.markdown("### MOPR — Star Topology")

    # Load data from the 'MOPR' worksheet
    try:
        df = load_sheet_from_db("MOPR")
    except Exception as e:
        st.error(f"Could not load MOPR sheet: {e}")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_err"):
            st.session_state.main_view = "dashboard"
        return

    if df.empty:
        st.info("No entries found in MOPR sheet. Add rows with columns: Department, PPT_URL, Date.")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_empty"):
            st.session_state.main_view = "dashboard"
        return

    # Normalize column names
    df = df.rename(columns={c: c.strip() for c in df.columns})
    lower_map = {c.lower(): c for c in df.columns}

    if "department" not in lower_map or "ppt_url" not in lower_map:
        st.error("MOPR sheet must have columns: Department, PPT_URL, (optional) Date")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_cols"):
            st.session_state.main_view = "dashboard"
        return

    dept_col = lower_map["department"]
    url_col  = lower_map["ppt_url"]
    date_col = lower_map.get("date", None)

    df_work = df[[dept_col, url_col] + ([date_col] if date_col else [])].copy()

    # Pick latest PPT per department (by Date if present)
    if date_col:
        df_work["__date"] = pd.to_datetime(df_work[date_col], errors="coerce")
        df_work = (df_work.sort_values("__date")
                           .dropna(subset=["__date"])
                           .groupby(dept_col, as_index=False)
                           .tail(1))
    else:
        df_work = df_work.drop_duplicates(subset=[dept_col], keep="last")

    # Build item list
    rows = []
    for _, r in df_work.iterrows():
        d = str(r.get(dept_col, "")).strip()
        u = str(r.get(url_col, "")).strip()
        if d and u:
            rows.append((d, u))

    if not rows:
        st.info("No valid Department / PPT_URL pairs found in MOPR.")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_nodata"):
            st.session_state.main_view = "dashboard"
        return

    # Simple radial (star) layout with HTML buttons
    import math
    n = len(rows)
    size = 480       # container size (px)
    center = size // 2
    radius = 180     # circle radius (px)
    btn_w = 140
    btn_h = 44

    nodes_html = []
    for i, (dept, url) in enumerate(rows):
        angle = 2 * math.pi * i / max(n, 1)
        x = center + int(radius * math.cos(angle)) - btn_w // 2
        y = center + int(radius * math.sin(angle)) - btn_h // 2
        # keep the pill fully inside the box (6px padding)
        x = max(6, min(size - btn_w - 6, x))
        y = max(6, min(size - btn_h - 6, y))

        nodes_html.append(f"""
            <a href="{url}" target="_blank" rel="noopener"
               style="
                 position:absolute; left:{x}px; top:{y}px;
                 padding:10px 18px; border-radius:18px;
                 background:linear-gradient(90deg,#299bff 10%, #55e386 90%);
                 color:#000; font-weight:900; text-decoration:none;
                 box-shadow:0 2px 12px #8fd3fe60; white-space:nowrap;">
               {dept}
            </a>
        """)

    html = f"""
    <div style="position:relative; width:{size}px; height:{size}px; margin:10px auto 20px auto;
                background:#ffffff; border-radius:16px; box-shadow:0 8px 32px #00000011;">
        <div style="position:absolute; left:{center-60}px; top:{center-24}px;
                    padding:12px 20px; border-radius:18px; background:#f4e7da;
                    color:#2056b5; font-weight:900; box-shadow:0 2px 12px #8fd3fe60;">
            MOPR
        </div>
        {''.join(nodes_html)}
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    if st.button("⬅️ Back to Dashboard", key="mopr_back_btn"):
        st.session_state.main_view = "dashboard"

# ---------- SHEETS & NAVIGATION STATE ----------
all_subsections = [
    "PLC DETAILS", "OS DETAILS", "SINGLE POINT TRIPPING", "PAIN POINT",
    "IO LIST", "CRITICAL SPARES","BACKUP","PANEL EARTHING","AUDIT","INVENTORY","DIRECTORY"
]
DASHBOARD_VIEW = "dashboard"
SHEET_VIEW = "sheet"
AREA_VIEW = "area"
SEARCH_VIEW = "search"
MOPR_VIEW = "mopr"


if "login" not in st.session_state:
    st.session_state.login = None
if "main_view" not in st.session_state:
    st.session_state.main_view = DASHBOARD_VIEW
if "selected_sheet" not in st.session_state:
    st.session_state.selected_sheet = None
if "selected_area" not in st.session_state:
    st.session_state.selected_area = None
if "search_sheet" not in st.session_state:
    st.session_state.search_sheet = all_subsections[0]
if "db_uploaded" not in st.session_state:
    st.session_state.db_uploaded = False
if "io_selected_sheet" not in st.session_state:
    st.session_state.io_selected_sheet = None

# --- LOGIN SMALL BOX, Centered ---
ADMIN_USERS = {'admin1': 'pass@1', 'danish': 'dk@1245@', 'avinash': 'avinash_1246#'}
VIEWERS = {'user': '@1234', 'guest': '@guest'}

set_bg_all()  # Always apply background

# --- User/Admin at the very top ---
if st.session_state.login:
    st.markdown(f"<div style='text-align:right;font-weight:bold;color:#2056b5; margin-bottom:0; margin-top:0;'>User: {st.session_state.login['user']} ({st.session_state.login['role']})</div>", unsafe_allow_html=True)

if st.session_state.login is None:
    show_logo_and_title()
    st.markdown(
        """
        <div style="max-width:200px;margin:16px auto 8px auto;">
        """,
        unsafe_allow_html=True
    )
    login_user = st.text_input("Username")
    login_pass = st.text_input("Password", type="password")
    login_btn = st.button("Login")
    st.markdown("</div>", unsafe_allow_html=True)
    if login_btn:
        if login_user in ADMIN_USERS and login_pass == ADMIN_USERS[login_user]:
            st.session_state.login = {"user": login_user, "role": "admin"}
            st.success("Admin login successful.")
            st.rerun()
        elif login_user in VIEWERS and login_pass == VIEWERS[login_user]:
            st.session_state.login = {"user": login_user, "role": "viewer"}
            st.success("Viewer login successful.")
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.stop()

login_name = st.session_state.login["user"]
login_role = st.session_state.login["role"]
IS_ADMIN = (login_role == "admin")  # <-- viewers are read-only
st.sidebar.markdown(f"**👤 Logged in as:** `{login_name}` ({login_role.capitalize()})")

# --- DB Upload/Init
uploaded_file = None
if login_role == "admin":
    uploaded_file = st.sidebar.file_uploader("Upload Excel (.xlsx) [admin only]", type=["xlsx"])
    if uploaded_file and not st.session_state.db_uploaded:
        xl = pd.ExcelFile(uploaded_file)
        available_sheets = [s for s in all_subsections if s in xl.sheet_names]
        if not available_sheets:
            st.error("No relevant sheets found in this Excel file.")
        else:
            skipped_sheets = []
            loaded_sheets = []
            for sheet in available_sheets:
                df = xl.parse(sheet)
                df = clean_df(df)
                if df.empty or len(df.columns) == 0:
                    skipped_sheets.append(sheet)
                    continue
                save_sheet_to_db(sheet, df)
                loaded_sheets.append(sheet)
            if not loaded_sheets:
                st.error("No sheets could be loaded from your Excel. Please check your file.")
                st.stop()
            msg = ""
            if loaded_sheets:
                msg += f"Loaded: {', '.join(loaded_sheets)}. "
            if skipped_sheets:
                msg += f"Skipped: {', '.join(skipped_sheets)}."
            st.success(f"Database refreshed! {msg}")
            st.session_state.db_uploaded = True
            st.rerun()

# ---- Always load available sheets from DB ----
available_sheets_db = []
for s in all_subsections:
    try:
        df_check = load_sheet_from_db(s)
        if not df_check.empty:
            available_sheets_db.append(s)
    except:
        pass

if not available_sheets_db:
    st.info("👈 Please (Admin) upload your Excel file once to initialize the database.")
    if st.sidebar.button("Logout"):
        st.session_state.login = None
        st.rerun()
    st.stop()

# =========== MAIN DASHBOARD VIEW ===========
if st.session_state.main_view == DASHBOARD_VIEW:
    show_logo_and_title()
    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
    # --- Navigation Buttons Only (NO Card/Box) ---
    btn_cols = st.columns(3)
    for idx, sheet in enumerate(all_subsections):
        with btn_cols[idx % 3]:
            if st.button(f"{sheet}", key=f"btn_sheet_{sheet}", use_container_width=True):
                st.session_state.selected_sheet = sheet
                st.session_state.main_view = SHEET_VIEW
                st.session_state.selected_area = None
        if (idx + 1) % 3 == 0 and idx != len(all_subsections) - 1:
            btn_cols = st.columns(3)

    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)
    st.markdown('<div style="width:100%;max-width:340px;margin:auto;">', unsafe_allow_html=True)
    if st.button("🔍 Universal Search", key="big_univ_search", use_container_width=True):
        st.session_state.main_view = SEARCH_VIEW
    st.markdown('</div>', unsafe_allow_html=True)
    # --- MOPR launcher (separate; keeps your current structure unchanged)
    st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)
    st.markdown('<div style="width:100%;max-width:340px;margin:auto;">', unsafe_allow_html=True)
    if st.button("⭐ MOPR", key="open_mopr", use_container_width=True):
       st.session_state.main_view = MOPR_VIEW
    st.markdown('</div>', unsafe_allow_html=True)


# =========== SHEET => AREAS VIEW ===========
elif st.session_state.main_view == SHEET_VIEW:
    show_logo_and_title()
    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
    sheet = st.session_state.selected_sheet
    st.markdown(f"#### {sheet}")
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # --- Special hierarchical flow for IO LIST ---
    if sheet == "IO LIST":
        # Build IO map from worksheet titles following IO_AREA_SHEETNAME convention
        try:
            sh = gsheet_helper.client.open(gsheet_helper.SHEET_NAME)
            io_titles = [ws.title for ws in sh.worksheets() if ws.title.upper().startswith("IO_")]
        except Exception as e:
            st.error(f"Could not list IO worksheets: {e}")
            io_titles = []

        io_map = {}
        for t in sorted(io_titles):
            name = t[3:]  # remove 'IO_'
            parts = name.split("_", 1)
            area = parts[0] if parts else name
            sub = parts[1] if len(parts) > 1 else "(Sheet)"
            io_map.setdefault(area, []).append((sub, t))

        # 1) Pick Area
        if st.session_state.selected_area is None:
            if not io_map:
                st.warning("No IO_* worksheets found in database. Add worksheets like IO_AREA_SHEETNAME.")
            else:
                st.markdown("##### Select Area:")
                areacols = st.columns(4)
                for idx, area in enumerate(sorted(io_map.keys())):
                    with areacols[idx % 4]:
                        if st.button(area, key=f"io_area_{area}", use_container_width=True):
                            st.session_state.selected_area = area
                            st.session_state.io_selected_sheet = None
                            st.rerun()
            if st.button("⬅️ Back to Dashboard"):
                st.session_state.main_view = DASHBOARD_VIEW
                st.session_state.selected_sheet = None
                st.session_state.selected_area = None
                st.session_state.io_selected_sheet = None
                st.rerun()

        # 2) Pick Sheet within Area
        elif st.session_state.io_selected_sheet is None:
            st.markdown(f"##### {st.session_state.selected_area} — Select Sheet:")
            sheets = io_map.get(st.session_state.selected_area, [])
            if not sheets:
                st.info("No sheets found for this area.")
            cols = st.columns(3)
            for i, (label, title) in enumerate(sheets):
                with cols[i % 3]:
                    if st.button(label, key=f"io_sheet_{title}", use_container_width=True):
                        st.session_state.io_selected_sheet = title
                        st.rerun()
            c1, c2 = st.columns([1,1])
            with c1:
                if st.button("⬅️ Back to Areas", key="io_back_areas"):
                    st.session_state.selected_area = None
                    st.session_state.io_selected_sheet = None
                    st.rerun()
            with c2:
                if st.button("⬅️ Back to Dashboard", key="io_back_dash_from_sheets"):
                    st.session_state.main_view = DASHBOARD_VIEW
                    st.session_state.selected_sheet = None
                    st.session_state.selected_area = None
                    st.session_state.io_selected_sheet = None
                    st.rerun()

        # 3) Show full sheet data with edit & export
        else:
            data_sheet_name = st.session_state.io_selected_sheet
            df = clean_df(load_sheet_from_db(data_sheet_name))
            search = st.text_input("🔎 Search in this Sheet...", key="search_in_io_sheet")
            if search:
                filtered_df2 = df[df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)]
            else:
                filtered_df2 = df
            filtered_df2 = filtered_df2.astype(str).replace(['nan', 'NaN', 'None', 'NONE'], '')

            st.dataframe(filtered_df2, use_container_width=True, height=480)
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

            # ---- Editing (admin only)
            if len(filtered_df2) > 0:
                if IS_ADMIN:
                    row_indices = filtered_df2.index.tolist()
                    row_options = [f"Row {i+1}" for i in range(len(row_indices))]
                    selected_row = st.selectbox("Select Row to Edit", options=row_options, key="edit_row_select_io")
                    edit_idx = row_indices[row_options.index(selected_row)]
                    edit_cols = {}
                    with st.expander("Edit Selected Row", expanded=False):
                        st.markdown("<span style='color:#299bff;font-weight:bold;font-size:1.18rem;'>Edit Selected Row</span>", unsafe_allow_html=True)
                        for col in filtered_df2.columns:
                            edit_cols[col] = st.text_input(f"{col}", filtered_df2.at[edit_idx, col], key=f"edit_col_io_{col}")
                        if st.button("Update Row", key="update_row_btn_io"):
                            for col in filtered_df2.columns:
                                filtered_df2.at[edit_idx, col] = edit_cols[col]
                            master_df = clean_df(load_sheet_from_db(data_sheet_name))
                            master_df.update(filtered_df2)
                            save_sheet_to_db(data_sheet_name, master_df)
                            st.success("Row updated and saved to database.")
                            st.rerun()
                else:
                    st.info("🔒 Viewer mode: you can view and export this sheet. Editing is restricted to admins.")
            else:
                st.info("No rows to edit.")

            st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
            st.markdown("""
            <div class="action-btn-container">
            """, unsafe_allow_html=True)
            excel_buffer = io.BytesIO()
            filtered_df2.to_excel(excel_buffer, index=False)
            c1, c2, c3 = st.columns([1,1,1])
            with c1:
                st.download_button(
                    label="⬇️ Export Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"{data_sheet_name}_export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="export_btn_io_sheet"
                )
            with c2:
                if st.button("⬅️ Back to Sheets", key="io_back_sheets"):
                    st.session_state.io_selected_sheet = None
                    st.rerun()
            with c3:
                if st.button("⬅️ Back to Areas", key="io_back_areas_from_data"):
                    st.session_state.io_selected_sheet = None
                    st.session_state.selected_area = None
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # --- Default behavior for all other sheets (unchanged) ---
    else:
        df = clean_df(load_sheet_from_db(sheet))
        area_col = None
        for col in df.columns:
            if col.strip().lower() == "area":
                area_col = col
                break
        if area_col:
            df[area_col] = df[area_col].astype(str).str.strip()
            areas = sorted(df[area_col].replace(['', ' ', 'nan', 'NaN', 'None', 'NONE'], pd.NA).dropna().unique())
            st.markdown("##### Select Area:")
            areacols = st.columns(4)
            for idx, area in enumerate(areas):
                with areacols[idx % 4]:
                    if st.button(area, key=f"area_{area}", use_container_width=True):
                        st.session_state.selected_area = area
                        st.session_state.main_view = AREA_VIEW
            if not areas:
                st.warning("No Area values found. Showing all data.")
                if st.button("Show All Data", use_container_width=True):
                    st.session_state.selected_area = "All"
                    st.session_state.main_view = AREA_VIEW
            if st.button("⬅️ Back to Dashboard"):
                st.session_state.main_view = DASHBOARD_VIEW
                st.session_state.selected_sheet = None
                st.session_state.selected_area = None
        else:
            st.warning("No 'Area' column found. Showing all data.")
            if st.button("Show Data Table", use_container_width=True):
                st.session_state.selected_area = "All"
                st.session_state.main_view = AREA_VIEW
            if st.button("⬅️ Back to Dashboard"):
                st.session_state.main_view = DASHBOARD_VIEW
                st.session_state.selected_sheet = None
                st.session_state.selected_area = None

# =========== AREA DATA TABLE VIEW (WITH ROW EDIT) ===========
elif st.session_state.main_view == AREA_VIEW:
    show_logo_and_title()
    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
    sheet = st.session_state.selected_sheet
    area = st.session_state.selected_area
    st.markdown(f"#### {sheet} - {area}")
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    df = clean_df(load_sheet_from_db(sheet))
    area_col = None
    for col in df.columns:
        if col.strip().lower() == "area":
            area_col = col
            break
    if area_col and area != "All":
        filtered_df = df[df[area_col] == area]
    else:
        filtered_df = df
    search = st.text_input("🔎 Search in this Area...", key="search_in_area")
    if search:
        filtered_df2 = filtered_df[filtered_df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)]
    else:
        filtered_df2 = filtered_df
    filtered_df2 = filtered_df2.astype(str).replace(['nan', 'NaN', 'None', 'NONE'], '')

    st.dataframe(filtered_df2, use_container_width=True, height=420)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ---- Editing (admin only)
    if len(filtered_df2) > 0:
        if IS_ADMIN:
            row_indices = filtered_df2.index.tolist()
            row_options = [f"Row {i+1}" for i in range(len(row_indices))]
            selected_row = st.selectbox("Select Row to Edit", options=row_options, key="edit_row_select")
            edit_idx = row_indices[row_options.index(selected_row)]
            edit_cols = {}
            with st.expander("Edit Selected Row", expanded=False):
                st.markdown("<span style='color:#299bff;font-weight:bold;font-size:1.18rem;'>Edit Selected Row</span>", unsafe_allow_html=True)
                for col in filtered_df2.columns:
                    edit_cols[col] = st.text_input(f"{col}", filtered_df2.at[edit_idx, col], key=f"edit_col_{col}")
                if st.button("Update Row", key="update_row_btn"):
                    for col in filtered_df2.columns:
                        filtered_df2.at[edit_idx, col] = edit_cols[col]
                    master_df = clean_df(load_sheet_from_db(sheet))
                    master_df.update(filtered_df2)
                    save_sheet_to_db(sheet, master_df)
                    st.success("Row updated and saved to database.")
                    st.rerun()
        else:
            st.info("🔒 Viewer mode: you can view and export. Editing is restricted to admins.")
    else:
        st.info("No rows to edit.")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="action-btn-container">
        """,
        unsafe_allow_html=True
    )
    excel_buffer = io.BytesIO()
    filtered_df2.to_excel(excel_buffer, index=False)
    c1, c2, c3 = st.columns([1,1,1])
    with c1:
        st.download_button(
            label="⬇️ Export Excel",
            data=excel_buffer.getvalue(),
            file_name=f"{sheet}_{area}_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="export_btn_area"
        )
    with c2:
        if st.button("⬅️ Back to Areas", key="back_areas_btn"):
            st.session_state.main_view = SHEET_VIEW
            st.session_state.selected_area = None
    with c3:
        if st.button("⬅️ Back to Dashboard", key="back_dash_btn"):
            st.session_state.main_view = DASHBOARD_VIEW
            st.session_state.selected_sheet = None
            st.session_state.selected_area = None
    st.markdown("</div>", unsafe_allow_html=True)

# =========== UNIVERSAL SEARCH VIEW ===========
elif st.session_state.main_view == SEARCH_VIEW:
    show_logo_and_title()
    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
    st.markdown("### Universal Search")
    cols = st.columns(len(all_subsections))
    for idx, sheet in enumerate(all_subsections):
        with cols[idx]:
            btn_style = "selected-btn" if st.session_state.search_sheet == sheet else ""
            if st.button(sheet, key=f"search_sheet_{sheet}", use_container_width=True):
                st.session_state.search_sheet = sheet

    sheet = st.session_state.search_sheet
    st.markdown(f"**{sheet}**")
    df = clean_df(load_sheet_from_db(sheet))
    search = st.text_input(f"Search in {sheet}...", key=f"univ_search_{sheet}")
    if search:
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)]
    else:
        filtered_df = df
    filtered_df = filtered_df.astype(str).replace(['nan', 'NaN', 'None', 'NONE'], '')
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    st.dataframe(filtered_df, use_container_width=True, height=650)
    st.markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="action-btn-container">
        """,
        unsafe_allow_html=True
    )
    excel_buffer = io.BytesIO()
    filtered_df.to_excel(excel_buffer, index=False)
    c1, c2 = st.columns([1,1])
    with c1:
        st.download_button(
            label="⬇️ Export Excel",
            data=excel_buffer.getvalue(),
            file_name=f"{sheet}_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="export_btn_search"
        )
    with c2:
        if st.button("⬅️ Back to Dashboard", key="back_dash_search_btn"):
            st.session_state.main_view = DASHBOARD_VIEW
            st.session_state.selected_sheet = None
            st.session_state.selected_area = None
    st.markdown("</div>", unsafe_allow_html=True)
    # =========== MOPR VIEW ===========
elif st.session_state.main_view == MOPR_VIEW:
    render_mopr()


# =========== SIDEBAR LOGOUT ===========
if st.sidebar.button("Logout"):
    st.session_state.login = None
    st.session_state.main_view = DASHBOARD_VIEW
    st.session_state.selected_sheet = None
    st.session_state.selected_area = None
    st.session_state.db_uploaded = False
    st.rerun()
