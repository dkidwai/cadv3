import streamlit as st
import pandas as pd
import base64
import io
import gsheet_helper
from gsheet_helper import load_sheet_from_db, save_sheet_to_db
from textwrap import dedent
import re
import os
import json

# --- Fix for Excel export: strip illegal XML/control chars ---
_ILLEGAL_CTRL = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def _sanitize_excel_str(v):
    s = "" if v is None else str(v)
    s = (s.replace("\u00A0", " ")  # NBSP -> space
         .replace("\uFEFF", "")    # BOM
         .replace("\u200B", "")    # zero-width space
         .replace("\u200C", "")    # ZWNJ
         .replace("\u200D", "")    # ZWJ
         .replace("\u2060", ""))   # word joiner
    s = _ILLEGAL_CTRL.sub("", s)   # remove ASCII control chars Excel disallows
    return s

def sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if hasattr(df, "map"):          # pandas >= 2.1
        return df.map(_sanitize_excel_str)
    return df.applymap(_sanitize_excel_str)   # purane pandas ke liye




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
    """MOPR star topology with Month + FY filters.

    - Preferred filtering uses a Date column.
    - Fallback uses optional Month/FY columns if Date is missing.
    - FY logic is India style (Apr to Mar).

    Required (aliases ok):
    - Department
    - PPT_URL / Link / URL / PPT
    - Date (recommended)
    """

    import math, html

    show_logo_and_title()
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    st.markdown("### MOPR — Star Topology")

    # --- dropdown label + selected value color (black) ---
    # Force label text + selected values to black.
    # (More aggressive selectors to override app theme styles.)
    st.markdown(
        """
        <style>
        /* 1) Global label force (covers most Streamlit versions) */
        label {
            color: #000000 !important;
            font-weight: 700 !important;
        }

        /* 2) Specific widget label wrapper (newer Streamlit) */
        div[data-testid="stWidgetLabel"] > label {
            color: #000000 !important;
            font-weight: 700 !important;
        }

        /* 3) Selectbox label inside columns/containers */
        div[data-testid="stSelectbox"] label {
            color: #000000 !important;
            font-weight: 700 !important;
        }

        /* 4) BaseWeb Select (Streamlit selectbox) - selected value + arrow */
        div[data-baseweb="select"] * {
            color: #000000 !important;
        }
        div[data-baseweb="select"] svg {
            fill: #000000 !important;
        }

        /* 5) Dropdown menu items */
        div[role="listbox"] * {
            color: #000000 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # --- load ---
    try:
        df = load_sheet_from_db("MOPR")
        df = df.loc[:, [c for c in df.columns if not str(c).lower().startswith("unnamed")]]
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

    # --- normalize column names (no regex to keep copy-safe) ---
    def _norm(name: str) -> str:
        s = str(name).replace(chr(160), " ").strip().lower()
        s = s.replace("-", " ")
        s = " ".join(s.split())
        s = s.replace(" ", "_")
        return s

    norm_map = {_norm(c): c for c in df.columns if str(c).strip()}

    dept_col = next((norm_map[k] for k in ("department", "dept", "departments") if k in norm_map), None)
    url_col  = next((norm_map[k] for k in ("ppt_url", "ppturl", "ppt_link", "ppt", "url", "link") if k in norm_map), None)
    date_col = next((norm_map[k] for k in ("date", "updated", "updated_on", "last_updated", "last_update") if k in norm_map), None)

    month_col = next((norm_map[k] for k in ("month", "mth", "period") if k in norm_map), None)
    fy_col    = next((norm_map[k] for k in ("financial_year", "financialyear", "fy") if k in norm_map), None)

    if not dept_col or not url_col:
        st.error("MOPR sheet must have columns like: Department + PPT_URL (aliases ok).")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_cols"):
            st.session_state.main_view = "dashboard"
        return

    keep_cols = [dept_col, url_col]
    if date_col:
        keep_cols.append(date_col)
    if (not date_col) and month_col:
        keep_cols.append(month_col)
    if (not date_col) and fy_col:
        keep_cols.append(fy_col)

    df_work = df[keep_cols].copy()
    df_work = df_work.dropna(subset=[dept_col])

    # ---------- URL extraction (no regex) ----------
    stop_chars = " \"<>" + chr(9) + chr(10) + chr(13)

    def extract_url(val) -> str:
        s = str(val or "")
        s = (s.replace(chr(160), " ")
               .replace(chr(8203), "")
               .replace(chr(8204), "")
               .replace(chr(8205), "")
               .strip())

        s_low = s.lower()

        # Case 1: HYPERLINK("url","text")
        if "hyperlink(" in s_low:
            start = s_low.find("hyperlink(")
            q1 = s.find('"', start)
            if q1 != -1:
                q2 = s.find('"', q1 + 1)
                if q2 != -1:
                    return s[q1 + 1:q2].strip()

        # Case 2: plain http(s)
        for proto in ("https://", "http://"):
            p = s_low.find(proto)
            if p != -1:
                out = []
                for ch in s[p:]:
                    if ch in stop_chars:
                        break
                    out.append(ch)
                return "".join(out).strip()

        return ""

    # ---------- FY helper (Apr–Mar) ----------
    def fiscal_year_label(dt: pd.Timestamp) -> str:
        if pd.isna(dt):
            return ""
        y = int(dt.year)
        if int(dt.month) >= 4:
            start = y
            end = y + 1
        else:
            start = y - 1
            end = y
        return f"FY {start}-{str(end)[-2:]}"

    selected_fy = "All"
    selected_month = "All"

    # ---------- Filtering ----------
    if date_col:
        df_work["__date"] = pd.to_datetime(df_work[date_col], errors="coerce")
        df_work["__month_label"] = df_work["__date"].dt.strftime("%b %Y")
        df_work["__fy"] = df_work["__date"].apply(fiscal_year_label)

        df_with_dates = df_work.dropna(subset=["__date"]).copy()

        fy_vals = [v for v in df_with_dates["__fy"].dropna().unique().tolist() if str(v).strip()]

        def _fy_sort_key(v: str) -> int:
            digits = "".join([c for c in str(v) if c.isdigit()])
            return int(digits[:4]) if len(digits) >= 4 else 0

        fy_vals = sorted(fy_vals, key=_fy_sort_key, reverse=True)

        c_fy, c_mo, _sp = st.columns([1.2, 1.3, 3.5])
        with c_fy:
            selected_fy = st.selectbox("Financial Year", options=["All"] + fy_vals, index=0, key="mopr_fy_select")

        tmp = df_with_dates
        if selected_fy != "All":
            tmp = tmp[tmp["__fy"] == selected_fy]

        month_df = tmp[["__month_label", "__date"]].dropna(subset=["__date"]).sort_values("__date")

        month_order = []
        seen = set()
        for _, r in month_df.iterrows():
            lab = str(r.get("__month_label", "")).strip()
            if lab and lab not in seen:
                month_order.append(lab)
                seen.add(lab)

        month_order = list(reversed(month_order))

        with c_mo:
            selected_month = st.selectbox("MOPR Month", options=["All"] + month_order, index=0, key="mopr_month_select")

        df_filt = df_work.copy()
        if selected_fy != "All":
            df_filt = df_filt[df_filt["__fy"] == selected_fy]
        if selected_month != "All":
            df_filt = df_filt[df_filt["__month_label"] == selected_month]

        df_filt = (
            df_filt.sort_values([dept_col, "__date"], na_position="first")
                   .groupby(dept_col, as_index=False)
                   .tail(1)
        )

    else:
        df_filt = df_work.copy()

        if month_col:
            df_filt["__month_label"] = df_filt[month_col].astype(str).str.strip()
        if fy_col:
            df_filt["__fy"] = df_filt[fy_col].astype(str).str.strip()

        fy_vals = []
        month_vals = []
        if fy_col:
            fy_vals = sorted([v for v in df_filt["__fy"].unique().tolist() if str(v).strip()], reverse=True)
        if month_col:
            month_vals = sorted([v for v in df_filt["__month_label"].unique().tolist() if str(v).strip()], reverse=True)

        if fy_vals or month_vals:
            c_fy, c_mo, _sp = st.columns([1.2, 1.3, 3.5])
            with c_fy:
                if fy_vals:
                    selected_fy = st.selectbox("Financial Year", options=["All"] + fy_vals, index=0, key="mopr_fy_select")
            with c_mo:
                if month_vals:
                    selected_month = st.selectbox("MOPR Month", options=["All"] + month_vals, index=0, key="mopr_month_select")

            if fy_col and selected_fy != "All":
                df_filt = df_filt[df_filt["__fy"] == selected_fy]
            if month_col and selected_month != "All":
                df_filt = df_filt[df_filt["__month_label"] == selected_month]

        df_filt = df_filt.drop_duplicates(subset=[dept_col], keep="last")

    # ---------- Build items ----------
    rows = []
    for _, r in df_filt.iterrows():
        d = str(r.get(dept_col, "")).strip()
        u = extract_url(r.get(url_col, ""))
        if d:
            rows.append((d, u))

    rows.sort(key=lambda x: x[0].lower())

    if not rows:
        st.info("No departments found for the selected Month/FY. Try another selection or check MOPR data.")
        if st.button("⬅️ Back to Dashboard", key="mopr_back_nodata"):
            st.session_state.main_view = "dashboard"
        return

    # --- layout (unchanged) ---
    n = len(rows)
    if n <= 12:
        size = 860;  btn_w, btn_h = 140, 44; radii_f = [0.38]
    elif n <= 30:
        size = 920;  btn_w, btn_h = 130, 42; radii_f = [0.30, 0.52]
    elif n <= 55:
        size = 980;  btn_w, btn_h = 120, 40; radii_f = [0.26, 0.43, 0.60]
    else:
        size = 1080; btn_w, btn_h = 118, 38; radii_f = [0.22, 0.36, 0.52, 0.68]

    center = size // 2
    radii = [int(size * r) for r in radii_f]
    rings = len(radii)

    weights = [r for r in radii]
    total_w = sum(weights)
    base = [max(0, int(round(n * (w / total_w)))) for w in weights]
    diff = n - sum(base)
    if diff != 0:
        order = sorted(range(rings), key=lambda i: weights[i], reverse=(diff > 0))
        k = 0
        while diff != 0:
            i = order[k % rings]
            if diff > 0:
                base[i] += 1; diff -= 1
            else:
                if base[i] > 0:
                    base[i] -= 1; diff += 1
            k += 1
    ring_counts = base

    nodes_html, lines = [], []
    idx = 0

    for r_idx, count in enumerate(ring_counts):
        if count <= 0:
            continue
        radius = radii[r_idx]
        angle_offset = (math.pi / max(count, 1)) * (r_idx % 2)

        for j in range(count):
            if idx >= n:
                break
            dept, url = rows[idx]; idx += 1

            angle = angle_offset + 2 * math.pi * j / max(count, 1)
            x = center + int(radius * math.cos(angle)) - btn_w // 2
            y = center + int(radius * math.sin(angle)) - btn_h // 2
            x = max(6, min(size - btn_w - 6, x))
            y = max(6, min(size - btn_h - 6, y))
            cx = x + btn_w / 2
            cy = y + btn_h / 2

            if url:
                lines.append(
                    f'<line x1="{center}" y1="{center}" x2="{int(cx)}" y2="{int(cy)}" '
                    f'stroke="url(#grad)" stroke-width="2.5" stroke-linecap="round" opacity="0.9" />'
                )
            else:
                lines.append(
                    f'<line x1="{center}" y1="{center}" x2="{int(cx)}" y2="{int(cy)}" '
                    f'stroke="#b9d7ff" stroke-width="2" stroke-linecap="round" '
                    f'stroke-dasharray="6 6" opacity="0.45" />'
                )

            label_html = html.escape(dept)
            common_css = (
                f"position:absolute; left:{x}px; top:{y}px; border-radius:18px;"
                f"display:inline-flex; align-items:center; justify-content:center;"
                f"box-sizing:border-box; width:{btn_w}px; height:{btn_h}px; padding:0 18px;"
                "overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"
                "box-shadow:0 2px 12px #8fd3fe60; font-weight:900; z-index:1; text-align:center;"
            )

            if url:
                safe_url = url.replace('"', '%22')
                nodes_html.append(
                    f'<a href="{safe_url}" target="_blank" rel="noopener" title="{label_html}" '
                    f'style="{common_css} background:linear-gradient(90deg,#299bff 10%, #55e386 90%); '
                    f'color:#000; text-decoration:none;">{label_html}</a>'
                )
            else:
                nodes_html.append(
                    f'<div aria-disabled="true" title="{label_html}" '
                    f'style="{common_css} background:linear-gradient(90deg,#e3f4ff 10%, #e9ffe4 90%); '
                    f'color:#2056b5; opacity:0.65; cursor:not-allowed; user-select:none;">{label_html}</div>'
                )

    edges_svg = (
        f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" '
        f'style="position:absolute; left:0; top:0; z-index:0; pointer-events:none;">'
        f'<defs>'
        f'  <linearGradient id="grad" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="{size}" y2="0">'
        f'    <stop offset="10%" stop-color="#299bff"/>'
        f'    <stop offset="90%" stop-color="#55e386"/>'
        f'  </linearGradient>'
        f'</defs>'
        f'{"".join(lines)}'
        f'</svg>'
    )

    nodes_layer = "".join(nodes_html)

    html_block = (
        f'<div style="position:relative; width:{size}px; height:{size}px; margin:10px auto 20px auto;'
        'background:#ffffff; border-radius:16px; box-shadow:0 8px 32px #00000011;">'
        '<div style="position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);'
        'padding:12px 20px; border-radius:18px; background:#f4e7da; color:#2056b5; font-weight:900;'
        'box-shadow:0 2px 12px #8fd3fe60;">MOPR</div>'
        f'{edges_svg}{nodes_layer}'
        '</div>'
    )

    st.components.v1.html(html_block, height=size + 60, scrolling=False)

    legend_html = """
     <div style='display:flex;gap:18px;justify-content:center;align-items:center;margin:6px 0 14px 0;flex-wrap:wrap;'>
      <span style="display:inline-flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:18px;height:12px;border-radius:6px;background:linear-gradient(90deg,#299bff 10%, #55e386 90%);box-shadow:0 1px 6px #8fd3fe60;"></span>
        <span style="color:#2056b5;font-weight:700;">Clickable (PPT available)</span>
      </span>
      <span style="display:inline-flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:18px;height:12px;border-radius:6px;background:linear-gradient(90deg,#e3f4ff 10%, #e9ffe4 90%);box-shadow:0 1px 6px #8fd3fe60;opacity:0.85;"></span>
        <span style="color:#2056b5;font-weight:700;">No PPT yet</span>
      </span>
     </div>
    """

    st.markdown(legend_html, unsafe_allow_html=True)

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
def load_credentials():
    admins = {}
    viewers = {}

    # 1) Streamlit secrets
    try:
        admins = dict(st.secrets.get("admins", {}))
        viewers = dict(st.secrets.get("viewers", {}))
    except Exception:
        admins = {}
        viewers = {}

    # 2) Optional env fallback
    if not admins:
        raw = os.getenv("CHANDRAGUPTA_ADMINS_JSON", "").strip()
        if raw:
            try:
                admins = json.loads(raw)
            except Exception:
                admins = {}

    if not viewers:
        raw = os.getenv("CHANDRAGUPTA_VIEWERS_JSON", "").strip()
        if raw:
            try:
                viewers = json.loads(raw)
            except Exception:
                viewers = {}

    return admins, viewers


ADMIN_USERS, VIEWERS = load_credentials()

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
            sanitize_df_for_excel(filtered_df2).to_excel(excel_buffer, index=False)

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
    sanitize_df_for_excel(filtered_df2).to_excel(excel_buffer, index=False)

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
    sanitize_df_for_excel(filtered_df).to_excel(excel_buffer, index=False)

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
