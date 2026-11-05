import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io as _io
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==============================================================================
# 1. PAGE CONFIG & CSS
# ==============================================================================
st.set_page_config(page_title="Saint Lucia Revenue Dashboard", layout="wide", page_icon="🇱🇨", initial_sidebar_state="expanded")

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1A3B47; color: white; }
    [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] label { color: white !important; }
    .kpi-card { background-color: white; border-radius: 10px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; border-top: 5px solid #00B4D8; margin-bottom: 20px; }
    .kpi-title { color: #6C757D; font-size: 14px; font-weight: 600; text-transform: uppercase; margin-bottom: 10px; }
    .kpi-value { color: #1A3B47; font-size: 28px; font-weight: bold; margin-bottom: 5px; }
    .kpi-delta { font-size: 14px; font-weight: 500; }
    .delta-pos { color: #2E8B57; }
    .delta-neg { color: #DC3545; }
    .narrative-box { background-color: #f8f9fa; border-left: 5px solid #00B4D8; padding: 15px; margin-bottom: 20px; border-radius: 5px; font-size: 16px; }
    .snapshot-badge { display: inline-block; background-color: #00B4D8; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .snapshot-badge-warn { display: inline-block; background-color: #F59E0B; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .integrity-ok { background-color: #D4EDDA; border: 1px solid #28A745; padding: 10px; border-radius: 5px; font-size: 13px; margin-bottom: 10px; }
    .integrity-warn { background-color: #FFF3CD; border: 1px solid #FFC107; padding: 10px; border-radius: 5px; font-size: 13px; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. CONSTANTS
# ==============================================================================
FISCAL_MONTHS = ["April", "May", "June", "July", "August", "September",
                 "October", "November", "December", "January", "February", "March"]
MONTH_ALIASES = {
    "APR": "April", "APRIL": "April", "MAY": "May", "JUN": "June", "JUNE": "June",
    "JUL": "July", "JULY": "July", "AUG": "August", "AUGUST": "August",
    "SEP": "September", "SEPT": "September", "SEPTEMBER": "September",
    "OCT": "October", "OCTOBER": "October", "NOV": "November", "NOVEMBER": "November",
    "DEC": "December", "DECEMBER": "December", "JAN": "January", "JANUARY": "January",
    "FEB": "February", "FEBRUARY": "February", "MAR": "March", "MARCH": "March",
}
TAX_HEADS = {
    "3111 001": "Personal Income Tax", "3111 002": "Arrears on PIT",
    "3112 001": "Corporate Income Tax", "3112 002": "Arrears on CIT",
    "3113 001": "Withholding Tax", "3113 002": "Arrears on Withholding",
    "3131 001": "Property Tax", "3141 001": "Value Added Tax (VAT)",
    "3141 004": "Hotel Accommodation Tax", "3141 005": "Domestic Sales Tax on Travel",
    "3141 006": "Health & Citizen Security Levy", "3142 005": "Insurance Premium Tax",
    "3142 006": "Stamp Duty - Inland Revenue", "3143 010": "Telecommunication Class Licence",
    "3151 009": "Passenger Facility Fee", "3422 008": "Collection Fees - Towns & Villages",
}
QUARTER_END_MONTHS = {"June": "Q1", "September": "Q2 (Mid-Year)", "December": "Q3", "March": "Q4 (Year-End)"}

# ==============================================================================
# 3. HELPERS
# ==============================================================================
def normalize_code(code):
    if pd.isna(code): return ""
    return re.sub(r'\s+', ' ', str(code).strip())

def find_month_in_text(text):
    if not text: return None
    text_upper = str(text).upper()
    for alias in sorted(MONTH_ALIASES.keys(), key=len, reverse=True):
        if re.search(r'\b' + alias + r'\b', text_upper):
            return MONTH_ALIASES[alias]
    return None

def safe_number(val):
    try:
        if pd.isna(val): return 0.0
        return float(str(val).replace(",", "").replace("$", "").strip())
    except:
        return 0.0

def get_quarter(month):
    if month in ["April", "May", "June"]: return "Q1 (Apr-Jun)"
    if month in ["July", "August", "September"]: return "Q2 (Jul-Sep)"
    if month in ["October", "November", "December"]: return "Q3 (Oct-Dec)"
    if month in ["January", "February", "March"]: return "Q4 (Jan-Mar)"
    return "Unknown"

def month_order(m):
    return FISCAL_MONTHS.index(m) if m in FISCAL_MONTHS else 99

# ==============================================================================
# 4. PARSERS
# ==============================================================================

# --- 4a. Monthly Actuals (from "Actual Monthly Revenue" sheet) — for TRENDS only ---
@st.cache_data
def parse_all_main_data(file_payload):
    all_tax, diagnostics = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            rev_sheet = next((s for s in xls.sheet_names if "actual monthly revenue" in s.lower()), None)
            if not rev_sheet: continue
            month = find_month_in_text(filename) or find_month_in_text(rev_sheet)
            if not month: continue
            df = pd.read_excel(file_obj, sheet_name=rev_sheet, header=None)
            rows_in_file = 0
            for i in range(5, min(35, len(df))):
                code = df.iloc[i, 0]
                name = df.iloc[i, 1]
                if pd.isna(code) or pd.isna(name): continue
                code_norm = normalize_code(code)
                if not re.match(r'^\d{4}\s\d{3}$', code_norm): continue
                if "total" in str(name).lower(): continue
                approved = safe_number(df.iloc[i, 2])
                monthly_actual = sum(safe_number(df.iloc[i, c]) for c in range(4, 15))
                ytd_actual = safe_number(df.iloc[i, 15]) or monthly_actual
                all_tax.append({
                    "Month": month, "Quarter": get_quarter(month),
                    "Code": code_norm, "Account": str(name).strip(),
                    "Approved_Annual": approved, "Monthly_Actual": monthly_actual, "YTD_Actual": ytd_actual
                })
                rows_in_file += 1
            diagnostics.append({"file": filename, "month": month, "rows": rows_in_file})
        except Exception as e:
            diagnostics.append({"file": filename, "month": None, "rows": 0, "error": str(e)})
    return pd.DataFrame(all_tax), pd.DataFrame(diagnostics)

# --- 4b. Summary Snapshot (from "Revenue Summary" sheet) — AUTHORITATIVE for YTD actuals ---
@st.cache_data
def parse_summary_snapshot(file_payload):
    """
    Reads ACTUAL REVENUE TO DATE (col E / idx 4) and VARIANCE (col N / idx 13)
    from the Revenue Summary sheet. Filters by explicit code regex only,
    NOT by "total" text, to preserve all individual tax rows.
    """
    all_snap, diagnostics = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month:
                for s in xls.sheet_names:
                    month = find_month_in_text(s)
                    if month: break
            if not month:
                diagnostics.append({"file": filename, "error": "month not detected"})
                continue

            summary_sheet = next((s for s in xls.sheet_names if "revenue summary" in s.lower()), None)
            if not summary_sheet:
                summary_sheet = next((s for s in xls.sheet_names if "summary" in s.lower()), None)
            if not summary_sheet:
                diagnostics.append({"file": filename, "error": "no summary sheet"})
                continue

            df = pd.read_excel(file_obj, sheet_name=summary_sheet, header=None)

            # Find the header row by locating "TO DATE"
            header_row = None
            for r in range(15):
                row_str = " ".join(str(v).upper() for v in df.iloc[r].values if pd.notna(v))
                if "TO DATE" in row_str:
                    header_row = r
                    break
            if header_row is None:
                header_row = 4

            rows = 0
            for i in range(header_row + 1, len(df)):
                raw_code = df.iloc[i, 0]
                desc = df.iloc[i, 1]
                if pd.isna(raw_code) or pd.isna(desc):
                    continue
                code = normalize_code(raw_code)
                # ONLY accept explicit 4-digit + space + 3-digit codes
                # This excludes "ACTIVITY TOTAL", "Total Personal Income Tax", etc.
                if not re.match(r'^\d{4}\s\d{3}$', code):
                    continue

                approved = safe_number(df.iloc[i, 2])
                actual_month = safe_number(df.iloc[i, 3]) if 3 < len(df.columns) else 0
                actual_ytd = safe_number(df.iloc[i, 4]) if 4 < len(df.columns) else 0
                variance = safe_number(df.iloc[i, 13]) if 13 < len(df.columns) else 0

                all_snap.append({
                    "Month": month, "Quarter": get_quarter(month),
                    "Code": code, "Account": str(desc).strip(),
                    "Approved_Annual": approved,
                    "Actual_Month": actual_month,
                    "Actual_YTD": actual_ytd,
                    "Snapshot_Variance": variance,
                })
                rows += 1
            diagnostics.append({"file": filename, "month": month, "rows": rows})
        except Exception as e:
            diagnostics.append({"file": filename, "error": str(e)})
    return pd.DataFrame(all_snap), pd.DataFrame(diagnostics)

# --- 4c. Calendarized Projections ---
@st.cache_data
def parse_calendarized_projections(file_payload):
    """
    Reads monthly projection columns (D,E,F,H,I,J,L,M,N,P,Q,R) from Projections file.
    Skips formula columns (G,K,O,S quarter subtotals and T annual total).
    """
    if not file_payload:
        return pd.DataFrame(), {"error": "no file"}
    filename, file_bytes = file_payload
    diagnostics = {"file": filename, "rows_parsed": 0, "error": None}
    try:
        file_obj = _io.BytesIO(file_bytes)
        xls = pd.ExcelFile(file_obj)
        sheet = next((s for s in xls.sheet_names if "projection" in s.lower()), xls.sheet_names[0])
        df = pd.read_excel(file_obj, sheet_name=sheet, header=None)

        MONTH_COL_MAP = {
            "April": 3, "May": 4, "June": 5,
            "July": 7, "August": 8, "September": 9,
            "October": 11, "November": 12, "December": 13,
            "January": 15, "February": 16, "March": 17,
        }

        all_proj = []
        for i in range(7, min(55, len(df))):
            raw_code = df.iloc[i, 0]
            desc = df.iloc[i, 1]
            if pd.isna(raw_code) or pd.isna(desc):
                continue
            code = normalize_code(raw_code)
            if not re.match(r'^\d{4}\s\d{3}$', code):
                continue

            for month, col_idx in MONTH_COL_MAP.items():
                raw_val = df.iloc[i, col_idx] if col_idx < len(df.columns) else 0
                # If value is a formula string like "=SUM(D9:F9)", treat as 0
                # (those are quarter subtotals, which we don't need here anyway)
                val = safe_number(raw_val)
                all_proj.append({
                    "Code": code,
                    "Account": str(desc).strip(),
                    "Month": month,
                    "Projection": val,
                })

        result = pd.DataFrame(all_proj)
        diagnostics["rows_parsed"] = len(result)
        return result, diagnostics
    except Exception as e:
        diagnostics["error"] = str(e)
        return pd.DataFrame(), diagnostics

# --- 4d. Daily Collections & Offices ---
@st.cache_data
def parse_daily_collections(file_payload):
    all_daily, all_offices = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            daily_sheet = next((s for s in xls.sheet_names if "daily collection" in s.lower()), None)
            if not daily_sheet or not month: continue
            df = pd.read_excel(file_obj, sheet_name=daily_sheet, header=None)
            for i in range(5, 35):
                date_val = df.iloc[i, 0]
                if pd.isna(date_val) or not isinstance(date_val, (datetime, pd.Timestamp)): continue
                total_col = safe_number(df.iloc[i, 20])
                if total_col > 0:
                    all_daily.append({"Month": month, "Date": pd.to_datetime(date_val).date(),
                                      "Day": pd.to_datetime(date_val).day, "Total": total_col})
            for i in range(30, 55):
                office_name = df.iloc[i, 0]
                office_total = safe_number(df.iloc[i, 20])
                if pd.notna(office_name) and isinstance(office_name, str):
                    s = office_name.strip()
                    if any(x in s for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        all_offices.append({"Month": month, "Office": s, "Total": office_total})
                    elif "Sub Total" in s:
                        all_offices.append({"Month": month, "Office": "Castries (Sub Total)", "Total": office_total})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(all_offices)

# --- 4e. ISIC Deep ---
@st.cache_data
def parse_isic_deep(file_payload):
    all_isic = []
    target_sheets = ["PIT", "CIT", "VAT", "HAT", "Travel Tax", "HCSL", "Insurance Premium", "Stamp Duty"]
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month: continue
            for target in target_sheets:
                sheet = next((s for s in xls.sheet_names if target.lower() in s.lower()), None)
                if not sheet: continue
                df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
                header_row, isic_col, current_col, arrears_col = None, None, None, None
                for r in range(15):
                    row_vals = [str(v).lower() for v in df.iloc[r].values if pd.notna(v)]
                    if any("isic description" in v for v in row_vals):
                        header_row = r
                        for c, val in enumerate(df.iloc[r].values):
                            v_low = str(val).lower()
                            if "isic description" in v_low: isic_col = c
                            if "current" in v_low and current_col is None: current_col = c
                            if "arrears" in v_low and arrears_col is None: arrears_col = c
                        break
                if header_row is None or isic_col is None: continue
                for i in range(header_row + 1, len(df)):
                    sector = df.iloc[i, isic_col]
                    if pd.isna(sector) or "grand total" in str(sector).lower(): break
                    cur = safe_number(df.iloc[i, current_col]) if current_col is not None else 0
                    arr = safe_number(df.iloc[i, arrears_col]) if arrears_col is not None else 0
                    if cur > 0 or arr > 0:
                        all_isic.append({"Month": month, "Tax_Type": sheet, "ISIC_Sector": str(sector).strip(),
                                         "Current": cur, "Arrears": arr, "Total": cur + arr})
        except Exception:
            pass
    return pd.DataFrame(all_isic)

# ==============================================================================
# 5. SNAPSHOT SELECTION
# ==============================================================================
def select_authoritative_snapshot(df_snap, selected_months, analysis_mode):
    if df_snap.empty: return None, None
    filtered = df_snap[df_snap['Month'].isin(selected_months)]
    if filtered.empty: return None, None
    if analysis_mode == "YTD (Latest Summary)":
        latest = sorted(filtered['Month'].unique(), key=month_order)[-1]
        return filtered[filtered['Month'] == latest], latest
    elif analysis_mode == "Quarter-End Snapshot":
        available = [m for m in filtered['Month'].unique() if m in QUARTER_END_MONTHS]
        if not available: return None, None
        latest_qe = sorted(available, key=month_order)[-1]
        return filtered[filtered['Month'] == latest_qe], latest_qe
    else:
        latest = sorted(filtered['Month'].unique(), key=month_order)[-1]
        return filtered[filtered['Month'] == latest], latest

# ==============================================================================
# 6. SIDEBAR (files only first)
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    calib_file = st.file_uploader("Calendarized Projections (.xlsx)", type=["xlsx"], accept_multiple_files=False,
                                  help="Optional — upload the annual projections file to get authoritative YTD targets")

# ==============================================================================
# 7. MAIN
# ==============================================================================
st.title("🇱🇨 Revenue Intelligence Dashboard")

if not main_files:
    st.info("👈 Upload your Monthly Revenue Reports in the sidebar to begin.")
    st.stop()

# --- Parse everything ---
main_payload = [(f.name, f.getvalue()) for f in main_files]
df_tax, diag_main = parse_all_main_data(main_payload)          # monthly slices (trends)
df_snap, diag_snap = parse_summary_snapshot(main_payload)       # YTD actuals (authoritative)
df_daily, df_offices = parse_daily_collections(main_payload)

df_isic = pd.DataFrame()
if isic_files:
    isic_payload = [(f.name, f.getvalue()) for f in isic_files]
    df_isic = parse_isic_deep(isic_payload)

df_calib = pd.DataFrame()
calib_diag = {}
if calib_file:
    calib_payload = (calib_file.name, calib_file.getvalue())
    df_calib, calib_diag = parse_calendarized_projections(calib_payload)

# --- Diagnostics ---
with st.expander("🔍 Data Diagnostics", expanded=df_snap.empty):
    st.write("**Monthly slices (trends):**", len(df_tax), "rows")
    st.dataframe(diag_main, use_container_width=True)
    st.write("**Summary snapshots (YTD actuals):**", len(df_snap), "rows")
    st.dataframe(diag_snap, use_container_width=True)
    if calib_file:
        if calib_diag.get("error"):
            st.error(f"Projections parse error: {calib_diag['error']}")
        else:
            st.success(f"✓ Projections loaded: {calib_diag['rows_parsed']} rows")
    else:
        st.info("ℹ️ No calendarized projections uploaded. YTD target will be inferred from the Revenue Summary's variance column.")

if df_snap.empty:
    st.error("Could not parse Summary snapshots. Verify your files contain a 'Revenue Summary' sheet.")
    st.stop()

# --- Build slicer options ---
available_months = sorted(df_snap['Month'].unique().tolist(), key=month_order)
available_quarters = sorted(df_snap['Quarter'].unique().tolist())
available_codes = sorted(df_snap['Code'].unique().tolist())
tax_display_options = [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in available_codes]
display_to_code = {f"{c} — {TAX_HEADS.get(c, 'Unknown')}": c for c in available_codes}

# --- Sidebar filters ---
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🎯 Analysis Mode")
    analysis_mode = st.radio(
        "Snapshot driver",
        ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.markdown("### 🎛️ Filters")

    # Month
    st.markdown("**Month**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="m_all", use_container_width=True):
        st.session_state["month_f"] = available_months
    if c2.button("None", key="m_none", use_container_width=True):
        st.session_state["month_f"] = []
    selected_months = st.multiselect("Months", available_months, default=available_months,
                                     key="month_f", label_visibility="collapsed")

    # Quarter
    st.markdown("**Quarter**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="q_all", use_container_width=True):
        st.session_state["qtr_f"] = available_quarters
    if c2.button("None", key="q_none", use_container_width=True):
        st.session_state["qtr_f"] = []
    selected_quarters = st.multiselect("Quarters", available_quarters, default=available_quarters,
                                       key="qtr_f", label_visibility="collapsed")

    # Tax Head
    st.markdown("**Tax Head**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="t_all", use_container_width=True):
        st.session_state["tax_f"] = tax_display_options
    if c2.button("None", key="t_none", use_container_width=True):
        st.session_state["tax_f"] = []
    selected_taxes_display = st.multiselect("Tax Heads", tax_display_options, default=tax_display_options,
                                            key="tax_f", label_visibility="collapsed")

    # Cash Basis
    st.markdown("**Cash Basis**")
    cash_basis = st.radio("Focus", ["All", "Current Only", "Arrears Only"],
                          horizontal=True, label_visibility="collapsed")

# --- Apply filters ---
if not selected_months:
    st.warning("⚠️ No months selected. Click 'All' under Month.")
    st.stop()
if not selected_quarters:
    st.warning("⚠️ No quarters selected. Click 'All' under Quarter.")
    st.stop()
if not selected_taxes_display:
    st.warning("⚠️ No tax heads selected. Click 'All' under Tax Head.")
    st.stop()

selected_codes = [display_to_code[d] for d in selected_taxes_display if d in display_to_code]

# --- Filter monthly slices (for trends) ---
df_f = df_tax[
    df_tax['Month'].isin(selected_months) &
    df_tax['Quarter'].isin(selected_quarters) &
    df_tax['Code'].isin(selected_codes)
].copy()

# --- Snapshot for YTD ---
snapshot_df, snapshot_month = select_authoritative_snapshot(df_snap, selected_months, analysis_mode)
if snapshot_df is None or snapshot_df.empty:
    st.error(f"No {analysis_mode} snapshot available.")
    st.stop()

# Restrict snapshot to selected codes and quarters
snapshot_df = snapshot_df[
    snapshot_df['Code'].isin(selected_codes) &
    snapshot_df['Quarter'].isin(selected_quarters)
].copy()

if snapshot_df.empty:
    st.error("Selected filters exclude all rows from the snapshot.")
    st.stop()

# --- YTD ACTUALS from snapshot ---
ytd_actuals = snapshot_df[['Code', 'Account', 'Actual_YTD', 'Actual_Month', 'Approved_Annual', 'Snapshot_Variance']].copy()

# --- YTD PROJECTION: calendarized (authoritative) or fallback ---
if not df_calib.empty:
    months_through = FISCAL_MONTHS[:month_order(snapshot_month) + 1]
    proj_window = df_calib[
        df_calib['Month'].isin(months_through) &
        df_calib['Code'].isin(selected_codes)
    ]
    ytd_proj = proj_window.groupby('Code').agg(YTD_Projection=('Projection', 'sum')).reset_index()
    projection_source = f"Calendarized (Apr–{snapshot_month}: {len(months_through)} months)"
    fallback_used = False
else:
    # Fallback: back out projection from Snapshot variance
    ytd_proj = snapshot_df[['Code', 'Snapshot_Variance']].copy()
    ytd_proj['YTD_Projection'] = snapshot_df['Actual_YTD'].values - snapshot_df['Snapshot_Variance'].values
    ytd_proj = ytd_proj[['Code', 'YTD_Projection']]
    projection_source = "Fallback (from Summary variance column)"
    fallback_used = True

# --- Merge and compute final variance ---
ytd = ytd_actuals.merge(ytd_proj, on='Code', how='left')
ytd['YTD_Projection'] = ytd['YTD_Projection'].fillna(0)
ytd['Variance'] = ytd['Actual_YTD'] - ytd['YTD_Projection']

total_ytd = ytd['Actual_YTD'].sum()
total_proj = ytd['YTD_Projection'].sum()
total_var = ytd['Variance'].sum()

# --- INTEGRITY CHECK: compare against Summary's own variance ---
snapshot_var_sum = snapshot_df['Snapshot_Variance'].sum()
delta_vs_summary = total_var - snapshot_var_sum

# Show diagnostic
st.caption(
    f"🔍 **Data check:** Σ(Actual_YTD) = ${total_ytd:,.2f} | "
    f"Summary reports ${snapshot_df['Actual_YTD'].sum():,.2f} | "
    f"Σ(Variance) = ${total_var:,.2f} vs Summary ${snapshot_var_sum:,.2f}"
)

if not fallback_used and abs(delta_vs_summary) > 1.0:
    st.markdown(f"""
    <div class="integrity-warn">
        ⚠️ <b>Calendarization discrepancy:</b> Dashboard variance 
        <b>${total_var:,.2f}</b> vs Summary variance 
        <b>${snapshot_var_sum:,.2f}</b> (difference: ${delta_vs_summary:,.2f}).
        This may indicate mid-year projection revisions or a parsing issue.
    </div>
    """, unsafe_allow_html=True)
elif not fallback_used:
    st.markdown(f"""
    <div class="integrity-ok">
        ✓ <b>Integrity check passed:</b> Dashboard variance ${total_var:,.2f} matches Summary ${snapshot_var_sum:,.2f}.
    </div>
    """, unsafe_allow_html=True)

# --- Top/Bottom ---
top_perf = ytd.nlargest(1, 'Variance').iloc[0]
top_short = ytd.nsmallest(1, 'Variance').iloc[0]

is_qe = snapshot_month in QUARTER_END_MONTHS
badge_class = "snapshot-badge" if is_qe else "snapshot-badge-warn"
badge_text = f"✓ Quarter-End: {snapshot_month}" if is_qe else f"Latest: {snapshot_month}"

# --- Narrative Box ---
st.markdown(f"""
<div class="narrative-box">
    <b>📊 Executive Summary</b>
    <span class="{badge_class}">{badge_text}</span><br><br>
    <b>Data Source:</b> <i>{snapshot_month} 2026 Revenue Summary</i> (Mode: {analysis_mode})<br>
    <b>Projection Source:</b> {projection_source}<br>
    {'<span style="color:#F59E0B">⚠️ Upload the Calendarized Projections file for authoritative targets.</span><br>' if fallback_used else ''}
    <b>YTD Actual:</b> ${total_ytd:,.2f} | <b>YTD Projection:</b> ${total_proj:,.2f}<br>
    <b>Net {'Overage' if total_var > 0 else 'Shortage'}:</b> ${abs(total_var):,.2f} ({(total_var/total_proj*100) if total_proj>0 else 0:.2f}%)<br><br>
    <b>Top Driver:</b> {top_perf['Account']} (+${top_perf['Variance']:,.2f})<br>
    <b>Main Drag:</b> {top_short['Account']} (${top_short['Variance']:,.2f})
</div>
""", unsafe_allow_html=True)

# --- KPI Cards ---
c1, c2, c3 = st.columns(3)
c1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Actual</div><div class="kpi-value">${total_ytd:,.0f}</div><div class="kpi-delta">Source: {snapshot_month}</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Projection</div><div class="kpi-value">${total_proj:,.0f}</div><div class="kpi-delta">{"Calendarized" if not fallback_used else "Derived"}</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_var > 0 else "#DC3545"}">${total_var:,.0f}</div></div>', unsafe_allow_html=True)

# --- Tabs ---
tab1, tab2, tab3, tab4 = st.tabs(["📈 Macro Trends", "📅 Daily & Office", "🏭 Sector (ISIC)", "📝 Reports"])

with tab1:
    st.markdown("### Tax Head Composition Over Time")
    if not df_f.empty:
        mc = df_f.groupby(['Month', 'Account'])['Monthly_Actual'].sum().reset_index()
        mc['Order'] = mc['Month'].map(month_order)
        mc = mc.sort_values('Order')
        st.plotly_chart(px.area(mc, x='Month', y='Monthly_Actual', color='Account', groupnorm='fraction'), use_container_width=True)

    st.markdown("### Month-over-Month Growth Rates")
    if not df_f.empty:
        mom = df_f.groupby(['Month', 'Account'])['Monthly_Actual'].sum().unstack(fill_value=0)
        mom_pct = mom.pct_change().replace([float('inf'), float('-inf')], 0) * 100
        st.dataframe(mom_pct.style.format("{:.1f}%"), use_container_width=True)

with tab2:
    if not df_daily.empty:
        df_d_f = df_daily[df_daily['Month'].isin(selected_months)]
        if not df_d_f.empty:
            st.markdown("### Daily Collection Heatmap")
            heatmap = df_d_f.pivot_table(index='Month', columns='Day', values='Total', aggfunc='sum').fillna(0)
            st.plotly_chart(
                px.imshow(heatmap, aspect="auto", color_continuous_scale="Blues",
                          labels=dict(x="Day of Month", y="Month", color="Revenue")),
                use_container_width=True
            )

            st.markdown("### Filing Deadline Spikes")
            df_d_f = df_d_f.copy()
            df_d_f['Deadline'] = df_d_f['Day'].apply(
                lambda x: 'Deadline Day' if x in [15, 21, 25, 28, 30, 31] else 'Normal Day'
            )
            st.plotly_chart(
                px.box(df_d_f, x='Deadline', y='Total', color='Deadline', points="all"),
                use_container_width=True
            )

    if not df_offices.empty:
        dfo = df_offices[df_offices['Month'].isin(selected_months)]
        if not dfo.empty:
            st.markdown("### Office Performance")
            s = dfo.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
            st.plotly_chart(
                px.bar(s, x='Office', y='Total', color='Total', color_continuous_scale='Blues'),
                use_container_width=True
            )
