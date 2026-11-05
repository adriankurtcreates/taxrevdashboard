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
from pathlib import Path

# ==============================================================================
# PAGE CONFIG
# ==============================================================================
st.set_page_config(page_title="Saint Lucia Revenue Dashboard", layout="wide", page_icon="🇱🇨")

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1A3B47; color: white; }
    [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] label { color: white !important; }
    .kpi-card { background-color: white; border-radius: 10px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; border-top: 5px solid #00B4D8; margin-bottom: 20px; }
    .kpi-title { color: #6C757D; font-size: 14px; font-weight: 600; text-transform: uppercase; margin-bottom: 10px; }
    .kpi-value { color: #1A3B47; font-size: 28px; font-weight: bold; margin-bottom: 5px; }
    .kpi-delta { font-size: 14px; font-weight: 500; }
    .narrative-box { background-color: #f8f9fa; border-left: 5px solid #00B4D8; padding: 15px; margin-bottom: 20px; border-radius: 5px; font-size: 16px; }
    .snapshot-badge { display: inline-block; background-color: #00B4D8; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .snapshot-badge-warn { display: inline-block; background-color: #F59E0B; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
</style>
""", unsafe_allow_html=True)

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
    "3111 001": "Personal Income Tax", "3111 002": "Arrears on Personal Income Tax",
    "3112 001": "Corporate Income Tax", "3112 002": "Arrears on Corporate Income Tax",
    "3113 001": "Withholding Income Tax", "3113 002": "Arrears on Withholding Income Tax",
    "3131 001": "Property Tax", "3141 001": "Value Added Tax (VAT)",
    "3141 004": "Hotel Accommodation Tax", "3141 005": "Domestic Sales Tax on Travel",
    "3141 006": "Health and Citizen Security Levy", "3142 005": "Insurance Premium Tax",
    "3142 006": "Stamp Duty - Inland Revenue", "3143 010": "Telecommunication Class Licence",
    "3151 009": "Passenger Facility Fee", "3422 008": "Collection Fees - Towns & Villages",
}
CODE_ALIASES = {"362 - 011": "3422 008", "362-011": "3422 008", "362 011": "3422 008"}
QUARTER_END_MONTHS = {"June": "Q1", "September": "Q2", "December": "Q3", "March": "Q4"}

def normalize_code(code):
    if pd.isna(code): return ""
    n = re.sub(r'\s+', ' ', str(code).strip())
    return CODE_ALIASES.get(n, n)

def find_month_in_text(text):
    if not text: return None
    t = str(text).upper()
    for a in sorted(MONTH_ALIASES.keys(), key=len, reverse=True):
        if re.search(r'\b' + a + r'\b', t):
            return MONTH_ALIASES[a]
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

def fmt_currency(val):
    if val < 0: return f"(${abs(val):,.2f})"
    return f"${val:,.2f}"

def fmt_pct(val):
    if val < 0: return f"({abs(val):.2f}%)"
    return f"{val:.2f}%"

def get_fiscal_context(month):
    q = get_quarter(month)
    if q == "Q1 (Apr-Jun)":
        pos = {"April": "first", "May": "second", "June": "third"}.get(month, "")
        return f"This month represents the {pos} month of the 1st quarter of the fiscal year."
    elif q == "Q2 (Jul-Sep)":
        pos = {"July": "first", "August": "second", "September": "third"}.get(month, "")
        return f"This month represents the {pos} month of the 2nd quarter of the fiscal year."
    elif q == "Q3 (Oct-Dec)":
        pos = {"October": "first", "November": "second", "December": "third"}.get(month, "")
        return f"This month represents the {pos} month of the 3rd quarter of the fiscal year."
    else:
        pos = {"January": "first", "February": "second", "March": "third"}.get(month, "")
        return f"This month represents the {pos} month of the 4th quarter of the fiscal year."

# ==============================================================================
# PARSERS
# ==============================================================================
@st.cache_data
def parse_collection_and_projection(file_payload):
    all_data = []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            sheet = next((s for s in xls.sheet_names if "collection & projection" in s.lower() or "collection and projection" in s.lower()), None)
            if not sheet: continue
            df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            
            header_row = None
            for r in range(20):
                row_str = " ".join(str(v).lower() for v in df.iloc[r].values if pd.notna(v))
                if "projection" in row_str and "collections" in row_str:
                    header_row = r
                    break
            if header_row is None: continue
                
            for i in range(header_row + 1, len(df)):
                raw_code = df.iloc[i, 0]
                if pd.isna(raw_code): continue
                code = normalize_code(raw_code)
                if not re.match(r'^\d{4}\s\d{3}$', code): continue
                desc = df.iloc[i, 1]
                if pd.isna(desc): continue
                
                projection = safe_number(df.iloc[i, 2])
                collections = safe_number(df.iloc[i, 3])
                
                all_data.append({
                    "Month": month, "Code": code, "Account": str(desc).strip(),
                    "Monthly_Projection": projection, "Monthly_Actual": collections,
                    "Monthly_Variance": collections - projection
                })
        except Exception:
            pass
    return pd.DataFrame(all_data)

@st.cache_data
def parse_summary_snapshot(file_payload):
    all_snap, diag = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month:
                for s in xls.sheet_names:
                    month = find_month_in_text(s)
                    if month: break
            if not month: continue
            summary = next((s for s in xls.sheet_names if "revenue summary" in s.lower()), None)
            if not summary:
                summary = next((s for s in xls.sheet_names if "summary" in s.lower()), None)
            if not summary: continue
            df = pd.read_excel(file_obj, sheet_name=summary, header=None)
            header_row = 4
            for r in range(15):
                row_str = " ".join(str(v).upper() for v in df.iloc[r].values if pd.notna(v))
                if "TO DATE" in row_str:
                    header_row = r; break
            for i in range(header_row + 1, len(df)):
                raw = df.iloc[i, 0]; desc = df.iloc[i, 1]
                if pd.isna(raw) or pd.isna(desc): continue
                code = normalize_code(raw)
                if not re.match(r'^\d{4}\s\d{3}$', code): continue
                ytd = safe_number(df.iloc[i, 4]) if 4 < len(df.columns) else 0
                all_snap.append({"Month": month, "Code": code, "Account": str(desc).strip(), "Actual_YTD": ytd})
        except Exception as e:
            pass
    return pd.DataFrame(all_snap)
@st.cache_data
def parse_calendarized(file_payload):
    if not file_payload: return pd.DataFrame(), {}
    filename, file_bytes = file_payload
    try:
        file_obj = _io.BytesIO(file_bytes)
        xls = pd.ExcelFile(file_obj)
        sheet = next((s for s in xls.sheet_names if "projection" in s.lower()), xls.sheet_names[0])
        df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
        MONTH_COLS = {"April": 3, "May": 4, "June": 5, "July": 7, "August": 8, "September": 9,
                      "October": 11, "November": 12, "December": 13, "January": 15, "February": 16, "March": 17}
        all_proj = []
        for i in range(7, min(55, len(df))):
            raw = df.iloc[i, 0]; desc = df.iloc[i, 1]
            if pd.isna(raw) or pd.isna(desc): continue
            code = normalize_code(raw)
            if not re.match(r'^\d{4}\s\d{3}$', code): continue
            for month, col in MONTH_COLS.items():
                val = safe_number(df.iloc[i, col]) if col < len(df.columns) else 0
                all_proj.append({"Code": code, "Account": str(desc).strip(), "Month": month, "Projection": val})
        return pd.DataFrame(all_proj), {}
    except Exception as e:
        return pd.DataFrame(), {"error": str(e)}

@st.cache_data
def parse_daily(file_payload):
    all_daily, all_offices = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            ds = next((s for s in xls.sheet_names if "daily collection" in s.lower()), None)
            if not ds or not month: continue
            df = pd.read_excel(file_obj, sheet_name=ds, header=None)
            for i in range(5, 35):
                dv = df.iloc[i, 0]
                if pd.isna(dv) or not isinstance(dv, (datetime, pd.Timestamp)): continue
                t = safe_number(df.iloc[i, 20])
                if t > 0:
                    all_daily.append({"Month": month, "Day": pd.to_datetime(dv).day, "Total": t})
            for i in range(30, 55):
                on = df.iloc[i, 0]; ot = safe_number(df.iloc[i, 20])
                if pd.notna(on) and isinstance(on, str):
                    s = on.strip()
                    if "Sub Total" in s:
                        all_offices.append({"Month": month, "Office": "Castries", "Total": ot})
                    elif any(x in s for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        all_offices.append({"Month": month, "Office": s, "Total": ot})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(all_offices)

@st.cache_data
def parse_isic(file_payload):
    all_isic = []
    targets = ["PIT", "CIT", "VAT", "HAT", "Travel Tax", "HCSL", "Insurance Premium", "Stamp Duty"]
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month: continue
            for t in targets:
                sheet = next((s for s in xls.sheet_names if t.lower() in s.lower()), None)
                if not sheet: continue
                df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
                hr, ic, cc = None, None, None
                for r in range(15):
                    vals = [str(v).lower() for v in df.iloc[r].values if pd.notna(v)]
                    if any("isic description" in v for v in vals):
                        hr = r
                        for c, val in enumerate(df.iloc[r].values):
                            vl = str(val).lower()
                            if "isic description" in vl: ic = c
                            if "current" in vl and cc is None: cc = c
                        break
                if hr is None or ic is None: continue
                for i in range(hr + 1, len(df)):
                    sector = df.iloc[i, ic]
                    if pd.isna(sector) or "grand total" in str(sector).lower(): break
                    cur = safe_number(df.iloc[i, cc]) if cc is not None else 0
                    if cur > 0:
                        all_isic.append({"Month": month, "Tax_Type": sheet, "ISIC_Sector": str(sector).strip(), "Current": cur})
        except Exception:
            pass
    return pd.DataFrame(all_isic)
def generate_tax_paragraph(actual, projection, variance):
    if projection > 0: pct = (variance / projection) * 100
    else: pct = 0.0
    var_str = fmt_currency(variance)
    pct_str = fmt_pct(pct)
    if variance == 0:
        return (f"Actual collections of {fmt_currency(actual)} matched the projected figure of {fmt_currency(projection)}. There is no variance to report.")
    direction = "exceeded" if variance > 0 else "fell short of"
    size = "an overage" if variance > 0 else "a shortfall"
    sentence = (f"Actual collections of {fmt_currency(actual)} {direction} the projected figure of {fmt_currency(projection)}, representing {size} of {var_str} or {pct_str}. ")
    if projection > 0:
        rate = (actual / projection) * 100
        sentence += f"This corresponds to a collection rate of {rate:.2f}% of the projection for the period."
    return sentence

def build_variance_notes_docx(snap_month, total_actual, total_projection, total_variance, ytd_df, df_isic):
    doc = Document()
    h = doc.add_heading('SAINT LUCIA', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('DEPARTMENT OF FINANCE', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('INLAND REVENUE DEPARTMENT', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph('Revenue Collection Volume 14: Monthly Performance Report'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(f'Explanation of Variances for the Month of {snap_month} 2026'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(14)

    doc.add_heading('Overview', level=1)
    direction = "an overage" if total_variance >= 0 else "a shortfall"
    pct_total = (total_variance / total_projection) * 100 if total_projection > 0 else 0
    overview_text = (f"For the month of {snap_month} 2026, the Inland Revenue Department projected total revenue of {fmt_currency(total_projection)}. Actual collections amounted to {fmt_currency(total_actual)}, resulting in {direction} of {fmt_currency(abs(total_variance))} or {fmt_pct(abs(pct_total))} {'above' if total_variance >= 0 else 'below'} projection.")
    doc.add_paragraph(overview_text)
    doc.add_paragraph(get_fiscal_context(snap_month))
    doc.add_paragraph(f"The aggregate performance for {snap_month} shows {'positive' if total_variance >= 0 else 'negative'} net variance. The detailed distribution of performance across individual tax heads is presented in the Variance Analysis section below.")

    doc.add_heading('Variance Analysis by Tax Type', level=1)
    doc.add_paragraph("The performances by individual tax type are outlined below:")

    for _, row in ytd_df.sort_values('Code').iterrows():
        code = row['Code']; account = row['Account']; actual = row['Actual_YTD']; projection = row['YTD_Projection']; variance = row['Variance']
        pct = (variance / projection * 100) if projection > 0 else 0.0
        doc.add_heading(f"{code} - {account}", level=2)
        line = (f"Projection: {fmt_currency(projection)} | Actual: {fmt_currency(actual)} | Variance: {fmt_currency(variance)} ({fmt_pct(pct)})")
        p = doc.add_paragraph(line); p.runs[0].bold = True
        doc.add_paragraph(generate_tax_paragraph(actual, projection, variance))
        
        if code == "3131 001": doc.add_paragraph("Note: Due to the ongoing moratorium on residential property tax, current payments apply exclusively to commercial properties. Arrears collections apply to both.")
        elif code == "3141 004": doc.add_paragraph("Note: The Hotel Accommodation Tax remains an active registry line for the collection and settlement of legacy arrears.")
        elif code == "3141 006": doc.add_paragraph("Note: The Health and Citizen Security Levy (HCSL) was implemented in March 2023; no historical data exists prior to that date.")
        elif code == "3142 006": doc.add_paragraph("Structural Insight: The top-performing sectors listed for Stamp Duty may not be a true indicator of which economic sectors are generating the underlying transactions. This is because the ISIC code used to log the payment is routinely assigned to the law firm or legal practitioner paying the tax on behalf of their clients.")
        elif code == "3422 008": doc.add_paragraph("Note: Due to the active Amnesty extending to December 2027, collections are minimal and result primarily from the waiver of interest, penalties, and fees.")

        if df_isic is not None and not df_isic.empty:
            prefix_map = {"3111": "PIT", "3112": "CIT", "3113": "Withholding", "3131": "Property Tax", "3141 001": "VAT", "3141 004": "HAT", "3141 005": "Travel Tax", "3141 006": "HCSL", "3142 005": "Insurance Premium", "3142 006": "Stamp Duty"}
            sheet_prefix = None
            for prefix, sheet in prefix_map.items():
                if code.startswith(prefix): sheet_prefix = sheet; break
            if sheet_prefix:
                filtered = df_isic[df_isic['Tax_Type'].str.contains(sheet_prefix, case=False, na=False)]
                filtered = filtered[filtered['Month'] == snap_month]
                if not filtered.empty:
                    top3 = filtered.groupby('ISIC_Sector')['Current'].sum().nlargest(3)
                    doc.add_heading('Top Performing Sectors', level=3)
                    for sector, amount in top3.items(): doc.add_paragraph(f"{sector}: {fmt_currency(amount)}", style='List Bullet')
                    doc.add_paragraph("These sectors accounted for the largest share of sector-attributed collections for the reporting period. Further observation of these possibilities will have to be made.")

    doc.add_heading('Conclusion', level=1)
    verdict = "positive at the aggregate level" if total_variance >= 0 else "negative at the aggregate level"
    doc.add_paragraph(f"The Inland Revenue Department's performance for {snap_month} 2026 produced a net variance of {fmt_currency(total_variance)} ({fmt_pct(pct_total)}), which is {verdict}.")
    if not ytd_df.empty:
        top = ytd_df.nlargest(1, 'Variance').iloc[0]; bottom = ytd_df.nsmallest(1, 'Variance').iloc[0]
        doc.add_paragraph(f"The largest positive variance was recorded in {top['Account']} ({fmt_currency(top['Variance'])}). The largest negative variance was recorded in {bottom['Account']} ({fmt_currency(bottom['Variance'])}).")
    doc.add_paragraph("These variances reflect the current collection status against the calendarized projections. The Department will continue to monitor these streams to assess whether the observed variances represent timing differences or structural shifts in compliance.")

    bio = _io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# ==============================================================================
# DATA LOADING
# ==============================================================================
DATA_DIR = Path("data")
main_files_on_disk, isic_files_on_disk, calib_file_on_disk = [], [], None
if DATA_DIR.exists():
    for p in DATA_DIR.glob("*.xlsx"):
        name_lower = p.name.lower()
        if "isic" in name_lower: isic_files_on_disk.append(p)
        elif "projection" in name_lower: calib_file_on_disk = p
        elif "revenue report" in name_lower or "monthly revenue" in name_lower or "revenue collection" in name_lower: main_files_on_disk.append(p)

def path_to_payload(paths): return [(p.name, p.read_bytes()) for p in paths]
main_payload = path_to_payload(main_files_on_disk)
isic_payload = path_to_payload(isic_files_on_disk)
calib_payload = (calib_file_on_disk.name, calib_file_on_disk.read_bytes()) if calib_file_on_disk else None

# ==============================================================================
# SIDEBAR
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    if main_payload: st.success(f"✓ Auto-loaded {len(main_payload)} monthly report(s) from `data/`")
    else: st.warning("No monthly reports found in `data/` folder. Upload below.")
    if isic_payload: st.success(f"✓ Auto-loaded {len(isic_payload)} ISIC file(s)")
    if calib_payload: st.success(f"✓ Auto-loaded projections: `{calib_payload[0]}`")

    st.markdown("---")
    with st.expander("🔄 Override with manual upload"):
        up_main = st.file_uploader("Monthly Revenue Reports", type=["xlsx"], accept_multiple_files=True, key="up_main")
        up_isic = st.file_uploader("ISIC Variance Reports", type=["xlsx"], accept_multiple_files=True, key="up_isic")
        up_calib = st.file_uploader("Calendarized Projections", type=["xlsx"], key="up_calib")

    if up_main: main_payload = [(f.name, f.getvalue()) for f in up_main]
    if up_isic: isic_payload = [(f.name, f.getvalue()) for f in up_isic]
    if up_calib: calib_payload = (up_calib.name, up_calib.getvalue())

st.title("🇱🇨 Revenue Intelligence Dashboard")

if not main_payload:
    st.error("No revenue reports found. Please add .xlsx files to the `data/` folder in GitHub, or upload manually.")
    st.stop()

# ==============================================================================
# PARSE DATA
# ==============================================================================
df_monthly_proj = parse_collection_and_projection(main_payload)
df_snapshot = parse_summary_snapshot(main_payload)
df_daily, df_offices = parse_daily(main_payload)
df_isic = parse_isic(isic_payload) if isic_payload else pd.DataFrame()
df_calib, _ = parse_calendarized(calib_payload) if calib_payload else (pd.DataFrame(), {})

if df_snapshot.empty:
    st.error("No snapshot rows parsed. Check your Excel files.")
    st.stop()

avail_months = sorted(df_snapshot['Month'].unique(), key=month_order)
avail_quarters = sorted(list(set([get_quarter(m) for m in avail_months])))
avail_codes = sorted(df_snapshot['Code'].unique())
tax_opts = [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in avail_codes]
display_to_code = {f"{c} — {TAX_HEADS.get(c, 'Unknown')}": c for c in avail_codes}

# ==============================================================================
# SIDEBAR FILTERS
# ==============================================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🎯 Analysis Mode")
    analysis_mode = st.radio("Snapshot driver", ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"], label_visibility="collapsed")
    st.markdown("---")
    st.markdown("### 🎛️ Filters")
    st.markdown("**Month**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="m_all", use_container_width=True): st.session_state["month_f"] = avail_months
    if c2.button("None", key="m_none", use_container_width=True): st.session_state["month_f"] = []
    selected_months = st.multiselect("Months", avail_months, default=avail_months, key="month_f", label_visibility="collapsed")

    st.markdown("**Quarter**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="q_all", use_container_width=True): st.session_state["qtr_f"] = avail_quarters
    if c2.button("None", key="q_none", use_container_width=True): st.session_state["qtr_f"] = []
    selected_quarters = st.multiselect("Quarters", avail_quarters, default=avail_quarters, key="qtr_f", label_visibility="collapsed")

    st.markdown("**Tax Head**")
    c1, c2 = st.columns(2)
    if c1.button("All", key="t_all", use_container_width=True): st.session_state["tax_f"] = tax_opts
    if c2.button("None", key="t_none", use_container_width=True): st.session_state["tax_f"] = []
    selected_taxes = st.multiselect("Tax Heads", tax_opts, default=tax_opts, key="tax_f", label_visibility="collapsed")

    st.markdown("**Cash Basis**")
    cash_basis = st.radio("Focus", ["All", "Current Only", "Arrears Only"], horizontal=True, label_visibility="collapsed")

if not selected_months or not selected_quarters or not selected_taxes:
    st.warning("⚠️ Please ensure all filters have at least one selection."); st.stop()

selected_codes = [display_to_code[d] for d in selected_taxes if d in display_to_code]

# ==============================================================================
# FILTERING LOGIC
# ==============================================================================
snap_filtered = df_snapshot[df_snapshot['Month'].isin(selected_months) & df_snapshot['Code'].isin(selected_codes)].copy()
if snap_filtered.empty: st.error("Filters produced 0 snapshot rows."); st.stop()

if analysis_mode == "Quarter-End Snapshot":
    qe_avail = [m for m in snap_filtered['Month'].unique() if m in QUARTER_END_MONTHS]
    if not qe_avail: st.error("No quarter-end snapshot available."); st.stop()
    snap_month = sorted(qe_avail, key=month_order)[-1]
else:
    snap_month = sorted(snap_filtered['Month'].unique(), key=month_order)[-1]

snapshot = snap_filtered[snap_filtered['Month'] == snap_month].copy()
ytd = snapshot[['Code', 'Account', 'Actual_YTD']].copy()

if not df_calib.empty:
    months_through = FISCAL_MONTHS[:month_order(snap_month) + 1]
    proj_win = df_calib[df_calib['Month'].isin(months_through) & df_calib['Code'].isin(selected_codes)]
    ytd_proj = proj_win.groupby('Code').agg(YTD_Projection=('Projection', 'sum')).reset_index()
    proj_source = f"Calendarized (Apr–{snap_month})"
else:
    # Fallback if no calib file: use the monthly projection * months
    ytd_proj = df_monthly_proj[df_monthly_proj['Month'].isin(months_through) & df_monthly_proj['Code'].isin(selected_codes)]
    ytd_proj = ytd_proj.groupby('Code')['Monthly_Projection'].sum().reset_index().rename(columns={'Monthly_Projection': 'YTD_Projection'})
    proj_source = "Monthly Projection Sum"

ytd = ytd.merge(ytd_proj, on='Code', how='left').fillna(0)
ytd['Variance'] = ytd['Actual_YTD'] - ytd['YTD_Projection']
total_actual, total_proj, total_var = ytd['Actual_YTD'].sum(), ytd['YTD_Projection'].sum(), ytd['Variance'].sum()

# ==============================================================================
# TABS
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Executive Dashboard", "📅 Daily & Branch", "🏭 Sector Deep-Dive", "📝 Reports"])

with tab1:
    st.markdown(f"### Monthly Revenue Performance Dashboard — {snap_month} 2026")
    
    if not df_monthly_proj.empty:
        df_current_month = df_monthly_proj[df_monthly_proj['Month'] == snap_month].copy()
        df_current_month = df_current_month[df_current_month['Code'].isin(selected_codes)]
        
        m_actual = df_current_month['Monthly_Actual'].sum()
        m_proj = df_current_month['Monthly_Projection'].sum()
        m_var = m_actual - m_proj
        m_pct = (m_var / m_proj * 100) if m_proj > 0 else 0
        
        st.markdown(f"""
        <div style="display: flex; gap: 20px; margin-bottom: 20px;">
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">Month Actual</div>
                <div style="color: #00B4D8; font-size: 32px; font-weight: bold;">${m_actual/1e6:,.2f}M</div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">Month Projection</div>
                <div style="color: #1A3B47; font-size: 32px; font-weight: bold;">${m_proj/1e6:,.2f}M</div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">Monthly Variance</div>
                <div style="color: {'#2E8B57' if m_var >= 0 else '#DC3545'}; font-size: 32px; font-weight: bold;">
                    {'+' if m_var >= 0 else ''}{m_var/1e6:,.2f}M ({m_pct:.1f}%)
                </div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">YTD Convergence</div>
                <div style="color: #1A3B47; font-size: 32px; font-weight: bold;">{(total_actual/total_proj*100 if total_proj>0 else 0):.1f}%</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📋 Analyst Notes")
        if not df_current_month.empty:
            top_over = df_current_month.nlargest(1, 'Monthly_Variance').iloc[0]
            top_under = df_current_month.nsmallest(1, 'Monthly_Variance').iloc[0]
            top_volume = df_current_month.nlargest(1, 'Monthly_Actual').iloc[0]
            
            st.info(f"""
            **Monthly Overview:** 
            In {snap_month}, actual collections totaled **${m_actual:,.2f}** against a projection of **${m_proj:,.2f}**, resulting in a net variance of **${m_var:,.2f}** ({m_pct:.1f}%). 
            
            The largest revenue driver for the month was **{top_volume['Account']}** (${top_volume['Monthly_Actual']:,.2f}), representing {top_volume['Monthly_Actual']/m_actual*100:.1f}% of total monthly intake. 
            
            **Variance Analysis:** 
            The most significant overperformance was in **{top_over['Account']}**, which exceeded its monthly projection by **${top_over['Monthly_Variance']:,.2f}**. Conversely, the largest shortfall occurred in **{top_under['Account']}**, which missed its projection by **${abs(top_under['Monthly_Variance']):,.2f}**. 
            
            **YTD Context:** 
            Year-to-Date collections currently stand at **${total_actual:,.2f}** against a pro-rata target of **${total_proj:,.2f}**, maintaining a YTD convergence rate of **{(total_actual/total_proj*100 if total_proj>0 else 0):.1f}%**.
            """)
        
        r1, r2 = st.columns(2)
        with r1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("#### Monthly Variance by Tax Head")
            plot_df = df_current_month.sort_values('Monthly_Variance', ascending=False)
            colors = ['#2E8B57' if v >= 0 else '#DC3545' for v in plot_df['Monthly_Variance']]
            fig = go.Figure(go.Bar(x=plot_df['Account'], y=plot_df['Monthly_Variance'], marker_color=colors, text=plot_df['Monthly_Variance'].apply(lambda x: f"${x/1e3:,.1f}K"), textposition='auto'))
            fig.update_layout(xaxis_title=None, yaxis_title="Variance (XCD)", margin=dict(t=20, b=20, l=20, r=20))
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with r2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("#### Target vs Actual (Monthly)")
            fig = go.Figure()
            fig.add_trace(go.Bar(x=plot_df['Account'], y=plot_df['Monthly_Projection'], name='Target', marker_color='#1A3B47'))
            fig.add_trace(go.Bar(x=plot_df['Account'], y=plot_df['Monthly_Actual'], name='Actual', marker_color='#00B4D8'))
            fig.update_layout(barmode='group', xaxis_title=None, yaxis_title="Revenue (XCD)", margin=dict(t=20, b=20, l=20, r=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("Collection & Projection data could not be parsed.")

with tab2:
    st.markdown("### Daily Collections & Branch Breakdown")
    if not df_daily.empty:
        d = df_daily[df_daily['Month'].isin(selected_months)]
        if not d.empty:
            daily_by_day = d.groupby('Day')['Total'].sum().reset_index()
            fig = go.Figure(go.Scatter(x=daily_by_day['Day'], y=daily_by_day['Total'], mode='lines+markers', line=dict(color='#00B4D8', width=3), marker=dict(size=8), fill='tozeroy', fillcolor='rgba(0,180,216,0.15)'))
            st.plotly_chart(fig, use_container_width=True)
    if not df_offices.empty:
        o = df_offices[df_offices['Month'].isin(selected_months)]
        if not o.empty:
            branch_sum = o.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
            st.plotly_chart(px.bar(branch_sum, x='Office', y='Total', color='Total', color_continuous_scale='Blues'), use_container_width=True)

with tab3:
    if not df_isic.empty:
        di = df_isic[df_isic['Month'].isin(selected_months)].copy()
        if cash_basis == "Current Only" or cash_basis == "All": di['Amount'] = di['Current']
        elif cash_basis == "Arrears Only": di['Amount'] = di['Arrears']
        di = di[di['Amount'] > 0]
        if not di.empty:
            sector_totals = di.groupby('ISIC_Sector')['Amount'].sum().reset_index().sort_values('Amount', ascending=False)
            st.plotly_chart(px.bar(sector_totals.head(15), x='Amount', y='ISIC_Sector', orientation='h', color='Amount', color_continuous_scale='Blues'), use_container_width=True)

with tab4:
    st.markdown(f"### Variance Notes — {snap_month} 2026")
    if st.button(f"Generate Report"):
        docx_bytes = build_variance_notes_docx(snap_month, total_actual, total_proj, total_var, ytd, df_isic)
        st.download_button("📥 Download Word Document", docx_bytes, f"Variance_Notes_{snap_month}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
