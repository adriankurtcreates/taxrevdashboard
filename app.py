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
@st.cache_data
def parse_monthly_slices(file_payload):
    all_tax = []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            rev = next((s for s in xls.sheet_names if "actual monthly revenue" in s.lower()), None)
            if not rev: continue
            month = find_month_in_text(filename) or find_month_in_text(rev)
            if not month: continue
            df = pd.read_excel(file_obj, sheet_name=rev, header=None)
            for i in range(5, min(35, len(df))):
                code = normalize_code(df.iloc[i, 0])
                name = df.iloc[i, 1]
                if not code or pd.isna(name): continue
                if not re.match(r'^\d{4}\s\d{3}$', code): continue
                monthly = sum(safe_number(df.iloc[i, c]) for c in range(4, 15))
                all_tax.append({"Month": month, "Quarter": get_quarter(month),
                                "Code": code, "Account": str(name).strip(),
                                "Monthly_Actual": monthly})
        except Exception:
            pass
    return pd.DataFrame(all_tax)

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
            if not month:
                diag.append({"file": filename, "error": "month not found"}); continue
            summary = next((s for s in xls.sheet_names if "revenue summary" in s.lower()), None)
            if not summary:
                summary = next((s for s in xls.sheet_names if "summary" in s.lower()), None)
            if not summary:
                diag.append({"file": filename, "error": "no summary sheet"}); continue
            df = pd.read_excel(file_obj, sheet_name=summary, header=None)
            header_row = 4
            for r in range(15):
                row_str = " ".join(str(v).upper() for v in df.iloc[r].values if pd.notna(v))
                if "TO DATE" in row_str:
                    header_row = r; break
            rows = 0
            for i in range(header_row + 1, len(df)):
                raw = df.iloc[i, 0]; desc = df.iloc[i, 1]
                if pd.isna(raw) or pd.isna(desc): continue
                code = normalize_code(raw)
                if not re.match(r'^\d{4}\s\d{3}$', code): continue
                approved = safe_number(df.iloc[i, 2])
                ytd = safe_number(df.iloc[i, 4]) if 4 < len(df.columns) else 0
                var = safe_number(df.iloc[i, 13]) if 13 < len(df.columns) else 0
                all_snap.append({"Month": month, "Quarter": get_quarter(month),
                                 "Code": code, "Account": str(desc).strip(),
                                 "Approved_Annual": approved, "Actual_YTD": ytd,
                                 "Snapshot_Variance": var})
                rows += 1
            diag.append({"file": filename, "month": month, "rows": rows})
        except Exception as e:
            diag.append({"file": filename, "error": str(e)})
    return pd.DataFrame(all_snap), pd.DataFrame(diag)
@st.cache_data
def parse_calendarized(file_payload):
    if not file_payload:
        return pd.DataFrame(), {"error": "no file"}
    filename, file_bytes = file_payload
    diag = {"file": filename, "rows_parsed": 0, "error": None}
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
                all_proj.append({"Code": code, "Account": str(desc).strip(),
                                 "Month": month, "Projection": val})
        result = pd.DataFrame(all_proj)
        diag["rows_parsed"] = len(result)
        return result, diag
    except Exception as e:
        diag["error"] = str(e)
        return pd.DataFrame(), diag

@st.cache_data
def parse_daily(file_payload):
    all_daily, castries_daily, offices = [], [], []
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
                        castries_daily.append({"Month": month, "Total": ot})
                        offices.append({"Month": month, "Office": "Castries", "Total": ot})
                    elif any(x in s for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        offices.append({"Month": month, "Office": s, "Total": ot})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(castries_daily), pd.DataFrame(offices)

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
                hr, ic, cc, ac = None, None, None, None
                for r in range(15):
                    vals = [str(v).lower() for v in df.iloc[r].values if pd.notna(v)]
                    if any("isic description" in v for v in vals):
                        hr = r
                        for c, val in enumerate(df.iloc[r].values):
                            vl = str(val).lower()
                            if "isic description" in vl: ic = c
                            if "current" in vl and cc is None: cc = c
                            if "arrears" in vl and ac is None: ac = c
                        break
                if hr is None or ic is None: continue
                for i in range(hr + 1, len(df)):
                    sector = df.iloc[i, ic]
                    if pd.isna(sector) or "grand total" in str(sector).lower(): break
                    cur = safe_number(df.iloc[i, cc]) if cc is not None else 0
                    arr = safe_number(df.iloc[i, ac]) if ac is not None else 0
                    if cur > 0 or arr > 0:
                        all_isic.append({"Month": month, "Tax_Type": sheet,
                                         "ISIC_Sector": str(sector).strip(),
                                         "Current": cur, "Arrears": arr,
                                         "Total": cur})
        except Exception:
            pass
    return pd.DataFrame(all_isic)
def generate_tax_paragraph(actual, projection, variance):
    if projection > 0:
        pct = (variance / projection) * 100
    else:
        pct = 0.0
    var_str = fmt_currency(variance)
    pct_str = fmt_pct(pct)
    if variance == 0:
        return (f"Actual collections of {fmt_currency(actual)} matched the projected figure "
                f"of {fmt_currency(projection)}. There is no variance to report.")
    direction = "exceeded" if variance > 0 else "fell short of"
    size = "an overage" if variance > 0 else "a shortfall"
    sentence = (f"Actual collections of {fmt_currency(actual)} {direction} the projected figure "
                f"of {fmt_currency(projection)}, representing {size} of {var_str} or {pct_str}. ")
    if projection > 0:
        rate = (actual / projection) * 100
        sentence += f"This corresponds to a collection rate of {rate:.2f}% of the projection for the period."
    return sentence

def build_variance_notes_docx(snap_month, total_actual, total_projection, total_variance, ytd_df, df_isic):
    doc = Document()
    h = doc.add_heading('SAINT LUCIA', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('DEPARTMENT OF FINANCE', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('INLAND REVENUE DEPARTMENT', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph('Revenue Collection Volume 14: Monthly Performance Report')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(f'Explanation of Variances for the Month of {snap_month} 2026')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(14)

    doc.add_heading('Overview', level=1)
    direction = "an overage" if total_variance >= 0 else "a shortfall"
    pct_total = (total_variance / total_projection) * 100 if total_projection > 0 else 0
    overview_text = (f"For the month of {snap_month} 2026, the Inland Revenue Department projected "
                     f"total revenue of {fmt_currency(total_projection)}. Actual collections amounted "
                     f"to {fmt_currency(total_actual)}, resulting in {direction} of "
                     f"{fmt_currency(abs(total_variance))} or {fmt_pct(abs(pct_total))} "
                     f"{'above' if total_variance >= 0 else 'below'} projection.")
    doc.add_paragraph(overview_text)
    doc.add_paragraph(get_fiscal_context(snap_month))
    doc.add_paragraph(
        f"The aggregate performance for {snap_month} shows "
        f"{'positive' if total_variance >= 0 else 'negative'} net variance. The detailed distribution "
        f"of performance across individual tax heads is presented in the Variance Analysis section below."
    )

    doc.add_heading('Variance Analysis by Tax Type', level=1)
    doc.add_paragraph("The performances by individual tax type are outlined below:")

    for _, row in ytd_df.sort_values('Code').iterrows():
        code = row['Code']; account = row['Account']
        actual = row['Actual_YTD']; projection = row['YTD_Projection']; variance = row['Variance']
        pct = (variance / projection * 100) if projection > 0 else 0.0

        doc.add_heading(f"{code} - {account}", level=2)
        line = (f"Projection: {fmt_currency(projection)} | Actual: {fmt_currency(actual)} | "
                f"Variance: {fmt_currency(variance)} ({fmt_pct(pct)})")
        p = doc.add_paragraph(line); p.runs[0].bold = True
        doc.add_paragraph(generate_tax_paragraph(actual, projection, variance))

        if code == "3131 001":
            doc.add_paragraph("Note: Due to the ongoing moratorium on residential property tax, current "
                              "payments apply exclusively to commercial properties. Arrears collections "
                              "apply to both.")
        elif code == "3141 004":
            doc.add_paragraph("Note: The Hotel Accommodation Tax remains an active registry line for the "
                              "collection and settlement of legacy arrears, accounting for ongoing revenue "
                              "despite no active monthly projections.")
        elif code == "3141 006":
            doc.add_paragraph("Note: The Health and Citizen Security Levy (HCSL) was implemented in March "
                              "2023; no historical data exists prior to that date.")
        elif code == "3142 006":
            doc.add_paragraph("Structural Insight: The top-performing sectors listed for Stamp Duty may not "
                              "be a true indicator of which economic sectors are generating the underlying "
                              "transactions. This is because the ISIC code used to log the payment is "
                              "routinely assigned to the law firm or legal practitioner paying the tax on "
                              "behalf of their clients.")
        elif code == "3422 008":
            doc.add_paragraph("Note: Due to the active Amnesty extending to December 2027, collections "
                              "are minimal and result primarily from the waiver of interest, penalties, "
                              "and fees.")

        if df_isic is not None and not df_isic.empty:
            prefix_map = {"3111": "PIT", "3112": "CIT", "3113": "Withholding", "3131": "Property Tax",
                          "3141 001": "VAT", "3141 004": "HAT", "3141 005": "Travel Tax",
                          "3141 006": "HCSL", "3142 005": "Insurance Premium", "3142 006": "Stamp Duty"}
            sheet_prefix = None
            for prefix, sheet in prefix_map.items():
                if code.startswith(prefix):
                    sheet_prefix = sheet; break
            if sheet_prefix:
                filtered = df_isic[df_isic['Tax_Type'].str.contains(sheet_prefix, case=False, na=False)]
                filtered = filtered[filtered['Month'] == snap_month]
                if not filtered.empty:
                    top3 = filtered.groupby('ISIC_Sector')['Total'].sum().nlargest(3)
                    doc.add_heading('Top Performing Sectors', level=3)
                    for sector, amount in top3.items():
                        doc.add_paragraph(f"{sector}: {fmt_currency(amount)}", style='List Bullet')
                    doc.add_paragraph(
                        "These sectors accounted for the largest share of sector-attributed "
                        "collections for the reporting period. Further observation of these "
                        "possibilities will have to be made."
                    )

    doc.add_heading('Conclusion', level=1)
    verdict = "positive at the aggregate level" if total_variance >= 0 else "negative at the aggregate level"
    doc.add_paragraph(f"The Inland Revenue Department's performance for {snap_month} 2026 produced a "
                      f"net variance of {fmt_currency(total_variance)} ({fmt_pct(pct_total)}), "
                      f"which is {verdict}.")
    if not ytd_df.empty:
        top = ytd_df.nlargest(1, 'Variance').iloc[0]
        bottom = ytd_df.nsmallest(1, 'Variance').iloc[0]
        doc.add_paragraph(f"The largest positive variance was recorded in {top['Account']} "
                          f"({fmt_currency(top['Variance'])}). The largest negative variance was "
                          f"recorded in {bottom['Account']} ({fmt_currency(bottom['Variance'])}).")
    doc.add_paragraph(
        "These variances reflect the current collection status against the calendarized projections. "
        "The Department will continue to monitor these streams to assess whether the observed variances "
        "represent timing differences or structural shifts in compliance."
    )

    bio = _io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# ===========================================================================
# DATA LOADING: Read Excel files from the data/ folder committed to GitHub.
# Falls back to manual upload if data/ is empty or missing.
# ===========================================================================
import os
from pathlib import Path

DATA_DIR = Path("data")

def discover_files(folder, patterns):
    """Return list of file paths in folder matching any pattern (case-insensitive)."""
    if not folder.exists():
        return []
    found = []
    for p in folder.iterdir():
        if not p.is_file():
            continue
        name_lower = p.name.lower()
        if any(pat in name_lower for pat in patterns):
            found.append(p)
    return found

# Discover files based on filename patterns
main_files_on_disk = []
for p in DATA_DIR.glob("*.xlsx") if DATA_DIR.exists() else []:
    name_lower = p.name.lower()
    # Skip files that are clearly ISIC or projections
    if "isic" in name_lower or "projection" in name_lower:
        continue
    # Match any file containing "revenue report", "monthly revenue", or "revenue collection"
    if ("revenue report" in name_lower or 
        "monthly revenue" in name_lower or 
        "revenue collection" in name_lower):
        main_files_on_disk.append(p)

isic_files_on_disk = [p for p in (DATA_DIR.glob("*.xlsx") if DATA_DIR.exists() else [])
                      if "isic" in p.name.lower()]

calib_file_on_disk = None
for p in (DATA_DIR.glob("*.xlsx") if DATA_DIR.exists() else []):
    if "projection" in p.name.lower():
        calib_file_on_disk = p
        break

# Build payload tuples from disk files
def path_to_payload(paths):
    return [(p.name, p.read_bytes()) for p in paths]

main_payload = path_to_payload(main_files_on_disk)
isic_payload = path_to_payload(isic_files_on_disk)
calib_payload = (calib_file_on_disk.name, calib_file_on_disk.read_bytes()) if calib_file_on_disk else None

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")

    # Show what was auto-loaded
    if main_payload:
        st.success(f"✓ Auto-loaded {len(main_payload)} monthly report(s) from `data/`")
    else:
        st.warning("No monthly reports found in `data/` folder. Upload below.")
    if isic_payload:
        st.success(f"✓ Auto-loaded {len(isic_payload)} ISIC file(s)")
    if calib_payload:
        st.success(f"✓ Auto-loaded projections: `{calib_payload[0]}`")

    st.markdown("---")
    with st.expander("🔄 Override with manual upload"):
        st.caption("If you need to test a file without committing it, upload here.")
        up_main = st.file_uploader("Monthly Revenue Reports", type=["xlsx"], accept_multiple_files=True, key="up_main")
        up_isic = st.file_uploader("ISIC Variance Reports", type=["xlsx"], accept_multiple_files=True, key="up_isic")
        up_calib = st.file_uploader("Calendarized Projections", type=["xlsx"], key="up_calib")

    # Manual uploads take priority if present
    if up_main:
        main_payload = [(f.name, f.getvalue()) for f in up_main]
        st.info(f"Using {len(main_payload)} manually uploaded report(s)")
    if up_isic:
        isic_payload = [(f.name, f.getvalue()) for f in up_isic]
    if up_calib:
        calib_payload = (up_calib.name, up_calib.getvalue())

st.title("🇱🇨 Revenue Intelligence Dashboard")

if not main_payload:
    st.error("No revenue reports found. Please add .xlsx files to the `data/` folder in GitHub, or upload manually using the sidebar expander.")
    st.info("📁 **How to add files:** In your GitHub repo, navigate to the `data/` folder, click 'Add file' → 'Upload files', drag your Excel files, and commit.")
    st.stop()

# --- Parse all data ---
df_monthly = parse_monthly_slices(main_payload)
df_snapshot, diag_snap = parse_summary_snapshot(main_payload)
df_daily, df_castries, df_offices = parse_daily(main_payload)

df_isic = pd.DataFrame()
if isic_payload:
    df_isic = parse_isic(isic_payload)

df_calib = pd.DataFrame()
calib_diag = {}
if calib_payload:
    df_calib, calib_diag = parse_calendarized(calib_payload)

with st.expander("🔍 Data Diagnostics", expanded=df_snapshot.empty):
    st.write(f"**Snapshot rows:** {len(df_snapshot)} | **Monthly slices:** {len(df_monthly)} | **ISIC rows:** {len(df_isic)}")
    if not df_snapshot.empty:
        st.write(f"**Σ Actual_YTD from parser:** ${df_snapshot['Actual_YTD'].sum():,.2f}")
        st.write(f"**Months:** {sorted(df_snapshot['Month'].unique())}")
    st.dataframe(diag_snap, use_container_width=True)
    if calib_payload:
        if calib_diag.get("error"):
            st.error(f"Projections error: {calib_diag['error']}")
        else:
            st.success(f"✓ Projections loaded: {calib_diag['rows_parsed']} rows")

if df_snapshot.empty:
    st.error("No snapshot rows parsed. Check the diagnostics above to see what went wrong.")
    st.stop()

avail_months = sorted(df_snapshot['Month'].unique(), key=month_order)
avail_quarters = sorted(df_snapshot['Quarter'].unique())
avail_codes = sorted(df_snapshot['Code'].unique())
tax_opts = [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in avail_codes]
display_to_code = {f"{c} — {TAX_HEADS.get(c, 'Unknown')}": c for c in avail_codes}

with st.sidebar:
    st.markdown("---")
    st.markdown("### 🎯 Analysis Mode")
    analysis_mode = st.radio("Snapshot driver",
                             ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"],
                             label_visibility="collapsed")
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
    st.caption("ℹ️ 'All' and 'Current Only' both match the Revenue Summary.")

if not selected_months:
    st.warning("⚠️ No months selected."); st.stop()
if not selected_quarters:
    st.warning("⚠️ No quarters selected."); st.stop()
if not selected_taxes:
    st.warning("⚠️ No tax heads selected."); st.stop()

selected_codes = [display_to_code[d] for d in selected_taxes if d in display_to_code]

snap_filtered = df_snapshot[
    df_snapshot['Month'].isin(selected_months) &
    df_snapshot['Quarter'].isin(selected_quarters) &
    df_snapshot['Code'].isin(selected_codes)
].copy()

if snap_filtered.empty:
    st.error("Filters produced 0 snapshot rows."); st.stop()

if analysis_mode == "Quarter-End Snapshot":
    qe_avail = [m for m in snap_filtered['Month'].unique() if m in QUARTER_END_MONTHS]
    if not qe_avail:
        st.error("No quarter-end snapshot available."); st.stop()
    snap_month = sorted(qe_avail, key=month_order)[-1]
else:
    snap_month = sorted(snap_filtered['Month'].unique(), key=month_order)[-1]

snapshot = snap_filtered[snap_filtered['Month'] == snap_month].copy()
ytd = snapshot[['Code', 'Account', 'Actual_YTD', 'Approved_Annual', 'Snapshot_Variance']].copy()

if not df_calib.empty:
    months_through = FISCAL_MONTHS[:month_order(snap_month) + 1]
    proj_win = df_calib[df_calib['Month'].isin(months_through) & df_calib['Code'].isin(selected_codes)]
    ytd_proj = proj_win.groupby('Code').agg(YTD_Projection=('Projection', 'sum')).reset_index()
    proj_source = f"Calendarized (Apr–{snap_month})"
    fallback = False
else:
    ytd_proj = snapshot[['Code', 'Snapshot_Variance']].copy()
    ytd_proj['YTD_Projection'] = snapshot['Actual_YTD'].values - snapshot['Snapshot_Variance'].values
    ytd_proj = ytd_proj[['Code', 'YTD_Projection']]
    proj_source = "Fallback (Summary variance)"
    fallback = True

ytd = ytd.merge(ytd_proj, on='Code', how='left')
ytd['YTD_Projection'] = ytd['YTD_Projection'].fillna(0)
ytd['Variance'] = ytd['Actual_YTD'] - ytd['YTD_Projection']

total_actual = ytd['Actual_YTD'].sum()
total_proj = ytd['YTD_Projection'].sum()
total_var = ytd['Variance'].sum()

top_perf = ytd.nlargest(1, 'Variance').iloc[0]
top_short = ytd.nsmallest(1, 'Variance').iloc[0]

is_qe = snap_month in QUARTER_END_MONTHS
badge = "snapshot-badge" if is_qe else "snapshot-badge-warn"
badge_txt = f"✓ Quarter-End: {snap_month}" if is_qe else f"Latest: {snap_month}"

st.markdown(f"""
<div class="narrative-box">
    <b>📊 Executive Summary</b>
    <span class="{badge}">{badge_txt}</span><br><br>
    <b>Source:</b> {snap_month} 2026 Revenue Summary &nbsp;|&nbsp; <b>Projection:</b> {proj_source}<br>
    <b>YTD Actual:</b> ${total_actual:,.2f} &nbsp;|&nbsp; <b>YTD Projection:</b> ${total_proj:,.2f}<br>
    <b>Net {'Overage' if total_var > 0 else 'Shortage'}:</b> ${abs(total_var):,.2f} ({(total_var/total_proj*100) if total_proj>0 else 0:.2f}%)<br><br>
    <b>Top Driver:</b> {top_perf['Account']} (+${top_perf['Variance']:,.2f})<br>
    <b>Main Drag:</b> {top_short['Account']} (${top_short['Variance']:,.2f})
</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
c1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Actual</div><div class="kpi-value">${total_actual:,.0f}</div><div class="kpi-delta">Source: {snap_month}</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Projection</div><div class="kpi-value">${total_proj:,.0f}</div><div class="kpi-delta">{proj_source}</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_var > 0 else "#DC3545"}">${total_var:,.0f}</div></div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📊 Executive Dashboard", "📅 Daily & Branch", "🏭 Sector Deep-Dive", "📝 Reports"])

# ==============================================================================
# DYNAMIC CALCULATION FOR CHARTS (Based on Analysis Mode)
# ==============================================================================
# Determine the exact time window to analyze based on the sidebar selection
if analysis_mode == "YTD (Latest Summary)":
    # Use all months up to and including the snapshot month
    months_in_scope = FISCAL_MONTHS[:month_order(snap_month) + 1]
    period_label = f"YTD to {snap_month}"
elif analysis_mode == "Quarter-End Snapshot":
    # Use only the months in the quarter of the snapshot month
    if snap_month in ["April", "May", "June"]: months_in_scope = ["April", "May", "June"]
    elif snap_month in ["July", "August", "September"]: months_in_scope = ["July", "August", "September"]
    elif snap_month in ["October", "November", "December"]: months_in_scope = ["October", "November", "December"]
    else: months_in_scope = ["January", "February", "March"]
    period_label = f"{get_quarter(snap_month)}"
else:
    # Month Snapshot: Use only the selected month
    months_in_scope = [snap_month]
    period_label = f"{snap_month} 2026"

# Filter the raw monthly data to just this scope
df_scope = df_monthly[
    df_monthly['Month'].isin(months_in_scope) &
    df_monthly['Code'].isin(selected_codes)
]

# Recalculate Totals for this specific scope
scope_actuals = df_scope.groupby('Code')['Monthly_Actual'].sum().reset_index()
scope_actuals.rename(columns={'Monthly_Actual': 'Scope_Actual'}, inplace=True)

# Get Targets for this scope
proj_scope = df_calib[
    df_calib['Month'].isin(months_in_scope) &
    df_calib['Code'].isin(selected_codes)
]
scope_proj = proj_scope.groupby('Code')['Projection'].sum().reset_index()

# Merge
chart_df = scope_actuals.merge(scope_proj, on='Code', how='outer').fillna(0)
chart_df = chart_df.merge(ytd[['Code', 'Account']], on='Code', how='left')
chart_df['Account'] = chart_df['Account'].fillna('Unknown')
chart_df['Variance'] = chart_df['Scope_Actual'] - chart_df['Projection']

# New totals for the KPI cards
scope_total_actual = chart_df['Scope_Actual'].sum()
scope_total_proj = chart_df['Projection'].sum()
scope_total_var = chart_df['Variance'].sum()
scope_conv_pct = (scope_total_actual / scope_total_proj * 100) if scope_total_proj > 0 else 0


# ==============================================================================
# TAB 1: EXECUTIVE DASHBOARD (Dynamic)
# ==============================================================================
with tab1:
    st.markdown(f"### Performance Dashboard — {period_label}")
    
    # --- ROW 1: KPI CARDS ---
    if not df_scope.empty:
        months_sorted = sorted(df_scope['Month'].unique(), key=month_order)
        latest_month_in_scope = months_sorted[-1] if months_sorted else None
        prev_month_in_scope = months_sorted[-2] if len(months_sorted) > 1 else None
        
        current_rev = df_scope[df_scope['Month'] == latest_month_in_scope]['Monthly_Actual'].sum() if latest_month_in_scope else 0
        prev_rev = df_scope[df_scope['Month'] == prev_month_in_scope]['Monthly_Actual'].sum() if prev_month_in_scope else 0
        
        if prev_rev > 0:
            delta_pct = ((current_rev - prev_rev) / prev_rev) * 100
            delta_color = "#2E8B57" if delta_pct >= 0 else "#DC3545"
            delta_arrow = "↑" if delta_pct >= 0 else "↓"
            delta_text = f"{delta_arrow} {abs(delta_pct):.1f}%"
        else:
            delta_color = "#6C757D"
            delta_text = "N/A"
            
        st.markdown(f"""
        <div style="display: flex; gap: 20px; margin-bottom: 20px;">
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">{period_label} Convergence</div>
                <div style="color: #00B4D8; font-size: 32px; font-weight: bold;">{scope_conv_pct:.1f}%</div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">{period_label} Revenue</div>
                <div style="color: #1A3B47; font-size: 32px; font-weight: bold;">${scope_total_actual/1e6:,.2f}M</div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">{period_label} Target</div>
                <div style="color: #1A3B47; font-size: 32px; font-weight: bold;">${scope_total_proj/1e6:,.2f}M</div>
            </div>
            <div style="flex: 1; padding: 20px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
                <div style="color: #6C757D; font-size: 14px; font-weight: 600; margin-bottom: 5px;">Latest Month</div>
                <div style="color: #1A3B47; font-size: 32px; font-weight: bold;">${current_rev/1e6:,.2f}M 
                <span style="font-size: 16px; color: {delta_color}; margin-left: 10px;">{delta_text}</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # --- ROW 2: THREE CHARTS ---
    r2_col1, r2_col2, r2_col3 = st.columns(3)
    
    with r2_col1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_2a = st.radio("View", ["Donut", "Bar"], horizontal=True, label_visibility="collapsed", key="v2a")
        
        met = len(chart_df[chart_df['Variance'] >= 0])
        missed = len(chart_df[chart_df['Variance'] < 0])
        status_df = pd.DataFrame({'Status': ['Met', 'Missed'], 'Count': [met, missed]})
        
        if view_2a == "Donut":
            fig_2a = px.pie(status_df, names='Status', values='Count', hole=0.6, color='Status', color_discrete_map={'Met': '#00B4D8', 'Missed': '#1A3B47'})
            fig_2a.update_traces(textposition='inside', textinfo='percent+label')
        else:
            fig_2a = px.bar(status_df, x='Status', y='Count', color='Status', color_discrete_map={'Met': '#00B4D8', 'Missed': '#1A3B47'})
            
        fig_2a.update_layout(margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
        st.plotly_chart(fig_2a, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with r2_col2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_2b = st.radio("View", ["Monthly Revenue", "Monthly Variance"], horizontal=True, label_visibility="collapsed", key="v2b")
        
        if view_2b == "Monthly Revenue":
            rev_by_month = df_scope.groupby('Month')['Monthly_Actual'].sum().reset_index()
            rev_by_month['Order'] = rev_by_month['Month'].map(month_order)
            rev_by_month = rev_by_month.sort_values('Order')
            fig_2b = px.bar(rev_by_month, x='Month', y='Monthly_Actual', color_discrete_sequence=['#00B4D8'])
            fig_2b.update_layout(xaxis_title=None, yaxis_title="Revenue (XCD)", margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
        else:
            # Calculate variance by month
            var_month = chart_df.copy() # if only one month, it's just the total variance
            if len(months_in_scope) > 1:
                var_by_month = df_scope.groupby('Month')['Monthly_Actual'].sum().reset_index()
                proj_by_month = proj_scope.groupby('Month')['Projection'].sum().reset_index()
                var_by_month = var_by_month.merge(proj_by_month, on='Month', how='outer').fillna(0)
                var_by_month['Variance'] = var_by_month['Monthly_Actual'] - var_by_month['Projection']
                var_by_month['Order'] = var_by_month['Month'].map(month_order)
                var_by_month = var_by_month.sort_values('Order')
                colors = ['#DC3545' if v < 0 else '#2E8B57' for v in var_by_month['Variance']]
                fig_2b = go.Figure(go.Bar(x=var_by_month['Month'], y=var_by_month['Variance'], marker_color=colors))
            else:
                fig_2b = go.Figure(go.Bar(x=['Total'], y=[scope_total_var], marker_color=['#DC3545' if scope_total_var < 0 else '#2E8B57']))
            
            fig_2b.update_layout(xaxis_title=None, yaxis_title="Variance (XCD)", margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
        st.plotly_chart(fig_2b, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with r2_col3:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_2c = st.radio("View", ["Target vs Actual", "Convergence Rate"], horizontal=True, label_visibility="collapsed", key="v2c")
        
        chart_df['Status'] = chart_df.apply(lambda r: 'Met' if r['Variance'] >= 0 else 'Missed', axis=1)
        
        if view_2c == "Target vs Actual":
            fig_2c = go.Figure()
            fig_2c.add_trace(go.Bar(x=chart_df['Code'], y=chart_df['Projection'], name='Target', marker_color='#1A3B47'))
            fig_2c.add_trace(go.Bar(x=chart_df['Code'], y=chart_df['Scope_Actual'], name='Actual', marker_color=chart_df['Status'].map({'Met': '#00B4D8', 'Missed': '#F59E0B'})))
            fig_2c.update_layout(barmode='group', xaxis_title=None, yaxis_title="Revenue (XCD)", showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        else:
            chart_df['Conv'] = (chart_df['Scope_Actual'] / chart_df['Projection'] * 100).fillna(0)
            fig_2c = px.bar(chart_df, x='Code', y='Conv', color='Conv', color_continuous_scale='RdYlGn', range_y=[0, 150])
            fig_2c.add_hline(y=100, line_dash="dash", line_color="black", annotation_text="100% Target")
            fig_2c.update_layout(xaxis_title=None, yaxis_title="Convergence %", showlegend=False)
            
        fig_2c.update_layout(margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_2c, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # --- ROW 3: THREE CHARTS ---
    r3_col1, r3_col2, r3_col3 = st.columns(3)

    with r3_col1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_3a = st.radio("View", ["Donut", "Treemap"], horizontal=True, label_visibility="collapsed", key="v3a")
        
        cat_map = {"3111": "Income Tax", "3112": "Income Tax", "3113": "Income Tax", "3131": "Property Tax", "3141": "Goods & Services", "3142": "Excises", "3143": "Permissions", "3151": "Customs", "3422": "Admin Fees"}
        chart_df['Category'] = chart_df['Code'].apply(lambda x: cat_map.get(x[:4], "Other"))
        cat_df = chart_df.groupby('Category')['Scope_Actual'].sum().reset_index()
        
        if view_3a == "Donut":
            fig_3a = px.pie(cat_df, names='Category', values='Scope_Actual', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_3a.update_traces(textposition='inside', textinfo='percent+label')
        else:
            fig_3a = px.treemap(cat_df, path=['Category'], values='Scope_Actual', color='Scope_Actual', color_continuous_scale='Blues')
            
        fig_3a.update_layout(margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
        st.plotly_chart(fig_3a, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with r3_col2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_3b = st.radio("View", ["Current vs Arrears", "Top vs Bottom"], horizontal=True, label_visibility="collapsed", key="v3b")
        
        if view_3b == "Current vs Arrears":
            # Split YTD data into Current vs Arrears based on code suffix
            chart_df['Type'] = chart_df['Code'].apply(lambda x: 'Arrears' if x.endswith('002') else 'Current')
            stack_df = chart_df.groupby(['Account', 'Type'])['Scope_Actual'].sum().reset_index()
            stack_df = stack_df[stack_df['Scope_Actual'] > 0]
            fig_3b = px.bar(stack_df, x='Account', y='Scope_Actual', color='Type', color_discrete_map={'Current': '#00B4D8', 'Arrears': '#1A3B47'}, barmode='relative')
            fig_3b.update_layout(xaxis_title=None, yaxis_title="Revenue (XCD)", showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        else:
            sorted_df = chart_df.sort_values('Scope_Actual', ascending=False)
            top5 = sorted_df.head(5)
            bottom5 = sorted_df.tail(5)
            combined = pd.concat([top5, bottom5])
            colors = ['#00B4D8']*5 + ['#DC3545']*5
            fig_3b = go.Figure(go.Bar(x=combined['Account'], y=combined['Scope_Actual'], marker_color=colors))
            fig_3b.update_layout(xaxis_title=None, yaxis_title="Revenue (XCD)", showlegend=False)
            
        fig_3b.update_layout(margin=dict(t=20, b=0, l=0, r=0))
        st.plotly_chart(fig_3b, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with r3_col3:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        view_3c = st.radio("View", ["Branch Bar", "Branch Pie"], horizontal=True, label_visibility="collapsed", key="v3c")
        
        if not df_offices.empty:
            o = df_offices[df_offices['Month'].isin(months_in_scope)]
            branch_sum = o.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
            
            if view_3c == "Branch Bar":
                fig_3c = px.bar(branch_sum, x='Total', y='Office', orientation='h', color='Office', color_discrete_sequence=px.colors.qualitative.Set3)
            else:
                fig_3c = px.pie(branch_sum, names='Office', values='Total', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
                
            fig_3c.update_layout(xaxis_title="Revenue (XCD)", yaxis_title=None, margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig_3c, use_container_width=True)
        else:
            st.info("Branch data not available.")
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# TAB 2: DAILY & BRANCHES
# ==============================================================================
with tab2:
    st.markdown(f"### Daily Collections — {period_label}")
    
    view_t2 = st.radio("Display Style", ["Daily Trend", "Branch Focus"], horizontal=True, label_visibility="collapsed", key="v_t2")
    
    if view_t2 == "Daily Trend":
        if not df_daily.empty:
            d = df_daily[df_daily['Month'].isin(months_in_scope)]
            if not d.empty:
                daily_by_day = d.groupby('Day')['Total'].sum().reset_index()
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=daily_by_day['Day'], y=daily_by_day['Total'], mode='lines+markers', line=dict(color='#00B4D8', width=3), marker=dict(size=8), fill='tozeroy', fillcolor='rgba(0,180,216,0.15)'))
                fig.update_layout(xaxis_title="Day of Month", yaxis_title="Revenue (XCD)", margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Daily data not available.")
            
    else:
        if not df_offices.empty:
            o = df_offices[df_offices['Month'].isin(months_in_scope)]
            if not o.empty:
                branch_sum = o.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
                fig = px.bar(branch_sum, x='Office', y='Total', color='Total', color_continuous_scale='Blues', text_auto='.2s')
                fig.update_layout(xaxis_title=None, yaxis_title="Revenue (XCD)", margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Branch data not available.")

# ==============================================================================
# TAB 3: SECTOR DEEP-DIVE
# ==============================================================================
with tab3:
    if not df_isic.empty:
        st.markdown(f"### ISIC Sector Deep-Dive — {period_label}")
        st.caption("⚠️ ISIC data reflects payments made via the IRD's online payment platform only.")

        view_t3 = st.radio("View", ["Sector Performance", "Tax Head Drill-Down"], horizontal=True, label_visibility="collapsed", key="v_t3")
        
        di = df_isic[df_isic['Month'].isin(months_in_scope)].copy()
        if cash_basis == "Current Only" or cash_basis == "All":
            di['Amount'] = di['Current']
        elif cash_basis == "Arrears Only":
            di['Amount'] = di['Arrears']
        di = di[di['Amount'] > 0]

        if not di.empty:
            if view_t3 == "Sector Performance":
                sector_totals = di.groupby('ISIC_Sector')['Amount'].sum().reset_index().sort_values('Amount', ascending=False)
                top_sectors = sector_totals.head(15)

                chart_type_t3 = st.radio("Chart Type", ["Bar", "Treemap", "Pie"], horizontal=True, label_visibility="collapsed", key="chart_t3")
                
                if chart_type_t3 == "Bar":
                    fig = px.bar(top_sectors, x='Amount', y='ISIC_Sector', orientation='h', color='Amount', color_continuous_scale='Blues')
                elif chart_type_t3 == "Treemap":
                    fig = px.treemap(top_sectors, path=['ISIC_Sector'], values='Amount', color='Amount', color_continuous_scale='Blues')
                else:
                    fig = px.pie(top_sectors, names='ISIC_Sector', values='Amount', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                
                fig.update_layout(xaxis_title="Revenue (XCD)", yaxis_title=None, margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
                
            else:
                tax_options = di['Tax_Type'].unique()
                drill_tax = st.selectbox("Select a Tax Head to see its top sectors:", tax_options)
                
                drill = di[di['Tax_Type'] == drill_tax]
                tax_sectors = drill.groupby('ISIC_Sector')['Amount'].sum().nlargest(10).reset_index()
                
                fig = px.bar(tax_sectors, x='Amount', y='ISIC_Sector', orientation='h', color='Amount', color_continuous_scale='Blues')
                fig.update_layout(xaxis_title="Revenue (XCD)", yaxis_title=None, margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No ISIC data for the selected filters.")
    else:
        st.info("Upload ISIC Variance Reports to unlock sector data.")

# ==============================================================================
# TAB 4: REPORTS
# ==============================================================================
with tab4:
    st.markdown(f"### Variance Notes — {snap_month} 2026")
    st.caption("Generates the IMF-style Word document using the strict Master Prompt template.")
    if st.button(f"Generate Report for {snap_month}"):
        with st.spinner("Generating Word document..."):
            docx_bytes = build_variance_notes_docx(
                snap_month, total_actual, total_proj, total_var, ytd, df_isic
            )
            st.download_button(
                label="📥 Download Variance Notes (.docx)",
                data=docx_bytes,
                file_name=f"Variance_Notes_{snap_month}_2026.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
