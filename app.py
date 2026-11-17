import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io as _io
from datetime import datetime, date
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
    .stDataFrame { font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONSTANTS & HELPERS
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

QUARTER_DEFS = {
    "Q1 (Apr-Jun)": ["April", "May", "June"],
    "Q2 (Jul-Sep)": ["July", "August", "September"],
    "Q3 (Oct-Dec)": ["October", "November", "December"],
    "Q4 (Jan-Mar)": ["January", "February", "March"],
}
QUARTER_END_MONTHS = {"June": "Q1 (Apr-Jun)", "September": "Q2 (Jul-Sep)",
                     "December": "Q3 (Oct-Dec)", "March": "Q4 (Jan-Mar)"}

SEPARATED_TAX_HEADS = {
    "3111": "Personal Income Tax",
    "3112": "Corporate Income Tax",
    "3113": "Withholding Income Tax",
}

ISIC_TAX_PREFIX_MAP = {
    "3111": "Personal Income Tax",
    "3112": "Corporate Income Tax",
    "3113": "Withholding",
    "3131": "Property Tax",
    "3141": "Value Added Tax",
    "3141 004": "Hotel Accommodation",
    "3141 005": "Travel Tax",
    "3141 006": "Health and Citizen Security",
    "3142 005": "Insurance Premium",
    "3142 006": "Stamp Duty",
}

ISIC_TARGETS = [
    "PIT", "CIT", "Withholding", "VAT", "Hotel Accommodation", "Travel Tax",
    "HCSL", "Health and Citizen Security", "Insurance Premium", "Stamp Duty",
]


def current_fiscal_year_label(month, reference_year=None):
    if month is None:
        return None
    if reference_year is None:
        reference_year = date.today().year
    if month in ("January", "February", "March"):
        return f"{reference_year - 1}/{reference_year}"
    return f"{reference_year}/{reference_year + 1}"


def extract_year_from_text(text):
    if not text:
        return None
    m = re.search(r"(20\d{2})", str(text))
    return int(m.group(1)) if m else None


def resolve_fiscal_label(month, filename):
    year = extract_year_from_text(filename)
    return current_fiscal_year_label(month, reference_year=year)


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
    except Exception:
        return 0.0


def get_quarter(month):
    for q, months in QUARTER_DEFS.items():
        if month in months:
            return q
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
    pos_map = {
        "Q1 (Apr-Jun)": {"April": "first", "May": "second", "June": "third"},
        "Q2 (Jul-Sep)": {"July": "first", "August": "second", "September": "third"},
        "Q3 (Oct-Dec)": {"October": "first", "November": "second", "December": "third"},
        "Q4 (Jan-Mar)": {"January": "first", "February": "second", "March": "third"},
    }
    qnum = q.split(" ")[0] if q != "Unknown" else ""
    pos = pos_map.get(q, {}).get(month, "")
    return f"This month represents the {pos} month of the {qnum} quarter of the fiscal year."


def is_separated_tax_head(code):
    return any(code.startswith(p) for p in SEPARATED_TAX_HEADS)


def parent_tax_prefix(code):
    m = re.match(r"^(\d{4})", str(code))
    return m.group(1) if m else ""


def months_through_for(snap_month):
    return [m for m in FISCAL_MONTHS[:month_order(snap_month) + 1]]


def months_in_quarter(q_label):
    return QUARTER_DEFS.get(q_label, [])


def apply_cash_basis(df, basis):
    """
    Filter an ISIC-like dataframe by current / arrears / all.
    Expects columns: 'Current' and 'Arrears'.
    """
    if df is None or df.empty:
        return df
    if basis == "Current Only":
        out = df.copy()
        out["Amount"] = out["Current"]
        return out[out["Amount"] > 0]
    if basis == "Arrears Only":
        out = df.copy()
        out["Amount"] = out["Arrears"]
        return out[out["Amount"] > 0]
    out = df.copy()
    out["Amount"] = out["Current"] + out["Arrears"]
    return out[out["Amount"] > 0]


# ==============================================================================
# PARSERS (unchanged structurally; carry FiscalYear everywhere)
# ==============================================================================
@st.cache_data
def parse_collection_and_projection(file_payload):
    all_data = []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            fiscal = resolve_fiscal_label(month, filename) if month else None
            sheet = next((s for s in xls.sheet_names if "collection & projection" in s.lower() or "collection and projection" in s.lower()), None)
            if not sheet: continue
            df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            header_row = None
            for r in range(20):
                row_str = " ".join(str(v).lower() for v in df.iloc[r].values if pd.notna(v))
                if "projection" in row_str and "collections" in row_str:
                    header_row = r; break
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
                    "Month": month, "FiscalYear": fiscal, "Code": code,
                    "Account": str(desc).strip(),
                    "Monthly_Projection": projection, "Monthly_Actual": collections,
                    "Monthly_Variance": collections - projection,
                })
        except Exception:
            pass
    return pd.DataFrame(all_data)


@st.cache_data
def parse_summary_snapshot(file_payload):
    all_snap = []
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
            fiscal = resolve_fiscal_label(month, filename)
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
                all_snap.append({
                    "Month": month, "FiscalYear": fiscal, "Code": code,
                    "Account": str(desc).strip(), "Actual_YTD": ytd,
                })
        except Exception:
            pass
    return pd.DataFrame(all_snap)


@st.cache_data
def parse_calendarized(file_payload):
    if not file_payload: return pd.DataFrame()
    filename, file_bytes = file_payload
    try:
        file_obj = _io.BytesIO(file_bytes)
        xls = pd.ExcelFile(file_obj)
        sheet = next((s for s in xls.sheet_names if "projection" in s.lower()), xls.sheet_names[0])
        df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
        MONTH_COLS = {"April": 3, "May": 4, "June": 5, "July": 7, "August": 8, "September": 9,
                      "October": 11, "November": 12, "December": 13, "January": 15, "February": 16, "March": 17}
        calib_fiscal = None
        for r in range(min(6, len(df))):
            for v in df.iloc[r].values:
                if pd.isna(v): continue
                y = extract_year_from_text(v)
                if y:
                    if "fiscal" in str(v).lower():
                        calib_fiscal = str(v).strip()
                    else:
                        calib_fiscal = f"{y}/{y+1}"
                    break
            if calib_fiscal: break
        all_proj = []
        for i in range(7, min(55, len(df))):
            raw = df.iloc[i, 0]; desc = df.iloc[i, 1]
            if pd.isna(raw) or pd.isna(desc): continue
            code = normalize_code(raw)
            if not re.match(r'^\d{4}\s\d{3}$', code): continue
            for month, col in MONTH_COLS.items():
                val = safe_number(df.iloc[i, col]) if col < len(df.columns) else 0
                all_proj.append({
                    "Code": code, "Account": str(desc).strip(),
                    "Month": month, "Projection": val, "FiscalYear": calib_fiscal,
                })
        return pd.DataFrame(all_proj)
    except Exception:
        return pd.DataFrame()


@st.cache_data
def parse_daily(file_payload):
    all_daily, all_offices = [], []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            fiscal = resolve_fiscal_label(month, filename) if month else None
            ds = next((s for s in xls.sheet_names if "daily collection" in s.lower()), None)
            if not ds or not month: continue
            df = pd.read_excel(file_obj, sheet_name=ds, header=None)
            for i in range(5, 35):
                dv = df.iloc[i, 0]
                if pd.isna(dv) or not isinstance(dv, (datetime, pd.Timestamp)): continue
                t = safe_number(df.iloc[i, 20])
                if t > 0:
                    all_daily.append({"Month": month, "FiscalYear": fiscal,
                                      "Day": pd.to_datetime(dv).day, "Total": t})
            for i in range(30, 55):
                on = df.iloc[i, 0]; ot = safe_number(df.iloc[i, 20])
                if pd.notna(on) and isinstance(on, str):
                    s = on.strip()
                    if "Sub Total" in s:
                        all_offices.append({"Month": month, "FiscalYear": fiscal,
                                            "Office": "Castries", "Total": ot})
                    elif any(x in s for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        all_offices.append({"Month": month, "FiscalYear": fiscal,
                                            "Office": s, "Total": ot})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(all_offices)


@st.cache_data
def parse_isic(file_payload):
    all_isic = []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month: continue
            fiscal = resolve_fiscal_label(month, filename)
            for t in ISIC_TARGETS:
                sheet = next((s for s in xls.sheet_names if t.lower() in s.lower()), None)
                if not sheet: continue
                df = pd.read_excel(file_obj, sheet_name=sheet, header=None)
                hr, ic, curc, arrc = None, None, None, None
                for r in range(15):
                    vals = [str(v).lower() for v in df.iloc[r].values if pd.notna(v)]
                    if any("isic description" in v for v in vals):
                        hr = r
                        for c, val in enumerate(df.iloc[r].values):
                            vl = str(val).lower()
                            if "isic description" in vl: ic = c
                            elif "current" in vl and curc is None: curc = c
                            elif "arrear" in vl and arrc is None: arrc = c
                        break
                if hr is None or ic is None: continue
                separated = any(k.lower() in sheet.lower() for k in SEPARATED_TAX_HEADS.values())
                for i in range(hr + 1, len(df)):
                    sector = df.iloc[i, ic]
                    if pd.isna(sector) or "grand total" in str(sector).lower(): break
                    cur = safe_number(df.iloc[i, curc]) if curc is not None else 0
                    arr = safe_number(df.iloc[i, arrc]) if (arrc is not None and separated) else 0
                    if cur > 0 or arr > 0:
                        all_isic.append({
                            "Month": month, "FiscalYear": fiscal,
                            "Tax_Type": sheet, "ISIC_Sector": str(sector).strip(),
                            "Current": cur, "Arrears": arr,
                        })
        except Exception:
            pass
    return pd.DataFrame(all_isic)


# ==============================================================================
# NARRATIVE HELPERS
# ==============================================================================
def generate_tax_paragraph(actual, projection, variance):
    if projection > 0: pct = (variance / projection) * 100
    else: pct = 0.0
    var_str = fmt_currency(variance); pct_str = fmt_pct(pct)
    if variance == 0:
        return (f"Actual collections of {fmt_currency(actual)} matched the projected figure "
                f"of {fmt_currency(projection)}. There is no variance to report.")
    direction = "exceeded" if variance > 0 else "fell short of"
    size = "an overage" if variance > 0 else "a shortfall"
    sentence = (f"Actual collections of {fmt_currency(actual)} {direction} the projected figure "
                f"of {fmt_currency(projection)}, representing {size} of {var_str} or {pct_str}. ")
    if projection > 0:
        sentence += f"This corresponds to a collection rate of {(actual/projection*100):.2f}% of the projection for the period."
    return sentence


def build_comment_section(comments_dict, month, fiscal_year):
    out = []
    for code, text in comments_dict.get((month, fiscal_year), {}).items():
        if text and text.strip():
            out.append((code, text.strip()))
    return out


# ==============================================================================
# REPORT BUILDER (structure preserved, comment key includes year)
# ==============================================================================
def build_variance_notes_docx(snap_month, snap_fiscal, total_actual, total_projection,
                              total_variance, ytd_df, df_isic, comments=None):
    comments = comments or {}
    doc = Document()

    h = doc.add_heading('SAINT LUCIA', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('DEPARTMENT OF FINANCE', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('INLAND REVENUE DEPARTMENT', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph('Revenue Collection Volume 14: Monthly Performance Report'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(f'Explanation of Variances for {snap_month} — Fiscal Year {snap_fiscal}'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(14)

    doc.add_heading('1. Overview', level=1)
    direction = "an overage" if total_variance >= 0 else "a shortfall"
    pct_total = (total_variance / total_projection * 100) if total_projection > 0 else 0.0
    overview_text = (f"For {snap_month} (FY {snap_fiscal}), the Inland Revenue Department projected total "
                     f"revenue of {fmt_currency(total_projection)}. Actual collections amounted to "
                     f"{fmt_currency(total_actual)}, resulting in {direction} of "
                     f"{fmt_currency(abs(total_variance))} or {fmt_pct(abs(pct_total))} "
                     f"{'above' if total_variance >= 0 else 'below'} projection.")
    doc.add_paragraph(overview_text)
    doc.add_paragraph(get_fiscal_context(snap_month))
    doc.add_paragraph("This note examines the aggregate and tax-head-level variance against the calendarised "
                      "projection for the period. Observations are presented on a current-versus-arrears basis "
                      "for the heads where those lines are separately maintained (Personal Income Tax, Corporate "
                      "Income Tax, and Withholding Tax), and on an aggregated basis for all other heads.")

    user_comments = build_comment_section(comments, snap_month, snap_fiscal)
    if user_comments:
        doc.add_heading('2. Confirming Explanations Provided by the Department', level=1)
        doc.add_paragraph("The Department has provided the following confirming explanations in support of the "
                          "observed outturns. These are reproduced verbatim to preserve the Department's account "
                          "of the period.")
        for code, text in user_comments:
            label = f"{code} — {TAX_HEADS.get(code, 'General')}" if re.match(r"^\d{4}\s\d{3}$", code) else code
            p = doc.add_paragraph(); p.add_run(f"{label}: ").bold = True; p.add_run(text)

    doc.add_heading('3. Variance Analysis by Tax Type', level=1)
    doc.add_paragraph("The tables and paragraphs below set out, for each reporting line, the period projection, "
                      "actual outturn, variance, and rate of attainment. Where the Department has supplied a "
                      "confirming explanation it is reproduced beneath the relevant line.")
    for _, row in ytd_df.sort_values('Code').iterrows():
        code = row['Code']; account = row['Account']
        actual = row['Actual_YTD']; projection = row['YTD_Projection']; variance = row['Variance']
        pct = (variance / projection * 100) if projection > 0 else 0.0
        doc.add_heading(f"{code} — {account}", level=2)
        line = (f"Projection: {fmt_currency(projection)} | Actual: {fmt_currency(actual)} | "
                f"Variance: {fmt_currency(variance)} ({fmt_pct(pct)})")
        p = doc.add_paragraph(line); p.runs[0].bold = True
        doc.add_paragraph(generate_tax_paragraph(actual, projection, variance))
        if is_separated_tax_head(code) and df_isic is not None and not df_isic.empty:
            prefix = parent_tax_prefix(code)
            sheet_match = ISIC_TAX_PREFIX_MAP.get(prefix)
            if sheet_match:
                filtered = df_isic[df_isic['Tax_Type'].str.contains(sheet_match, case=False, na=False)]
                filtered = filtered[filtered['Month'] == snap_month]
                if not filtered.empty:
                    cur_amt = filtered['Current'].sum()
                    arr_amt = filtered['Arrears'].sum() if 'Arrears' in filtered.columns else 0.0
                    if cur_amt > 0 or arr_amt > 0:
                        doc.add_heading('Current vs Arrears Split', level=3)
                        doc.add_paragraph(f"Current collections: {fmt_currency(cur_amt)} | "
                                          f"Arrears collections: {fmt_currency(arr_amt)}")
        if df_isic is not None and not df_isic.empty:
            sheet_prefix = None
            for prefix, sheet in ISIC_TAX_PREFIX_MAP.items():
                if code.startswith(prefix):
                    sheet_prefix = sheet; break
            if sheet_prefix:
                filtered = df_isic[df_isic['Tax_Type'].str.contains(sheet_prefix, case=False, na=False)]
                filtered = filtered[filtered['Month'] == snap_month]
                if not filtered.empty:
                    top3 = filtered.groupby('ISIC_Sector')['Current'].sum().nlargest(3)
                    if (top3 > 0).any():
                        doc.add_heading('Top Performing Sectors', level=3)
                        for sector, amount in top3.items():
                            doc.add_paragraph(f"{sector}: {fmt_currency(amount)}", style='List Bullet')
                        doc.add_paragraph("These sectors accounted for the largest share of sector-attributed "
                                          "collections for the reporting period. No inference is drawn as to the "
                                          "underlying drivers unless explicitly stated in the Department's comments above.")

    doc.add_heading('4. Conclusion', level=1)
    verdict = "positive at the aggregate level" if total_variance >= 0 else "negative at the aggregate level"
    doc.add_paragraph(f"The Inland Revenue Department's outturn for {snap_month} (FY {snap_fiscal}) produced a net "
                      f"variance of {fmt_currency(total_variance)} ({fmt_pct(pct_total)}), which is {verdict}.")
    if not ytd_df.empty:
        top = ytd_df.nlargest(1, 'Variance').iloc[0]; bottom = ytd_df.nsmallest(1, 'Variance').iloc[0]
        doc.add_paragraph(f"The largest positive variance was recorded in {top['Account']} "
                          f"({fmt_currency(top['Variance'])}). The largest negative variance was recorded in "
                          f"{bottom['Account']} ({fmt_currency(bottom['Variance'])}).")
    doc.add_paragraph("These variances reflect the current collection status against the calendarised projections. "
                      "The Department will continue to monitor these streams to assess whether the observed variances "
                      "represent timing differences or structural shifts in compliance. Where the Department has "
                      "provided a confirming explanation, it has been reproduced in Section 2 above.")

    bio = _io.BytesIO(); doc.save(bio); return bio.getvalue()


# ==============================================================================
# DATA LOADING
# ==============================================================================
DATA_DIR = Path("data")
main_files_on_disk, isic_files_on_disk, calib_file_on_disk = [], [], None
if DATA_DIR.exists():
    for p in DATA_DIR.glob("*.xlsx"):
        n = p.name.lower()
        if "isic" in n: isic_files_on_disk.append(p)
        elif "projection" in n: calib_file_on_disk = p
        elif "revenue report" in n or "monthly revenue" in n or "revenue collection" in n:
            main_files_on_disk.append(p)


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
    if main_payload: st.success(f"✓ Auto-loaded {len(main_payload)} monthly report(s)")
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
    st.error("No revenue reports found. Add .xlsx files to `data/` or upload manually.")
    st.stop()

# ==============================================================================
# PARSE DATA
# ==============================================================================
df_monthly_proj = parse_collection_and_projection(main_payload)
df_snapshot = parse_summary_snapshot(main_payload)
df_daily, df_offices = parse_daily(main_payload)
df_isic = parse_isic(isic_payload) if isic_payload else pd.DataFrame()
df_calib = parse_calendarized(calib_payload) if calib_payload else pd.DataFrame()

if df_snapshot.empty:
    st.error("No snapshot rows parsed. Check your Excel files.")
    st.stop()

avail_months = sorted(df_snapshot['Month'].unique(), key=month_order)
avail_quarters = [q for q in QUARTER_DEFS if any(m in avail_months for m in QUARTER_DEFS[q])]
avail_codes = sorted(df_snapshot['Code'].unique())
tax_opts = [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in avail_codes]
display_to_code = {f"{c} — {TAX_HEADS.get(c, 'Unknown')}": c for c in avail_codes}

# ==============================================================================
# SIDEBAR FILTERS
# ==============================================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🎯 Analysis Mode")
    analysis_mode = st.radio(
        "Snapshot driver",
        ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"],
        label_visibility="collapsed",
    )
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
    cash_basis = st.radio("Focus", ["All", "Current Only", "Arrears Only"],
                          horizontal=True, label_visibility="collapsed")

    # -------- Variance Notes comment tool (now includes YEAR) --------
    st.markdown("---")
    st.markdown("### 📝 Variance Notes — Departmental Comments")
    st.caption("Pick the exact month AND year. Comments are stored per (month, fiscal year).")

    if "comments_store" not in st.session_state:
        st.session_state["comments_store"] = {}

    # Build a defensible list of (Year, Month) pairs from the snapshot data.
    year_month_pairs = []
    for _, r in df_snapshot[["Month", "FiscalYear"]].drop_duplicates().iterrows():
        if pd.notna(r["FiscalYear"]):
            year = str(r["FiscalYear"]).split("/")[0]  # label year
            # For Jan–Mar, the label year is year-1; the "month year" is +1.
            if r["Month"] in ("January", "February", "March"):
                month_year = int(year) + 1
            else:
                month_year = int(year)
            year_month_pairs.append((month_year, r["Month"], r["FiscalYear"]))
    year_month_pairs = sorted(set(year_month_pairs))

    if not year_month_pairs:
        st.info("No months available yet.")
    else:
        month_years = sorted({y for (y, m, fy) in year_month_pairs}, reverse=True)
        cm_year = st.selectbox("Year", month_years, key="cm_year")
        avail_months_for_year = sorted(
            {m for (y, m, fy) in year_month_pairs if y == cm_year}, key=month_order
        )
        cm_month = st.selectbox("Month", avail_months_for_year, key="cm_month")
        comment_fy = current_fiscal_year_label(cm_month, reference_year=cm_year)
        st.markdown(f"**Fiscal Year:** `{comment_fy}`  ·  **Month:** `{cm_month} {cm_year}`")

        key = (cm_month, comment_fy)
        existing = st.session_state["comments_store"].get(key, {})

        cm_tax_label = st.selectbox(
            "Tax head (or General)",
            ["General"] + [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in avail_codes],
            key="cm_tax",
        )
        if cm_tax_label == "General":
            label_for_store = "General"
            prefill = existing.get("General", "")
        else:
            label_for_store = display_to_code[cm_tax_label]
            prefill = existing.get(label_for_store, "")

        comment_text = st.text_area("Comment", value=prefill,
                                    key=f"cm_text_{comment_fy}_{label_for_store}", height=100)

        col_c1, col_c2 = st.columns(2)
        if col_c1.button("Save comment", use_container_width=True):
            store = st.session_state["comments_store"].setdefault(key, {})
            if comment_text.strip():
                store[label_for_store] = comment_text.strip()
            else:
                store.pop(label_for_store, None)
            st.success("Saved.")
        if col_c2.button("Clear comment", use_container_width=True):
            store = st.session_state["comments_store"].setdefault(key, {})
            store.pop(label_for_store, None)
            st.info("Cleared.")

        if existing:
            st.markdown(f"**Saved for {cm_month} {cm_year} (FY {comment_fy}):**")
            for k2, v2 in existing.items():
                label = f"{k2} — {TAX_HEADS.get(k2, 'General')}" if re.match(r"^\d{4}\s\d{3}$", k2) else k2
                st.markdown(f"- **{label}**: {v2[:80]}{'…' if len(v2) > 80 else ''}")

if not selected_months or not selected_quarters or not selected_taxes:
    st.warning("⚠️ Please ensure all filters have at least one selection.")
    st.stop()

selected_codes = [display_to_code[d] for d in selected_taxes if d in display_to_code]

# ==============================================================================
# FILTERING LOGIC
# ==============================================================================
snap_filtered = df_snapshot[
    df_snapshot['Month'].isin(selected_months) & df_snapshot['Code'].isin(selected_codes)
].copy()
if snap_filtered.empty:
    st.error("Filters produced 0 snapshot rows.")
    st.stop()

if analysis_mode == "Quarter-End Snapshot":
    qe_avail = [m for m in snap_filtered['Month'].unique() if m in QUARTER_END_MONTHS]
    snap_month = sorted(qe_avail, key=month_order)[-1] if qe_avail else sorted(snap_filtered['Month'].unique(), key=month_order)[-1]
else:
    snap_month = sorted(snap_filtered['Month'].unique(), key=month_order)[-1]

fys = snap_filtered[snap_filtered['Month'] == snap_month]['FiscalYear'].dropna().unique()
snap_fiscal = fys[0] if len(fys) else current_fiscal_year_label(snap_month)

snapshot = snap_filtered[snap_filtered['Month'] == snap_month].copy()
ytd = snapshot[['Code', 'Account', 'Actual_YTD']].copy()

months_through = months_through_for(snap_month)

if not df_calib.empty:
    if "FiscalYear" in df_calib.columns and df_calib["FiscalYear"].notna().any():
        calib_fy = df_calib["FiscalYear"].dropna().unique()[0]
    else:
        calib_fy = snap_fiscal
    proj_win = df_calib[df_calib['Month'].isin(months_through) & df_calib['Code'].isin(selected_codes)]
    ytd_proj = proj_win.groupby('Code').agg(YTD_Projection=('Projection', 'sum')).reset_index()
    proj_source = f"Calendarised (Apr–{snap_month}, FY {calib_fy})"
else:
    ytd_proj = df_monthly_proj[
        df_monthly_proj['Month'].isin(months_through) & df_monthly_proj['Code'].isin(selected_codes)
    ]
    ytd_proj = (ytd_proj.groupby('Code')['Monthly_Projection'].sum()
                       .reset_index()
                       .rename(columns={'Monthly_Projection': 'YTD_Projection'}))
    proj_source = "Monthly Projection Sum (fallback)"

ytd = ytd.merge(ytd_proj, on='Code', how='left').fillna(0)
ytd['Variance'] = ytd['Actual_YTD'] - ytd['YTD_Projection']
total_actual = ytd['Actual_YTD'].sum()
total_proj = ytd['YTD_Projection'].sum()
total_var = ytd['Variance'].sum()

# ==============================================================================
# SHARED ANALYSIS HELPERS
# ==============================================================================
def build_kpi_row(actual, projection, variance, label_prefix="YTD"):
    pct = (variance / projection * 100) if projection > 0 else 0.0
    ach = (actual / projection * 100) if projection > 0 else 0.0
    var_color = "#2E8B57" if variance >= 0 else "#DC3545"
    st.markdown(f"""
    <div style="display: flex; gap: 15px; margin-bottom: 20px;">
        <div style="flex: 1; padding: 18px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
            <div style="color: #6C757D; font-size: 13px; font-weight: 600;">{label_prefix} Actual</div>
            <div style="color: #00B4D8; font-size: 26px; font-weight: bold;">${actual/1e6:,.2f}M</div>
        </div>
        <div style="flex: 1; padding: 18px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
            <div style="color: #6C757D; font-size: 13px; font-weight: 600;">{label_prefix} Projection</div>
            <div style="color: #1A3B47; font-size: 26px; font-weight: bold;">${projection/1e6:,.2f}M</div>
        </div>
        <div style="flex: 1; padding: 18px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
            <div style="color: #6C757D; font-size: 13px; font-weight: 600;">{label_prefix} Variance</div>
            <div style="color: {var_color}; font-size: 26px; font-weight: bold;">
                {'+' if variance >= 0 else ''}${variance/1e6:,.2f}M ({pct:+.1f}%)
            </div>
        </div>
        <div style="flex: 1; padding: 18px; background-color: #f8f9fa; border-radius: 10px; border: 1px solid #e9ecef;">
            <div style="color: #6C757D; font-size: 13px; font-weight: 600;">Achievement</div>
            <div style="color: #1A3B47; font-size: 26px; font-weight: bold;">{ach:.1f}%</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def achievement_gauge(actual, projection, title):
    ach = (actual / projection * 100) if projection > 0 else 0.0
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=min(ach, 150),
        number={"suffix": "%", "font": {"size": 28}},
        title={"text": title, "font": {"size": 14}},
        gauge={
            "axis": {"range": [0, 150], "tickwidth": 1},
            "bar": {"color": "#00B4D8" if ach >= 100 else ("#F59E0B" if ach >= 85 else "#DC3545")},
            "steps": [
                {"range": [0, 85], "color": "#FFE5E5"},
                {"range": [85, 100], "color": "#FFF4D6"},
                {"range": [100, 150], "color": "#E1F5E8"},
            ],
            "threshold": {"line": {"color": "#1A3B47", "width": 3}, "thickness": 0.8, "value": 100},
        },
    ))
    fig.update_layout(height=240, margin=dict(t=40, b=10, l=10, r=10))
    return fig


def tax_head_bar_chart(df, value_col, label_col='Account', title="", top_n=None, horizontal=True):
    if df.empty:
        return None
    d = df.copy()
    if top_n:
        d = d.nlargest(top_n, value_col)
    d = d.sort_values(value_col, ascending=horizontal)
    if horizontal:
        fig = go.Figure(go.Bar(
            x=d[value_col], y=d[label_col], orientation='h',
            marker_color=['#2E8B57' if v >= 0 else '#DC3545' for v in d[value_col]],
            text=d[value_col].apply(lambda x: f"${x/1e6:,.2f}M"),
            textposition='outside',
        ))
        fig.update_layout(title=title, height=max(300, len(d) * 32),
                          margin=dict(t=40, b=20, l=20, r=20),
                          xaxis_title="XCD", yaxis_title=None)
    else:
        fig = go.Figure(go.Bar(
            x=d[label_col], y=d[value_col],
            marker_color=['#2E8B57' if v >= 0 else '#DC3545' for v in d[value_col]],
            text=d[value_col].apply(lambda x: f"${x/1e6:,.2f}M"),
            textposition='auto',
        ))
        fig.update_layout(title=title, height=380, margin=dict(t=40, b=20, l=20, r=20))
    return fig


def target_vs_actual_bar(df, title="Target vs Actual"):
    if df.empty:
        return None
    d = df.sort_values('YTD_Projection', ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=d['YTD_Projection'], y=d['Account'], orientation='h',
                         name='Projection', marker_color='#1A3B47'))
    fig.add_trace(go.Bar(x=d['Actual_YTD'], y=d['Account'], orientation='h',
                         name='Actual', marker_color='#00B4D8'))
    fig.update_layout(barmode='group', title=title, height=max(320, len(d) * 35),
                      margin=dict(t=40, b=20, l=20, r=20),
                      xaxis_title="XCD", yaxis_title=None,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    return fig


def tax_mix_donut(df, title="Tax Mix"):
    if df.empty:
        return None
    fig = px.pie(df, values='Actual_YTD', names='Account', hole=0.5, title=title,
                 color_discrete_sequence=px.colors.sequential.Blues_r)
    fig.update_traces(textposition='inside', textinfo='percent+label',
                      textfont=dict(size=10))
    fig.update_layout(height=400, margin=dict(t=40, b=10, l=10, r=10), showlegend=False)
    return fig


def ytd_trend_chart(df_monthly, code_filter, fiscal_year):
    """
    Build a YTD cumulative trend of Actual and Projection across fiscal months.
    """
    d = df_monthly[(df_monthly['Code'].isin(code_filter)) & (df_monthly['FiscalYear'] == fiscal_year)].copy()
    if d.empty:
        return None
    d = d.groupby('Month').agg(
        Monthly_Actual=('Monthly_Actual', 'sum'),
        Monthly_Projection=('Monthly_Projection', 'sum'),
    ).reset_index()
    d['Month_Order'] = d['Month'].apply(month_order)
    d = d.sort_values('Month_Order')
    d['Cum_Actual'] = d['Monthly_Actual'].cumsum()
    d['Cum_Projection'] = d['Monthly_Projection'].cumsum()

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=d['Month'], y=d['Cum_Projection'], name='Cumulative Projection',
                             mode='lines+markers', line=dict(color='#1A3B47', width=2, dash='dash')))
    fig.add_trace(go.Scatter(x=d['Month'], y=d['Cum_Actual'], name='Cumulative Actual',
                             mode='lines+markers', line=dict(color='#00B4D8', width=3),
                             fill='tonexty', fillcolor='rgba(0,180,216,0.1)'))
    fig.update_layout(title="YTD Cumulative Actual vs Projection",
                      height=340, margin=dict(t=40, b=20, l=20, r=20),
                      xaxis_title=None, yaxis_title="XCD",
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    return fig


def top_sectors_table(df_isic, tax_type_contains, months, basis="All", top_n=10):
    """
    Filter ISIC rows for a tax type across a set of months and return the top
    sectors table.  basis='All'/'Current Only'/'Arrears Only'.
    """
    if df_isic is None or df_isic.empty:
        return None
    d = df_isic[df_isic['Tax_Type'].str.contains(tax_type_contains, case=False, na=False)].copy()
    d = d[d['Month'].isin(months)]
    d = apply_cash_basis(d, basis)
    if d.empty:
        return None
    g = d.groupby('ISIC_Sector')['Amount'].sum().reset_index()
    g = g.sort_values('Amount', ascending=False).head(top_n)
    g['Amount'] = g['Amount'].apply(lambda x: f"${x:,.2f}")
    g = g.rename(columns={'ISIC_Sector': 'Sector', 'Amount': 'Collected'})
    return g


def year_end_outlook(snap_month, ytd_df, df_calib, selected_codes):
    """
    Return a dataframe describing (per tax head) YTD actual, YTD projection,
    remaining full-year projection, and projected FY outturn assuming current
    run-rate holds.
    """
    if df_calib.empty:
        return None
    mth_through = months_through_for(snap_month)
    full_year = FISCAL_MONTHS
    remaining = [m for m in full_year if m not in mth_through]

    full_proj = df_calib[df_calib['Code'].isin(selected_codes) & df_calib['Month'].isin(full_year)]
    full_proj = full_proj.groupby('Code').agg(FullYear_Projection=('Projection', 'sum')).reset_index()

    remaining_proj = df_calib[df_calib['Code'].isin(selected_codes) & df_calib['Month'].isin(remaining)]
    remaining_proj = remaining_proj.groupby('Code').agg(Remaining_Projection=('Projection', 'sum')).reset_index()

    out = ytd_df.merge(full_proj, on='Code', how='left').fillna(0)
    out = out.merge(remaining_proj, on='Code', how='left').fillna(0)
    out['Months_Elapsed'] = len(mth_through)
    out['Months_Remaining'] = len(remaining)
    out['RunRate'] = out.apply(
        lambda r: (r['Actual_YTD'] / r['Months_Elapsed']) if r['Months_Elapsed'] > 0 else 0, axis=1
    )
    out['Projected_FY_Outturn'] = out['Actual_YTD'] + out['RunRate'] * out['Months_Remaining']
    out['Projected_FY_Variance'] = out['Projected_FY_Outturn'] - out['FullYear_Projection']
    return out


# ==============================================================================
# TABS
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Executive Dashboard", "📅 Daily & Branch",
    "🏭 Sector Deep-Dive", "📝 Reports"
])

with tab1:
    # ==========================================================================
    # ANALYSIS MODE: YTD (Latest Summary)  →  comprehensive YTD dashboard
    # ==========================================================================
    if analysis_mode == "YTD (Latest Summary)":
        st.markdown(f"### YTD Performance Dashboard — through {snap_month} (FY {snap_fiscal})")

        # --- KPI row ---
        build_kpi_row(total_actual, total_proj, total_var, label_prefix="YTD")

        # --- Row 1: trend + achievement gauge ---
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = ytd_trend_chart(df_monthly_proj, selected_codes, snap_fiscal)
            if fig: st.plotly_chart(fig, use_container_width=True)
            else: st.info("No monthly trend data available for this fiscal year.")
        with c2:
            st.plotly_chart(achievement_gauge(total_actual, total_proj, "YTD Achievement"),
                            use_container_width=True)

        # --- Row 2: tax mix donut + target vs actual ---
        c1, c2 = st.columns(2)
        with c1:
            fig = tax_mix_donut(ytd[ytd['Actual_YTD'] > 0], "YTD Actual Revenue Mix")
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = target_vs_actual_bar(ytd, "YTD Target vs Actual by Tax Head")
            if fig: st.plotly_chart(fig, use_container_width=True)

        # --- Row 3: variance by tax head (horizontal bar) ---
        fig = tax_head_bar_chart(ytd, 'Variance', title="YTD Variance by Tax Head (XCD)")
        if fig: st.plotly_chart(fig, use_container_width=True)

        # --- Row 4: Analyst Notes as a table + bullets ---
        st.markdown("#### 📋 Analyst Notes")
        top_over = ytd.nlargest(1, 'Variance').iloc[0]
        top_under = ytd.nsmallest(1, 'Variance').iloc[0]
        top_volume = ytd.nlargest(1, 'Actual_YTD').iloc[0]
        achievement = (total_actual / total_proj * 100) if total_proj > 0 else 0
        notes_table = pd.DataFrame([
            {"Metric": "Reporting Period", "Value": f"{snap_month} (FY {snap_fiscal})"},
            {"Metric": "Months Elapsed", "Value": f"{len(months_through)} of 12"},
            {"Metric": "YTD Actual", "Value": fmt_currency(total_actual)},
            {"Metric": "YTD Projection", "Value": fmt_currency(total_proj)},
            {"Metric": "YTD Variance", "Value": fmt_currency(total_var)},
            {"Metric": "Achievement Rate", "Value": f"{achievement:.1f}%"},
            {"Metric": "Top Volume Head", "Value": f"{top_volume['Account']} ({fmt_currency(top_volume['Actual_YTD'])})"},
            {"Metric": "Largest Positive Variance", "Value": f"{top_over['Account']} ({fmt_currency(top_over['Variance'])})"},
            {"Metric": "Largest Negative Variance", "Value": f"{top_under['Account']} ({fmt_currency(top_under['Variance'])})"},
        ])
        st.table(notes_table)

        st.markdown("**Key Observations:**")
        obs = [
            f"YTD collections of **{fmt_currency(total_actual)}** represent **{achievement:.1f}%** of the "
            f"pro-rata target of **{fmt_currency(total_proj)}** after **{len(months_through)}** months.",
            f"**{top_volume['Account']}** is the largest contributor at **{fmt_currency(top_volume['Actual_YTD'])}** "
            f"({top_volume['Actual_YTD']/total_actual*100:.1f}% of YTD intake)." if total_actual > 0 else "",
            f"Best-performing head vs target: **{top_over['Account']}** at **{fmt_currency(top_over['Variance'])}** "
            f"above projection.",
            f"Weakest head vs target: **{top_under['Account']}** at **{fmt_currency(abs(top_under['Variance']))}** "
            f"below projection.",
        ]
        for o in obs:
            if o: st.markdown(f"- {o}")

        # --- Row 5: YTD per tax head table ---
        st.markdown("#### 📊 YTD Performance by Tax Head")
        tbl = ytd.copy()
        tbl['Achievement %'] = tbl.apply(
            lambda r: (r['Actual_YTD'] / r['YTD_Projection'] * 100) if r['YTD_Projection'] > 0 else 0, axis=1
        )
        display_tbl = tbl[['Code', 'Account', 'Actual_YTD', 'YTD_Projection', 'Variance', 'Achievement %']].copy()
        display_tbl = display_tbl.rename(columns={
            'Actual_YTD': 'YTD Actual', 'YTD_Projection': 'YTD Projection',
            'Variance': 'YTD Variance',
        })
        display_tbl['YTD Actual'] = display_tbl['YTD Actual'].apply(fmt_currency)
        display_tbl['YTD Projection'] = display_tbl['YTD Projection'].apply(fmt_currency)
        display_tbl['YTD Variance'] = display_tbl['YTD Variance'].apply(fmt_currency)
        display_tbl['Achievement %'] = display_tbl['Achievement %'].apply(lambda x: f"{x:.1f}%")
        st.dataframe(display_tbl, use_container_width=True, hide_index=True)

        # --- Row 6: Top-10 ISIC sectors per tax head (YTD) ---
        if not df_isic.empty:
            st.markdown("#### 🏭 Top-10 ISIC Sectors per Tax Head (YTD)")
            isic_ytd = df_isic[
                (df_isic['Month'].isin(months_through)) & (df_isic['FiscalYear'] == snap_fiscal)
            ].copy() if 'FiscalYear' in df_isic.columns else df_isic[df_isic['Month'].isin(months_through)].copy()

            # Tax heads actually in scope (respecting selected filters)
            in_scope_prefixes = sorted({parent_tax_prefix(c) for c in selected_codes})
            prefix_to_sheet = {p: ISIC_TAX_PREFIX_MAP.get(p) for p in in_scope_prefixes}
            prefix_to_sheet = {k: v for k, v in prefix_to_sheet.items() if v}

            if not prefix_to_sheet:
                st.info("No ISIC-mapped tax heads in the current filter.")
            else:
                tabs_isic = st.tabs(list(prefix_to_sheet.values()))
                for i, (prefix, sheet) in enumerate(prefix_to_sheet.items()):
                    with tabs_isic[i]:
                        t = top_sectors_table(isic_ytd, sheet, months_through, basis=cash_basis, top_n=10)
                        if t is None or t.empty:
                            st.info(f"No ISIC data for {sheet} in the current period.")
                        else:
                            col_a, col_b = st.columns([1, 1])
                            with col_a:
                                st.dataframe(t, use_container_width=True, hide_index=True)
                            with col_b:
                                chart_df = t.copy()
                                chart_df['Collected'] = chart_df['Collected'].str.replace(r'[$,]', '', regex=True).astype(float)
                                fig = px.bar(chart_df, x='Collected', y='Sector', orientation='h',
                                             title=f"Top Sectors — {sheet}",
                                             color='Collected', color_continuous_scale='Blues')
                                fig.update_layout(height=380, yaxis={'categoryorder': 'total ascending'},
                                                  margin=dict(t=40, b=10, l=10, r=10))
                                st.plotly_chart(fig, use_container_width=True)

        # --- Row 7: Year-end outlook ---
        if not df_calib.empty:
            st.markdown("#### 📈 Fiscal Year-End Outlook (Run-Rate Extrapolation)")
            outlook = year_end_outlook(snap_month, ytd, df_calib, selected_codes)
            if outlook is not None and not outlook.empty:
                out_disp = outlook[[
                    'Code', 'Account', 'Actual_YTD', 'FullYear_Projection',
                    'Projected_FY_Outturn', 'Projected_FY_Variance'
                ]].copy()
                for c in ['Actual_YTD', 'FullYear_Projection', 'Projected_FY_Outturn', 'Projected_FY_Variance']:
                    out_disp[c] = out_disp[c].apply(fmt_currency)
                out_disp = out_disp.rename(columns={
                    'Actual_YTD': 'YTD Actual', 'FullYear_Projection': 'Full-Year Projection',
                    'Projected_FY_Outturn': 'Projected FY Outturn (Run-Rate)',
                    'Projected_FY_Variance': 'Projected Variance',
                })
                st.dataframe(out_disp, use_container_width=True, hide_index=True)
                st.caption(
                    "Projected FY outturn assumes each tax head continues to collect at its current "
                    "average monthly rate for the remainder of the fiscal year. It is indicative only "
                    "and does not incorporate policy or seasonality assumptions."
                )
        else:
            st.info("Calendarised projections not loaded — fiscal year-end outlook unavailable.")

    # ==========================================================================
    # ANALYSIS MODE: Quarter-End Snapshot  →  per-quarter dashboard tabs
    # ==========================================================================
    elif analysis_mode == "Quarter-End Snapshot":
        st.markdown(f"### Quarterly Performance — FY {snap_fiscal}")
        st.caption("Each tab shows that quarter's actuals, projections, achievement, and top sectors. "
                   "Incomplete quarters show what has been collected and what remains against projection.")

        quarter_tabs = st.tabs([q for q in QUARTER_DEFS if q in selected_quarters])
        active_quarters = [q for q in QUARTER_DEFS if q in selected_quarters]

        for i, q_label in enumerate(active_quarters):
            with quarter_tabs[i]:
                q_months = months_in_quarter(q_label)

                # Actuals from snapshots (month-end summaries within the quarter)
                q_snap = df_snapshot[
                    df_snapshot['Month'].isin(q_months) &
                    df_snapshot['Code'].isin(selected_codes) &
                    (df_snapshot['FiscalYear'] == snap_fiscal if 'FiscalYear' in df_snapshot.columns else True)
                ]
                months_with_data = sorted(q_snap['Month'].unique(), key=month_order)

                # YTD snapshot for the last month of the quarter that has data
                if months_with_data:
                    q_end_month = months_with_data[-1]
                    q_actual = q_snap[q_snap['Month'] == q_end_month][['Code', 'Account', 'Actual_YTD']].copy()
                else:
                    q_end_month = None
                    q_actual = df_snapshot[df_snapshot['Month'].isin(q_months)][['Code', 'Account']].drop_duplicates().copy()
                    q_actual['Actual_YTD'] = 0

                # Quarter-only collections from monthly collection files
                q_monthly_actuals = df_monthly_proj[
                    df_monthly_proj['Month'].isin(q_months) & df_monthly_proj['Code'].isin(selected_codes)
                ]
                if 'FiscalYear' in q_monthly_actuals.columns:
                    q_monthly_actuals = q_monthly_actuals[q_monthly_actuals['FiscalYear'] == snap_fiscal]
                quarter_actual_sum = q_monthly_actuals['Monthly_Actual'].sum()
                quarter_proj_sum = q_monthly_actuals['Monthly_Projection'].sum()

                # Full quarter projection (from calendarised) - sum of all 3 months
                if not df_calib.empty:
                    q_proj = df_calib[
                        df_calib['Month'].isin(q_months) & df_calib['Code'].isin(selected_codes)
                    ]
                    quarter_proj_sum = q_proj['Projection'].sum()
                    quarter_proj_by_code = q_proj.groupby('Code').agg(
                        Quarter_Projection=('Projection', 'sum')).reset_index()
                else:
                    quarter_proj_by_code = (
                        q_monthly_actuals.groupby('Code')['Monthly_Projection'].sum()
                        .reset_index().rename(columns={'Monthly_Projection': 'Quarter_Projection'})
                    )

                status = "Complete" if len(months_with_data) == 3 else (
                    "In Progress" if len(months_with_data) > 0 else "Not Started"
                )
                remaining = quarter_proj_sum - quarter_actual_sum
                st.markdown(f"**Quarter Status:** `{status}`  ·  "
                            f"**Months with data:** {', '.join(months_with_data) if months_with_data else 'none'}  ·  "
                            f"**Quarter Reference Month:** {q_end_month if q_end_month else '—'}")

                if status == "Not Started":
                    st.info(f"No collections recorded for {q_label} yet. "
                            f"Full-quarter projection: {fmt_currency(quarter_proj_sum)}.")
                    if not quarter_proj_by_code.empty:
                        pd_tbl = quarter_proj_by_code.copy()
                        pd_tbl = pd_tbl.merge(
                            df_snapshot[['Code', 'Account']].drop_duplicates(),
                            on='Code', how='left'
                        )
                        pd_tbl['Quarter_Projection'] = pd_tbl['Quarter_Projection'].apply(fmt_currency)
                        st.dataframe(pd_tbl.rename(columns={
                            'Quarter_Projection': 'Quarter Projection'}), use_container_width=True, hide_index=True)
                    continue

                # KPI row for the quarter
                build_kpi_row(quarter_actual_sum, quarter_proj_sum,
                              quarter_actual_sum - quarter_proj_sum, label_prefix="Quarter")

                c1, c2 = st.columns([1, 1])
                with c1:
                    st.plotly_chart(
                        achievement_gauge(quarter_actual_sum, quarter_proj_sum, f"{q_label} Achievement"),
                        use_container_width=True
                    )
                with c2:
                    if remaining > 0:
                        st.markdown(f"""
                        <div style="padding: 20px; background-color: #FFF4D6; border-radius: 10px;
                                    border-left: 5px solid #F59E0B; height: 200px; display: flex;
                                    flex-direction: column; justify-content: center;">
                            <div style="font-size: 13px; color: #6C757D; font-weight: 600;">Remaining to Collect</div>
                            <div style="font-size: 32px; color: #F59E0B; font-weight: bold; margin: 8px 0;">
                                {fmt_currency(remaining)}
                            </div>
                            <div style="font-size: 13px; color: #6C757D;">
                                Across {3 - len(months_with_data)} remaining month(s) in {q_label}.
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style="padding: 20px; background-color: #E1F5E8; border-radius: 10px;
                                    border-left: 5px solid #2E8B57; height: 200px; display: flex;
                                    flex-direction: column; justify-content: center;">
                            <div style="font-size: 13px; color: #6C757D; font-weight: 600;">Quarter Fully Collected</div>
                            <div style="font-size: 32px; color: #2E8B57; font-weight: bold; margin: 8px 0;">
                                +{fmt_currency(abs(remaining))}
                            </div>
                            <div style="font-size: 13px; color: #6C757D;">Over the quarter projection.</div>
                        </div>
                        """, unsafe_allow_html=True)

                # Per-head quarter table
                st.markdown(f"#### {q_label} — Performance by Tax Head")
                q_by_head = quarter_proj_by_code.merge(
                    q_monthly_actuals.groupby('Code').agg(Quarter_Actual=('Monthly_Actual', 'sum')).reset_index(),
                    on='Code', how='outer'
                ).fillna(0)
                accts = df_snapshot[['Code', 'Account']].drop_duplicates()
                q_by_head = q_by_head.merge(accts, on='Code', how='left')
                q_by_head['Variance'] = q_by_head['Quarter_Actual'] - q_by_head['Quarter_Projection']
                q_by_head['Achievement %'] = q_by_head.apply(
                    lambda r: (r['Quarter_Actual'] / r['Quarter_Projection'] * 100) if r['Quarter_Projection'] > 0 else 0,
                    axis=1
                )
                disp = q_by_head[[
                    'Code', 'Account', 'Quarter_Actual', 'Quarter_Projection', 'Variance', 'Achievement %'
                ]].copy()
                disp['Quarter_Actual'] = disp['Quarter_Actual'].apply(fmt_currency)
                disp['Quarter_Projection'] = disp['Quarter_Projection'].apply(fmt_currency)
                disp['Variance'] = disp['Variance'].apply(fmt_currency)
                disp['Achievement %'] = disp['Achievement %'].apply(lambda x: f"{x:.1f}%")
                disp = disp.rename(columns={
                    'Quarter_Actual': 'Quarter Actual',
                    'Quarter_Projection': 'Quarter Projection',
                })
                st.dataframe(disp, use_container_width=True, hide_index=True)

                # Top sectors for the quarter
                if not df_isic.empty:
                    st.markdown(f"#### {q_label} — Top Sectors")
                    q_isic = df_isic[df_isic['Month'].isin(q_months)].copy()
                    if 'FiscalYear' in q_isic.columns:
                        q_isic = q_isic[q_isic['FiscalYear'] == snap_fiscal]
                    q_isic = apply_cash_basis(q_isic, cash_basis)
                    if q_isic.empty:
                        st.info("No ISIC data for this quarter.")
                    else:
                        top_all = (q_isic.groupby('ISIC_Sector')['Amount'].sum()
                                   .reset_index().sort_values('Amount', ascending=False).head(10))
                        top_all['Amount'] = top_all['Amount'].apply(fmt_currency)
                        top_all = top_all.rename(columns={'ISIC_Sector': 'Sector', 'Amount': 'Collected'})
                        st.dataframe(top_all, use_container_width=True, hide_index=True)

    # ==========================================================================
    # ANALYSIS MODE: Month Snapshot  →  keep the original monthly dashboard
    # ==========================================================================
    else:
        st.markdown(f"### Monthly Revenue Performance Dashboard — {snap_month} (FY {snap_fiscal})")

        if not df_monthly_proj.empty:
            df_current_month = df_monthly_proj[
                (df_monthly_proj['Month'] == snap_month) & (df_monthly_proj['Code'].isin(selected_codes))
            ].copy()

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

            # --- Analyst Notes: clean table + bullets ---
            st.markdown("#### 📋 Analyst Notes")
            if not df_current_month.empty:
                top_over = df_current_month.nlargest(1, 'Monthly_Variance').iloc[0]
                top_under = df_current_month.nsmallest(1, 'Monthly_Variance').iloc[0]
                top_volume = df_current_month.nlargest(1, 'Monthly_Actual').iloc[0]

                month_tbl = pd.DataFrame([
                    {"Metric": "Month", "Value": f"{snap_month} (FY {snap_fiscal})"},
                    {"Metric": "Month Actual", "Value": fmt_currency(m_actual)},
                    {"Metric": "Month Projection", "Value": fmt_currency(m_proj)},
                    {"Metric": "Month Variance", "Value": fmt_currency(m_var)},
                    {"Metric": "Month Variance %", "Value": f"{m_pct:.2f}%"},
                    {"Metric": "Top Volume Head", "Value": f"{top_volume['Account']} ({fmt_currency(top_volume['Monthly_Actual'])})"},
                    {"Metric": "Largest Positive Variance", "Value": f"{top_over['Account']} ({fmt_currency(top_over['Monthly_Variance'])})"},
                    {"Metric": "Largest Negative Variance", "Value": f"{top_under['Account']} ({fmt_currency(abs(top_under['Monthly_Variance']))})"},
                    {"Metric": "YTD Convergence", "Value": f"{(total_actual/total_proj*100 if total_proj>0 else 0):.1f}%"},
                ])
                st.table(month_tbl)
                st.markdown("**Key Observations:**")
                st.markdown(f"- The largest revenue driver for the month was **{top_volume['Account']}** "
                            f"at **{fmt_currency(top_volume['Monthly_Actual'])}** "
                            f"({top_volume['Monthly_Actual']/m_actual*100:.1f}% of monthly intake).")
                st.markdown(f"- Overperformance: **{top_over['Account']}** exceeded projection by "
                            f"**{fmt_currency(top_over['Monthly_Variance'])}**.")
                st.markdown(f"- Underperformance: **{top_under['Account']}** fell short of projection by "
                            f"**{fmt_currency(abs(top_under['Monthly_Variance']))}**.")

            # --- Monthly charts ---
            r1, r2 = st.columns(2)
            with r1:
                st.markdown("#### Monthly Variance by Tax Head")
                plot_df = df_current_month.sort_values('Monthly_Variance', ascending=False)
                colors = ['#2E8B57' if v >= 0 else '#DC3545' for v in plot_df['Monthly_Variance']]
                fig = go.Figure(go.Bar(
                    x=plot_df['Account'], y=plot_df['Monthly_Variance'],
                    marker_color=colors,
                    text=plot_df['Monthly_Variance'].apply(lambda x: f"${x/1e3:,.1f}K"),
                    textposition='auto',
                ))
                fig.update_layout(xaxis_title=None, yaxis_title="Variance (XCD)",
                                  margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
            with r2:
                st.markdown("#### Target vs Actual (Monthly)")
                fig = go.Figure()
                fig.add_trace(go.Bar(x=plot_df['Account'], y=plot_df['Monthly_Projection'],
                                     name='Target', marker_color='#1A3B47'))
                fig.add_trace(go.Bar(x=plot_df['Account'], y=plot_df['Monthly_Actual'],
                                     name='Actual', marker_color='#00B4D8'))
                fig.update_layout(
                    barmode='group', xaxis_title=None, yaxis_title="Revenue (XCD)",
                    margin=dict(t=20, b=20, l=20, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Collection & Projection data could not be parsed.")

# ==============================================================================
# TAB 2: DAILY & BRANCH  (unchanged from prior version)
# ==============================================================================
with tab2:
    st.markdown("### Daily Collections & Branch Breakdown")
    if not df_daily.empty:
        d = df_daily[df_daily['Month'].isin(selected_months)]
        if not d.empty:
            daily_by_day = d.groupby('Day')['Total'].sum().reset_index()
            fig = go.Figure(go.Scatter(
                x=daily_by_day['Day'], y=daily_by_day['Total'],
                mode='lines+markers', line=dict(color='#00B4D8', width=3),
                marker=dict(size=8), fill='tozeroy', fillcolor='rgba(0,180,216,0.15)'
            ))
            st.plotly_chart(fig, use_container_width=True)
    if not df_offices.empty:
        o = df_offices[df_offices['Month'].isin(selected_months)]
        if not o.empty:
            branch_sum = o.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
            st.plotly_chart(px.bar(branch_sum, x='Office', y='Total',
                                   color='Total', color_continuous_scale='Blues'),
                            use_container_width=True)

# ==============================================================================
# TAB 3: SECTOR DEEP-DIVE  (uses the shared cash-basis helper — this is why
#                          the toggle now actually changes the figures)
# ==============================================================================
with tab3:
    st.markdown("### ISIC Sector Deep-Dive")
    if df_isic.empty:
        st.info("No ISIC data loaded.")
    else:
        di = df_isic[df_isic['Month'].isin(selected_months)].copy()
        di = apply_cash_basis(di, cash_basis)

        if di.empty:
            st.info(f"No rows match the current filter (Cash Basis: {cash_basis}).")
        else:
            # Aggregate per sector
            sector_totals = (di.groupby('ISIC_Sector')['Amount'].sum()
                             .reset_index().sort_values('Amount', ascending=False))
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown(f"#### Top 15 Sectors — {cash_basis}")
                top15 = sector_totals.head(15).copy()
                top15['Amount_Label'] = top15['Amount'].apply(fmt_currency)
                st.dataframe(top15[['ISIC_Sector', 'Amount_Label']].rename(
                    columns={'ISIC_Sector': 'Sector', 'Amount_Label': 'Collected'}),
                    use_container_width=True, hide_index=True)
            with c2:
                fig = px.bar(sector_totals.head(15), x='Amount', y='ISIC_Sector', orientation='h',
                             color='Amount', color_continuous_scale='Blues',
                             title=f"Top Sectors ({cash_basis})")
                fig.update_layout(height=500, yaxis={'categoryorder': 'total ascending'},
                                  margin=dict(t=40, b=10, l=10, r=10))
                st.plotly_chart(fig, use_container_width=True)

            # Per-tax-head breakdown, current vs arrears split shown explicitly
            st.markdown(f"#### Sector Detail by Tax Head — {cash_basis}")
            detail = di.groupby(['Tax_Type', 'ISIC_Sector']).agg(
                Current=('Current', 'sum'),
                Arrears=('Arrears', 'sum'),
                Amount=('Amount', 'sum'),
            ).reset_index().sort_values(['Tax_Type', 'Amount'], ascending=[True, False])
            detail_disp = detail.copy()
            for c in ['Current', 'Arrears', 'Amount']:
                detail_disp[c] = detail_disp[c].apply(fmt_currency)
            st.dataframe(detail_disp, use_container_width=True, hide_index=True)

# ==============================================================================
# TAB 4: REPORTS
# ==============================================================================
with tab4:
    st.markdown(f"### Variance Notes — {snap_month} (FY {snap_fiscal})")
    st.caption("Departmental comments saved in the sidebar for this exact month and fiscal year "
               "are reproduced verbatim in Section 2 of the report.")
    if st.button("Generate Report"):
        docx_bytes = build_variance_notes_docx(
            snap_month, snap_fiscal,
            total_actual, total_proj, total_var,
            ytd, df_isic,
            comments=st.session_state.get("comments_store", {}),
        )
        st.download_button(
            "📥 Download Word Document",
            docx_bytes,
            f"Variance_Notes_{snap_month.replace(' ', '_')}_{snap_fiscal.replace('/', '-')}.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
