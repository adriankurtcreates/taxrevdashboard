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
# PAGE CONFIG & CSS
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

# ==============================================================================
# CONSTANTS
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
QUARTER_END_MONTHS = {"June": "Q1", "September": "Q2", "December": "Q3", "March": "Q4"}

# ==============================================================================
# HELPERS
# ==============================================================================
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
        m_pos = {"April": "first", "May": "second", "June": "third"}.get(month, "")
        return f"This month represents the {m_pos} month of the 1st quarter of the fiscal year."
    elif q == "Q2 (Jul-Sep)":
        m_pos = {"July": "first", "August": "second", "September": "third"}.get(month, "")
        return f"This month represents the {m_pos} month of the 2nd quarter of the fiscal year."
    elif q == "Q3 (Oct-Dec)":
        m_pos = {"October": "first", "November": "second", "December": "third"}.get(month, "")
        return f"This month represents the {m_pos} month of the 3rd quarter of the fiscal year."
    else:
        m_pos = {"January": "first", "February": "second", "March": "third"}.get(month, "")
        return f"This month represents the {m_pos} month of the 4th quarter of the fiscal year."

# ==============================================================================
# PARSERS
# ==============================================================================
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

# ==============================================================================
# VARIANCE NOTES DOCX GENERATOR (STRICT TEMPLATE)
# ==============================================================================
def generate_tax_paragraph(code, account, actual, projection, variance):
    if projection > 0:
        pct = (variance / projection) * 100
    else:
        pct = 0.0

    var_str = fmt_currency(variance)
    pct_str = fmt_pct(pct)

    if variance == 0:
        return f"Actual collections of {fmt_currency(actual)} matched the projected figure of {fmt_currency(projection)}. There is no variance to report."
    
    if variance > 0:
        direction = "exceeded"
        size = "an overage"
    else:
        direction = "fell short of"
        size = "a shortfall"

    sentence = (
        f"Actual collections of {fmt_currency(actual)} {direction} the projected figure of "
        f"{fmt_currency(projection)}, representing {size} of {var_str} or {pct_str}. "
    )
    if projection > 0:
        rate = (actual / projection) * 100
        sentence += f"This corresponds to a collection rate of {rate:.2f}% of the projection for the period."
    return sentence

def build_variance_notes_docx(
    snap_month, total_actual, total_projection, total_variance, ytd_df, df_isic
):
    doc = Document()

    # --- HEADER ---
    h = doc.add_heading('SAINT LUCIA', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('DEPARTMENT OF FINANCE', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = doc.add_heading('INLAND REVENUE DEPARTMENT', 1); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph('Revenue Collection Volume 14: Monthly Performance Report')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(f'Explanation of Variances for the Month of {snap_month} 2026')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(14)

    # --- OVERVIEW ---
    doc.add_heading('Overview', level=1)
    
    direction = "an overage" if total_variance >= 0 else "a shortfall"
    pct_total = (total_variance / total_projection) * 100 if total_projection > 0 else 0

    overview_text = (
        f"For the month of {snap_month} 2026, the Inland Revenue Department projected total revenue "
        f"of {fmt_currency(total_projection)}. Actual collections amounted to {fmt_currency(total_actual)}, "
        f"resulting in {direction} of {fmt_currency(abs(total_variance))} or {fmt_pct(abs(pct_total))} "
        f"{'above' if total_variance >= 0 else 'below'} projection. "
    )
    doc.add_paragraph(overview_text)
    
    doc.add_paragraph(get_fiscal_context(snap_month))
    
    # Placeholder for the general narrative
    narrative = (
        f"The aggregate performance for {snap_month} shows {'positive' if total_variance >= 0 else 'negative'} "
        f"net variance. The detailed distribution of performance across individual tax heads is presented in the Variance Analysis section below."
    )
    doc.add_paragraph(narrative)

    # --- VARIANCE ANALYSIS ---
    doc.add_heading('Variance Analysis by Tax Type', level=1)
    doc.add_paragraph("The performances by individual tax type are outlined below:")

    for _, row in ytd_df.sort_values('Code').iterrows():
        code = row['Code']
        account = row['Account']
        actual = row['Actual_YTD']
        projection = row['YTD_Projection']
        variance = row['Variance']
        pct = (variance / projection * 100) if projection > 0 else 0.0

        # Heading
        doc.add_heading(f"{code} - {account}", level=2)
        
        # Projection/Actual/Variance Line
        var_str = fmt_currency(variance)
        pct_str = fmt_pct(pct)
        line = f"Projection: {fmt_currency(projection)} | Actual: {fmt_currency(actual)} | Variance: {var_str} ({pct_str})"
        p = doc.add_paragraph(line)
        p.runs[0].bold = True

        # Analytical Paragraph
        doc.add_paragraph(generate_tax_paragraph(code, account, actual, projection, variance))

        # Special Rules
        if code == "3111 001":
            doc.add_paragraph("Note: This tax head includes Personal Income Tax, Contract Tax, and PAYE.")
        elif code == "3131 001":
            doc.add_paragraph(
                "Note: Due to the ongoing moratorium on residential property tax, current payments "
                "apply exclusively to commercial properties. Arrears collections apply to both."
            )
        elif code == "3141 004":
            doc.add_paragraph(
                "Note: Do not label this tax as obsolete. The Hotel Accommodation Tax remains an active "
                "registry line for the collection and settlement of legacy arrears, which accounts for "
                "ongoing revenue despite no active monthly projections."
            )
        elif code == "3141 006":
            doc.add_paragraph(
                "Note: The Health and Citizen Security Levy (HCSL) was implemented in March 2023. "
                "Therefore, no historical data exists prior to that date for trend comparison."
            )
        elif code == "3142 006":
            doc.add_paragraph(
                "Structural Insight: The top-performing sectors listed for Stamp Duty may not be a true "
                "indicator of which economic sectors are generating the underlying transactions. This is "
                "because the ISIC code used to log the payment is routinely assigned to the law firm or "
                "legal practitioner paying the tax on behalf of their clients."
            )
        elif code == "3422 008":
            doc.add_paragraph(
                "Note: Due to the active Amnesty extending to December 2027, collections are minimal and "
                "result primarily from the waiver of interest, penalties, and fees."
            )

        # Top Performing Sectors
        if df_isic is not None and not df_isic.empty:
            sheet_prefix = None
            prefix_map = {"3111": "PIT", "3112": "CIT", "3113": "Withholding", "3131": "Property Tax",
                          "3141 001": "VAT", "3141 004": "HAT", "3141 005": "Travel Tax", "3141 006": "HCSL",
                          "3142 005": "Insurance Premium", "3142 006": "Stamp Duty"}
            for prefix, sheet in prefix_map.items():
                if code.startswith(prefix):
                    sheet_prefix = sheet
                    break
            
            if sheet_prefix:
                filtered = df_isic[df_isic['Tax_Type'].str.contains(sheet_prefix, case=False, na=False)]
                filtered = filtered[filtered['Month'] == snap_month]
                if not filtered.empty:
                    top3 = filtered.groupby('ISIC_Sector')['Total'].sum().nlargest(3)
                    doc.add_heading('Top Performing Sectors', level=3)
                    for sector, amount in top3.items():
                        doc.add_paragraph(f"{sector}: {fmt_currency(amount)}", style='List Bullet')
                    
                    # Required closing phrase
                    doc.add_paragraph(
                        "These sectors accounted for the largest share of sector-attributed collections "
                        "for the reporting period. Further observation of these possibilities will have to be made."
                    )

    # --- CONCLUSION ---
    doc.add_heading('Conclusion', level=1)
    
    if total_variance >= 0:
        verdict = "positive at the aggregate level"
    else:
        verdict = "negative at the aggregate level"

    conclusion_text = (
        f"The Inland Revenue Department's performance for {snap_month} 2026 produced a net variance "
        f"of {fmt_currency(total_variance)} ({fmt_pct(pct_total)}), which is {verdict}. "
    )
    doc.add_paragraph(conclusion_text)
    
    if not ytd_df.empty:
        top = ytd_df.nlargest(1, 'Variance').iloc[0]
        bottom = ytd_df.nsmallest(1, 'Variance').iloc[0]
        doc.add_paragraph(
            f"The largest positive variance was recorded in {top['Account']} "
            f"({fmt_currency(top['Variance'])}). The largest negative variance was recorded in "
            f"{bottom['Account']} ({fmt_currency(bottom['Variance'])})."
        )
        
    doc.add_paragraph(
        "These variances reflect the current collection status against the calendarized projections. "
        "The Department will continue to monitor these streams to assess whether the observed variances "
        "represent timing differences or structural shifts in compliance."
    )

    bio = _io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==============================================================================
# SIDEBAR & MAIN RENDERING (SAME AS PREVIOUS)
# ==============================================================================
# ... [Keep the same sidebar uploaders, diagnostics, filters, KPI cards, and tabs 1-3 as the previous code] ...

# TAB 4: REPORTS
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
