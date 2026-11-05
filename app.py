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
    "3111 001": "Personal Income Tax", "3111 002": "Arrears on PIT",
    "3112 001": "Corporate Income Tax", "3112 002": "Arrears on CIT",
    "3113 001": "Withholding Tax", "3113 002": "Arrears on Withholding",
    "3131 001": "Property Tax", "3141 001": "Value Added Tax (VAT)",
    "3141 004": "Hotel Accommodation Tax", "3141 005": "Domestic Sales Tax on Travel",
    "3141 006": "Health & Citizen Security Levy", "3142 005": "Insurance Premium Tax",
    "3142 006": "Stamp Duty - Inland Revenue", "3143 010": "Telecommunication Class Licence",
    "3151 009": "Passenger Facility Fee", "3422 008": "Collection Fees - Towns & Villages",
}
QUARTER_END_MONTHS = {"June": "Q1", "September": "Q2", "December": "Q3", "March": "Q4"}

# ==============================================================================
# HELPERS
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
# PARSER: MONTHLY SLICES (for trends)
# ==============================================================================
@st.cache_data
def parse_monthly_slices(file_payload):
    all_tax = []
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            rev_sheet = next((s for s in xls.sheet_names if "actual monthly revenue" in s.lower()), None)
            if not rev_sheet: continue
            month = find_month_in_text(filename) or find_month_in_text(rev_sheet)
            if not month: continue
            df = pd.read_excel(file_obj, sheet_name=rev_sheet, header=None)
            for i in range(5, min(35, len(df))):
                code = normalize_code(df.iloc[i, 0])
                name = df.iloc[i, 1]
                if not code or pd.isna(name): continue
                if not re.match(r'^\d{4}\s\d{3}$', code): continue
                approved = safe_number(df.iloc[i, 2])
                monthly = sum(safe_number(df.iloc[i, c]) for c in range(4, 15))
                all_tax.append({
                    "Month": month, "Quarter": get_quarter(month),
                    "Code": code, "Account": str(name).strip(),
                    "Monthly_Actual": monthly
                })
        except Exception:
            pass
    return pd.DataFrame(all_tax)

# ==============================================================================
# PARSER: SUMMARY SNAPSHOT (YTD actuals, authoritative)
# ==============================================================================
@st.cache_data
def parse_summary_snapshot(file_payload):
    all_snap = []
    diag = []
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
                diag.append({"file": filename, "error": "month not found"})
                continue

            summary_sheet = next((s for s in xls.sheet_names if "revenue summary" in s.lower()), None)
            if not summary_sheet:
                summary_sheet = next((s for s in xls.sheet_names if "summary" in s.lower()), None)
            if not summary_sheet:
                diag.append({"file": filename, "error": "no summary sheet"})
                continue

            df = pd.read_excel(file_obj, sheet_name=summary_sheet, header=None)

            header_row = 4
            for r in range(15):
                row_str = " ".join(str(v).upper() for v in df.iloc[r].values if pd.notna(v))
                if "TO DATE" in row_str:
                    header_row = r
                    break

            rows = 0
            for i in range(header_row + 1, len(df)):
                raw_code = df.iloc[i, 0]
                desc = df.iloc[i, 1]
                if pd.isna(raw_code) or pd.isna(desc): continue
                code = normalize_code(raw_code)
                if not re.match(r'^\d{4}\s\d{3}$', code): continue

                approved = safe_number(df.iloc[i, 2])
                actual_ytd = safe_number(df.iloc[i, 4]) if 4 < len(df.columns) else 0
                variance = safe_number(df.iloc[i, 13]) if 13 < len(df.columns) else 0

                all_snap.append({
                    "Month": month, "Quarter": get_quarter(month),
                    "Code": code, "Account": str(desc).strip(),
                    "Approved_Annual": approved,
                    "Actual_YTD": actual_ytd,
                    "Snapshot_Variance": variance,
                })
                rows += 1
            diag.append({"file": filename, "month": month, "rows": rows})
        except Exception as e:
            diag.append({"file": filename, "error": str(e)})

    return pd.DataFrame(all_snap), pd.DataFrame(diag)

# ==============================================================================
# PARSER: CALENDARIZED PROJECTIONS
# ==============================================================================
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

        MONTH_COLS = {
            "April": 3, "May": 4, "June": 5,
            "July": 7, "August": 8, "September": 9,
            "October": 11, "November": 12, "December": 13,
            "January": 15, "February": 16, "March": 17,
        }
        all_proj = []
        for i in range(7, min(55, len(df))):
            raw_code = df.iloc[i, 0]
            desc = df.iloc[i, 1]
            if pd.isna(raw_code) or pd.isna(desc): continue
            code = normalize_code(raw_code)
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

# ==============================================================================
# PARSER: DAILY + OFFICES
# ==============================================================================
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
                total = safe_number(df.iloc[i, 20])
                if total > 0:
                    all_daily.append({"Month": month, "Day": pd.to_datetime(dv).day, "Total": total})
            for i in range(30, 55):
                on = df.iloc[i, 0]
                ot = safe_number(df.iloc[i, 20])
                if pd.notna(on) and isinstance(on, str):
                    s = on.strip()
                    if any(x in s for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        all_offices.append({"Month": month, "Office": s, "Total": ot})
                    elif "Sub Total" in s:
                        all_offices.append({"Month": month, "Office": "Castries (Sub Total)", "Total": ot})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(all_offices)

# ==============================================================================
# PARSER: ISIC
# ==============================================================================
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
                                         "Current": cur, "Arrears": arr, "Total": cur + arr})
        except Exception:
            pass
    return pd.DataFrame(all_isic)

# ==============================================================================
# SIDEBAR: UPLOADERS
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    calib_file = st.file_uploader("Calendarized Projections (.xlsx)", type=["xlsx"])

# ==============================================================================
# MAIN
# ==============================================================================
st.title("🇱🇨 Revenue Intelligence Dashboard")

if not main_files:
    st.info("👈 Upload your Monthly Revenue Reports in the sidebar to begin.")
    st.stop()

# Parse all
main_payload = [(f.name, f.getvalue()) for f in main_files]
df_monthly = parse_monthly_slices(main_payload)
df_snapshot, diag_snap = parse_summary_snapshot(main_payload)
df_daily, df_offices = parse_daily(main_payload)

df_isic = pd.DataFrame()
if isic_files:
    isic_payload = [(f.name, f.getvalue()) for f in isic_files]
    df_isic = parse_isic(isic_payload)

df_calib = pd.DataFrame()
calib_diag = {}
if calib_file:
    calib_payload = (calib_file.name, calib_file.getvalue())
    df_calib, calib_diag = parse_calendarized(calib_payload)

# DIAGNOSTIC — always visible so we can verify
with st.expander("🔍 Data Diagnostics", expanded=True):
    st.write(f"**Monthly slices:** {len(df_monthly)} rows")
    st.write(f"**Snapshot rows:** {len(df_snapshot)}")
    if not df_snapshot.empty:
        st.write(f"**Σ Actual_YTD from parser:** ${df_snapshot['Actual_YTD'].sum():,.2f}")
        st.write(f"**Months in snapshot:** {sorted(df_snapshot['Month'].unique())}")
        st.write(f"**Codes parsed:** {sorted(df_snapshot['Code'].unique())}")
    st.dataframe(diag_snap, use_container_width=True)
    if calib_file:
        if calib_diag.get("error"):
            st.error(f"Projections error: {calib_diag['error']}")
        else:
            st.success(f"✓ Projections loaded: {calib_diag['rows_parsed']} rows")
    else:
        st.info("No projections file uploaded — YTD projection will fall back to Summary variance.")

if df_snapshot.empty:
    st.error("Could not parse Summary snapshots. Check diagnostics above.")
    st.stop()

# Slicer options
avail_months = sorted(df_snapshot['Month'].unique(), key=month_order)
avail_quarters = sorted(df_snapshot['Quarter'].unique())
avail_codes = sorted(df_snapshot['Code'].unique())
tax_opts = [f"{c} — {TAX_HEADS.get(c, 'Unknown')}" for c in avail_codes]
display_to_code = {f"{c} — {TAX_HEADS.get(c, 'Unknown')}": c for c in avail_codes}

# Sidebar filters
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

# Apply filters
if not selected_months:
    st.warning("⚠️ No months selected."); st.stop()
if not selected_quarters:
    st.warning("⚠️ No quarters selected."); st.stop()
if not selected_taxes:
    st.warning("⚠️ No tax heads selected."); st.stop()

selected_codes = [display_to_code[d] for d in selected_taxes if d in display_to_code]

# Snapshot selection
snap_filtered = df_snapshot[
    df_snapshot['Month'].isin(selected_months) &
    df_snapshot['Quarter'].isin(selected_quarters) &
    df_snapshot['Code'].isin(selected_codes)
].copy()

if snap_filtered.empty:
    st.error("Filters produced 0 snapshot rows."); st.stop()

# Pick authoritative snapshot
if analysis_mode == "YTD (Latest Summary)":
    snap_month = sorted(snap_filtered['Month'].unique(), key=month_order)[-1]
elif analysis_mode == "Quarter-End Snapshot":
    qe_avail = [m for m in snap_filtered['Month'].unique() if m in QUARTER_END_MONTHS]
    if not qe_avail:
        st.error("No quarter-end snapshot available."); st.stop()
    snap_month = sorted(qe_avail, key=month_order)[-1]
else:
    snap_month = sorted(snap_filtered['Month'].unique(), key=month_order)[-1]

snapshot = snap_filtered[snap_filtered['Month'] == snap_month].copy()

# YTD actuals
ytd = snapshot[['Code', 'Account', 'Actual_YTD', 'Approved_Annual', 'Snapshot_Variance']].copy()

# YTD projection
if not df_calib.empty:
    months_through = FISCAL_MONTHS[:month_order(snap_month) + 1]
    proj_win = df_calib[
        df_calib['Month'].isin(months_through) & df_calib['Code'].isin(selected_codes)
    ]
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

# Sanity check caption
snap_var_sum = snapshot['Snapshot_Variance'].sum()
st.caption(f"🔍 Data check → ΣActual: ${total_actual:,.2f} | ΣSummary variance: ${snap_var_sum:,.2f} | ΣDashboard variance: ${total_var:,.2f} | Codes: {len(selected_codes)}")

top_perf = ytd.nlargest(1, 'Variance').iloc[0]
top_short = ytd.nsmallest(1, 'Variance').iloc[0]

is_qe = snap_month in QUARTER_END_MONTHS
badge = "snapshot-badge" if is_qe else "snapshot-badge-warn"
badge_txt = f"✓ Quarter-End: {snap_month}" if is_qe else f"Latest: {snap_month}"

st.markdown(f"""
<div class="narrative-box">
    <b>📊 Executive Summary</b>
    <span class="{badge}">{badge_txt}</span><br><br>
    <b>Source:</b> {snap_month} 2026 Revenue Summary<br>
    <b>Projection:</b> {proj_source}<br>
    <b>YTD Actual:</b> ${total_actual:,.2f} | <b>YTD Projection:</b> ${total_proj:,.2f}<br>
    <b>Net {'Overage' if total_var > 0 else 'Shortage'}:</b> ${abs(total_var):,.2f} ({(total_var/total_proj*100) if total_proj>0 else 0:.2f}%)<br><br>
    <b>Top Driver:</b> {top_perf['Account']} (+${top_perf['Variance']:,.2f})<br>
    <b>Main Drag:</b> {top_short['Account']} (${top_short['Variance']:,.2f})
</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
c1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Actual</div><div class="kpi-value">${total_actual:,.0f}</div><div class="kpi-delta">Source: {snap_month}</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Projection</div><div class="kpi-value">${total_proj:,.0f}</div><div class="kpi-delta">{proj_source}</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_var > 0 else "#DC3545"}">${total_var:,.0f}</div></div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📈 Macro Trends", "📅 Daily & Office", "🏭 Sector (ISIC)", "📝 Reports"])

with tab1:
    df_m = df_monthly[
        df_monthly['Month'].isin(selected_months) &
        df_monthly['Quarter'].isin(selected_quarters) &
        df_monthly['Code'].isin(selected_codes)
    ]
    if not df_m.empty:
        st.markdown("### Tax Head Composition Over Time")
        mc = df_m.groupby(['Month', 'Account'])['Monthly_Actual'].sum().reset_index()
        mc['Order'] = mc['Month'].map(month_order)
        mc = mc.sort_values('Order')
        st.plotly_chart(px.area(mc, x='Month', y='Monthly_Actual', color='Account', groupnorm='fraction'), use_container_width=True)

        st.markdown("### MoM Growth")
        mom = df_m.groupby(['Month', 'Account'])['Monthly_Actual'].sum().unstack(fill_value=0)
        st.dataframe(mom.pct_change().replace([float('inf'), float('-inf')], 0).fillna(0).style.format("{:.1%}"), use_container_width=True)
    else:
        st.info("No monthly slices available for these filters.")

with tab2:
    if not df_daily.empty:
        d = df_daily[df_daily['Month'].isin(selected_months)]
        if not d.empty:
            st.markdown("### Daily Collection Heatmap")
            hm = d.pivot_table(index='Month', columns='Day', values='Total', aggfunc='sum').fillna(0)
            st.plotly_chart(px.imshow(hm, aspect="auto", color_continuous_scale="Blues"), use_container_width=True)
    if not df_offices.empty:
        o = df_offices[df_offices['Month'].isin(selected_months)]
        if not o.empty:
            st.markdown("### Office Performance")
            s = o.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
            st.plotly_chart(px.bar(s, x='Office', y='Total', color='Total', color_continuous_scale='Blues'), use_container_width=True)
    if df_daily.empty and df_offices.empty:
        st.info("Daily/office data unavailable.")

with tab3:
    if not df_isic.empty:
        di = df_isic[df_isic['Month'].isin(selected_months)].copy()
        if cash_basis == "Current Only": di['Amount'] = di['Current']
        elif cash_basis == "Arrears Only": di['Amount'] = di['Arrears']
        else: di['Amount'] = di['Total']
        di = di[di['Amount'] > 0]
        if not di.empty:
            tax_pick = st.selectbox("Tax Head", di['Tax_Type'].unique())
            drill = di[di['Tax_Type'] == tax_pick].groupby('ISIC_Sector')['Amount'].sum().nlargest(15).reset_index()
            st.plotly_chart(px.bar(drill, x='Amount', y='ISIC_Sector', orientation='h'), use_container_width=True)
    else:
        st.info("Upload ISIC files to unlock sector data.")

with tab4:
    st.markdown(f"### Variance Notes — {snap_month} 2026")
    if st.button(f"Generate Report"):
        doc = Document()
        doc.add_heading('SAINT LUCIA — DEPARTMENT OF FINANCE', 0)
        doc.add_heading(f'Variance Notes — {snap_month} 2026', 1)
        doc.add_paragraph(f"Total YTD Actual: ${total_actual:,.2f}")
        doc.add_paragraph(f"Total YTD Projection: ${total_proj:,.2f}")
        doc.add_paragraph(f"Variance: ${total_var:,.2f}")
        for _, r in ytd.sort_values('Code').iterrows():
            doc.add_heading(f"{r['Code']} — {r['Account']}", level=2)
            doc.add_paragraph(f"Actual: ${r['Actual_YTD']:,.2f} | Projection: ${r['YTD_Projection']:,.2f} | Variance: ${r['Variance']:,.2f}")
        bio = _io.BytesIO()
        doc.save(bio)
        st.download_button("📥 Download", bio.getvalue(), f"Variance_Notes_{snap_month}.docx")
