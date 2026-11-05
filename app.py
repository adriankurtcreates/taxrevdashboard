import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io as _io
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==============================================================================
# 1. PAGE CONFIG & CUSTOM CSS
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
    .chart-container { background-color: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .narrative-box { background-color: #f8f9fa; border-left: 5px solid #00B4D8; padding: 15px; margin-bottom: 20px; border-radius: 5px; font-size: 16px; }
    .snapshot-badge { display: inline-block; background-color: #00B4D8; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .snapshot-badge-warn { display: inline-block; background-color: #F59E0B; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. CONSTANTS & HELPERS
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
    "3151 009": "Passenger Facility Fee", "3422 008": "Collection Fees - Towns & Villages"
}

# Quarter-end snapshots (only these months are authoritative for quarterly figures)
QUARTER_END_MONTHS = {"June": "Q1", "September": "Q2 (Mid-Year)", "December": "Q3", "March": "Q4 (Year-End)"}

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

def month_order(month_name):
    """Returns fiscal order index (April=0, March=11)."""
    return FISCAL_MONTHS.index(month_name) if month_name in FISCAL_MONTHS else 99

# ==============================================================================
# 3. DATA PARSERS
# ==============================================================================
@st.cache_data
def parse_all_main_data(file_payload):
    """
    Parses all monthly reports and builds:
    - df_tax: All monthly data (for trends)
    - df_snapshots: Only authoritative summary snapshots (latest + quarter ends)
    """
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
            header_row = 2
            
            for i in range(header_row + 3, min(header_row + 35, len(df))):
                code = str(df.iloc[i, 0]).strip()
                name = str(df.iloc[i, 1]).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code) or "total" in name.lower(): continue
                
                approved = safe_number(df.iloc[i, 2])
                monthly_actual = sum(safe_number(df.iloc[i, c]) for c in range(4, 15))
                ytd_actual = safe_number(df.iloc[i, 15]) or monthly_actual
                
                all_tax.append({
                    "Month": month, "Quarter": get_quarter(month), "Code": code, "Account": name,
                    "Approved_Annual": approved, "Monthly_Actual": monthly_actual, "YTD_Actual": ytd_actual
                })
            diagnostics.append({"file": filename, "month": month, "status": "OK"})
        except Exception as e:
            diagnostics.append({"file": filename, "month": None, "status": f"Error: {e}"})
    return pd.DataFrame(all_tax), pd.DataFrame(diagnostics)

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
                    all_daily.append({
                        "Month": month, "Date": pd.to_datetime(date_val).date(),
                        "Day": pd.to_datetime(date_val).day, "Total": total_col
                    })
            
            for i in range(30, 55):
                office_name = df.iloc[i, 0]
                office_total = safe_number(df.iloc[i, 20])
                if pd.notna(office_name) and isinstance(office_name, str):
                    office_str = office_name.strip()
                    if any(x in office_str for x in ["Office", "Online", "Payroll", "Treasury", "City Council"]):
                        all_offices.append({"Month": month, "Office": office_str, "Total": office_total})
                    elif "Sub Total" in office_str:
                        all_offices.append({"Month": month, "Office": "Castries (Sub Total)", "Total": office_total})
        except Exception:
            pass
    return pd.DataFrame(all_daily), pd.DataFrame(all_offices)

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
                    current_amt = safe_number(df.iloc[i, current_col]) if current_col is not None else 0
                    arrears_amt = safe_number(df.iloc[i, arrears_col]) if arrears_col is not None else 0
                    if current_amt > 0 or arrears_amt > 0:
                        all_isic.append({
                            "Month": month, "Tax_Type": sheet, "ISIC_Sector": str(sector).strip(),
                            "Current": current_amt, "Arrears": arrears_amt, "Total": current_amt + arrears_amt
                        })
        except Exception:
            pass
    return pd.DataFrame(all_isic)

# ==============================================================================
# 4. SNAPSHOT PRIORITY ENGINE
# ==============================================================================
def select_authoritative_snapshot(df_tax, selected_months, analysis_mode):
    """
    Returns the authoritative snapshot dataframe based on analysis mode.
    
    Modes:
      - "YTD (Latest Summary)": uses the MOST RECENT uploaded month
      - "Quarter-End Snapshot": uses only quarter-end months (Jun, Sep, Dec, Mar)
      - "Month Snapshot": uses the latest month within filter
    """
    filtered = df_tax[df_tax['Month'].isin(selected_months)]
    if filtered.empty:
        return None, None
    
    if analysis_mode == "YTD (Latest Summary)":
        # Latest uploaded month wins — captures most recent cumulative state
        latest = sorted(filtered['Month'].unique(), key=month_order)[-1]
        snapshot = filtered[filtered['Month'] == latest]
        return snapshot, latest
    
    elif analysis_mode == "Quarter-End Snapshot":
        # Only use quarter-end months (Jun, Sep, Dec, Mar) that were uploaded
        available_quarter_ends = [m for m in filtered['Month'].unique() if m in QUARTER_END_MONTHS]
        if not available_quarter_ends:
            return None, None
        # Latest quarter-end available
        latest_qe = sorted(available_quarter_ends, key=month_order)[-1]
        snapshot = filtered[filtered['Month'] == latest_qe]
        return snapshot, latest_qe
    
    else:  # "Month Snapshot"
        latest = sorted(filtered['Month'].unique(), key=month_order)[-1]
        snapshot = filtered[filtered['Month'] == latest]
        return snapshot, latest

def compute_ytd_from_snapshot(snapshot_df):
    """From a single authoritative snapshot, compute YTD per account."""
    return snapshot_df.groupby(['Code', 'Account']).agg(
        YTD_Actual=('YTD_Actual', 'first'),   # YTD is cumulative — use as-is
        Approved_Annual=('Approved_Annual', 'first')
    ).reset_index()

# ==============================================================================
# 5. SIDEBAR
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    
    if main_files:
        st.markdown("### 🎯 Analysis Mode")
        analysis_mode = st.radio(
            "Which snapshot should drive YTD numbers?",
            ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"],
            help="YTD uses the most recent Revenue Summary. Quarter-End uses only Jun/Sep/Dec/Mar summaries."
        )
        
        st.markdown("---")
        selected_months = st.multiselect("Month Slicer", FISCAL_MONTHS, default=FISCAL_MONTHS)
        available_quarters = list(set([get_quarter(m) for m in selected_months]))
        selected_quarters = st.multiselect("Quarter Slicer", available_quarters, default=available_quarters)
        tax_options = list(TAX_HEADS.values())
        selected_taxes = st.multiselect("Tax Head Slicer", tax_options, default=tax_options)
        cash_basis = st.radio("Cash Basis Focus", ["All", "Current Only", "Arrears Only"], horizontal=True)

# ==============================================================================
# 6. MAIN EXECUTION
# ==============================================================================
st.title("🇱🇨 Revenue Intelligence Dashboard")

if main_files:
    main_payload = [(f.name, f.getvalue()) for f in main_files]
    df_tax, diag_main = parse_all_main_data(main_payload)
    df_daily, df_offices = parse_daily_collections(main_payload)
    
    df_isic = pd.DataFrame()
    if isic_files:
        isic_payload = [(f.name, f.getvalue()) for f in isic_files]
        df_isic = parse_isic_deep(isic_payload)

    if not df_tax.empty:
        # --- Apply Slicers ---
        df_f = df_tax[df_tax['Month'].isin(selected_months)]
        df_f = df_f[df_f['Quarter'].isin(selected_quarters)]
        selected_codes = [code for code, name in TAX_HEADS.items() if name in selected_taxes]
        df_f = df_f[df_f['Code'].isin(selected_codes)]
        
        if df_f.empty:
            st.warning("No data matches the selected slicers.")
        else:
            # --- SNAPSHOT PRIORITY ENGINE ---
            snapshot_df, snapshot_month = select_authoritative_snapshot(df_f, selected_months, analysis_mode)
            
            if snapshot_df is None or snapshot_df.empty:
                st.error(f"No {analysis_mode} available. "
                         f"Upload the required month or change the analysis mode.")
            else:
                months_elapsed = month_order(snapshot_month) + 1
                ytd = compute_ytd_from_snapshot(snapshot_df)
                ytd['Pro_Rata'] = ytd['Approved_Annual'] * (months_elapsed / 12)
                ytd['Variance'] = ytd['YTD_Actual'] - ytd['Pro_Rata']
                
                total_ytd = ytd['YTD_Actual'].sum()
                total_pro_rata = ytd['Pro_Rata'].sum()
                total_var = total_ytd - total_pro_rata
                
                top_perf = ytd.nlargest(1, 'Variance').iloc[0] if not ytd.empty else None
                top_short = ytd.nsmallest(1, 'Variance').iloc[0] if not ytd.empty else None
                
                # --- Snapshot Badge ---
                is_quarter_end = snapshot_month in QUARTER_END_MONTHS
                badge_class = "snapshot-badge" if is_quarter_end else "snapshot-badge-warn"
                badge_text = f"✓ Quarter-End Authoritative: {snapshot_month}" if is_quarter_end else f"Latest Snapshot: {snapshot_month}"
                
                st.markdown(f"""
                <div class="narrative-box">
                    <b>📊 Executive Summary</b>
                    <span class="{badge_class}">{badge_text}</span><br><br>
                    <b>Data Source:</b> <i>{snapshot_month} 2026 Revenue Summary</i> (Mode: {analysis_mode})<br>
                    Total YTD Revenue: <b>${total_ytd:,.0f}</b> against pro-rata target of <b>${total_pro_rata:,.0f}</b>.<br>
                    Net {'overage' if total_var > 0 else 'shortage'}: <b>${abs(total_var):,.0f}</b> ({(total_var/total_pro_rata*100):.1f}%).<br><br>
                    <b>Top Driver:</b> {top_perf['Account']} (${top_perf['Variance']:,.0f}).<br>
                    <b>Main Drag:</b> {top_short['Account']} (${abs(top_short['Variance']):,.0f}).
                </div>
                """, unsafe_allow_html=True)

                # --- Informational Callout for Prior Months ---
                if analysis_mode == "YTD (Latest Summary)":
                    available_qe = [m for m in df_f['Month'].unique() if m in QUARTER_END_MONTHS]
                    if available_qe:
                        st.caption(f"ℹ️ Prior quarter-end snapshots available: {', '.join(available_qe)}. Switch to 'Quarter-End Snapshot' mode to view historical quarter closes.")

                # --- KPI CARDS ---
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Revenue</div><div class="kpi-value">${total_ytd:,.0f}</div><div class="kpi-delta">Source: {snapshot_month}</div></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="kpi-card"><div class="kpi-title">Pro-Rata Target</div><div class="kpi-value">${total_pro_rata:,.0f}</div><div class="kpi-delta">{months_elapsed}/12 Months</div></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_var > 0 else "#DC3545"}">${total_var:,.0f}</div></div>', unsafe_allow_html=True)

                # --- TABS ---
                tab_trends, tab_daily, tab_isic, tab_reports = st.tabs(["📈 Macro Trends", "📅 Daily & Office", "🏭 Sector (ISIC)", "📝 Reports"])

                with tab_trends:
                    st.markdown("### Tax Head Composition Over Time")
                    monthly_comp = df_f.groupby(['Month', 'Account'])['Monthly_Actual'].sum().reset_index()
                    monthly_comp['Order'] = monthly_comp['Month'].map(month_order)
                    monthly_comp = monthly_comp.sort_values('Order')
                    st.plotly_chart(px.area(monthly_comp, x='Month', y='Monthly_Actual', color='Account', groupnorm='fraction'), use_container_width=True)

                    st.markdown("### Month-over-Month (MoM) Growth Rates")
                    mom = df_f.groupby(['Month', 'Account'])['Monthly_Actual'].sum().unstack(fill_value=0)
                    mom_pct = mom.pct_change().replace([float('inf'), float('-inf')], 0) * 100
                    st.dataframe(mom_pct.style.format("{:.1f}%"), use_container_width=True)

                with tab_daily:
                    if not df_daily.empty:
                        df_daily_f = df_daily[df_daily['Month'].isin(selected_months)]
                        st.markdown("### Daily Collection Heatmap")
                        heatmap = df_daily_f.pivot_table(index='Month', columns='Day', values='Total', aggfunc='sum').fillna(0)
                        st.plotly_chart(px.imshow(heatmap, aspect="auto", color_continuous_scale="Blues", labels=dict(x="Day of Month", y="Month", color="Revenue")), use_container_width=True)
                        
                        st.markdown("### Filing Deadline Spikes")
                        df_daily_f['Deadline'] = df_daily_f['Day'].apply(lambda x: 'Deadline Day' if x in [15, 21, 25, 28, 30, 31] else 'Normal Day')
                        st.plotly_chart(px.box(df_daily_f, x='Deadline', y='Total', color='Deadline', points="all"), use_container_width=True)
                    
                    if not df_offices.empty:
                        st.markdown("### Office Performance")
                        df_off_f = df_offices[df_offices['Month'].isin(selected_months)]
                        office_summary = df_off_f.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
                        st.plotly_chart(px.bar(office_summary, x='Office', y='Total', color='Total', color_continuous_scale='Blues'), use_container_width=True)

                with tab_isic:
                    if not df_isic.empty:
                        df_isic_f = df_isic[df_isic['Month'].isin(selected_months)]
                        
                        if cash_basis == "Current Only":
                            df_isic_f = df_isic_f.copy()
                            df_isic_f['Amount'] = df_isic_f['Current']
                        elif cash_basis == "Arrears Only":
                            df_isic_f = df_isic_f.copy()
                            df_isic_f['Amount'] = df_isic_f['Arrears']
                        else:
                            df_isic_f = df_isic_f.copy()
                            df_isic_f['Amount'] = df_isic_f['Total']
                        
                        df_isic_f = df_isic_f[df_isic_f['Amount'] > 0]
                        
                        st.markdown("### Sector Deep-Dive by Tax Head")
                        selected_isic_tax = st.selectbox("Select Tax Head to Drill Down", df_isic_f['Tax_Type'].unique())
                        drill_df = df_isic_f[df_isic_f['Tax_Type'] == selected_isic_tax]
                        sector_summary = drill_df.groupby('ISIC_Sector')['Amount'].sum().nlargest(15).reset_index()
                        st.plotly_chart(px.bar(sector_summary, x='Amount', y='ISIC_Sector', orientation='h', title=f"Top Sectors for {selected_isic_tax}"), use_container_width=True)
                    else:
                        st.info("Upload the ISIC Variance Reports in the sidebar to unlock sector data.")

                with tab_reports:
                    st.markdown(f"### Variance Notes Generator — {snapshot_month} 2026")
                    st.caption(f"Report will be generated using the **{analysis_mode}** snapshot: **{snapshot_month} 2026**")
                    
                    if st.button(f"Generate {snapshot_month} 2026 Variance Notes"):
                        doc = Document()
                        
                        doc.add_heading('SAINT LUCIA', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_heading('DEPARTMENT OF FINANCE', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_heading('INLAND REVENUE DEPARTMENT', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph('Revenue Collection Volume 14: Monthly Performance Report')
                        doc.add_heading(f'Explanation of Variances for the Month of {snapshot_month} 2026', 2)
                        
                        doc.add_heading('Overview', level=1)
                        doc.add_paragraph(f"For the month of {snapshot_month} 2026, the Inland Revenue Department projected total revenue of ${total_pro_rata:,.2f}. Actual collections amounted to ${total_ytd:,.2f}, resulting in a {'overage' if total_var > 0 else 'shortfall'} of ${abs(total_var):,.2f} or {(total_var/total_pro_rata*100):.1f}% {'above' if total_var > 0 else 'below'} projection.")
                        doc.add_paragraph(f"Note: This report uses the {snapshot_month} 2026 Revenue Summary as the authoritative snapshot for cumulative YTD figures.")
                        
                        doc.add_heading('Variance Analysis by Tax Type', level=1)
                        
                        for _, row in ytd.sort_values('Code').iterrows():
                            code = row['Code']
                            desc = row['Account']
                            proj = row['Pro_Rata']
                            act = row['YTD_Actual']
                            var = row['Variance']
                            pct = (var / proj * 100) if proj > 0 else 0
                            
                            var_str = f"(${abs(var):,.2f})" if var < 0 else f"${var:,.2f}"
                            pct_str = f"({abs(pct):.2f}%)" if pct < 0 else f"{pct:.2f}%"
                            
                            doc.add_heading(f"{code} - {desc}", level=2)
                            doc.add_paragraph(f"Projection: ${proj:,.2f} | Actual: ${act:,.2f} | Variance: {var_str} ({pct_str})")
                            
                            if var > 0:
                                doc.add_paragraph(f"Performance for {desc} was above expectations, yielding a positive variance. This performance suggests favorable filing patterns and robust underlying economic activity in the associated sectors.")
                            else:
                                doc.add_paragraph(f"Collections for {desc} fell short of the pro-rata target. This shortfall may be attributable to delayed filings or lower-than-anticipated economic throughput in the relevant sectors.")
                            
                            if code == "3131 001":
                                doc.add_paragraph("Moratorium Notice: Due to the ongoing moratorium on residential property tax, current payments apply exclusively to commercial properties. Arrears collections apply to both.")
                            elif code == "3142 006":
                                doc.add_paragraph("Stamp Duty Disclaimer: The top-performing sectors listed for Stamp Duty may not be a true indicator of which economic sectors are actually generating the underlying transactions, as ISIC codes are often assigned to legal practitioners paying on behalf of clients.")
                            elif code == "3422 008":
                                doc.add_paragraph("Amnesty Framework: Due to the active Amnesty extending to December 2027, collections are minimal.")
                            
                            if not df_isic.empty:
                                tax_isic = df_isic[(df_isic['Tax_Type'].str.contains(desc.split()[0], case=False, na=False)) & (df_isic['Month'].isin(selected_months))]
                                if not tax_isic.empty:
                                    top_sectors = tax_isic.groupby('ISIC_Sector')['Total'].sum().nlargest(3)
                                    doc.add_heading('Top Performing Sectors', level=3)
                                    for sector, amt in top_sectors.items():
                                        doc.add_paragraph(f"{sector}: ${amt:,.2f}", style='List Bullet')
                                    doc.add_paragraph("These sectors indicate robust activity, potentially driven by seasonal demand or structural shifts. Further observation of these possibilities will have to be made.")
                        
                        doc.add_heading('Conclusion', level=1)
                        doc.add_paragraph(f"The fiscal period ending {snapshot_month} 2026 demonstrates {'resilient' if total_var > 0 else 'challenging'} revenue performance. The primary structural overperformer was {top_perf['Account']}, while {top_short['Account']} requires compliance focus. Heading into the next month, the Department will prioritize {top_short['Account']} recovery and maintain monitoring on high-yield sectors.")
                        
                        bio = _io.BytesIO()
                        doc.save(bio)
                        st.download_button(
                            label=f"📥 Download Variance Notes ({snapshot_month}).docx",
                            data=bio.getvalue(),
                            file_name=f"Variance_Notes_{snapshot_month}_2026.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

elif not main_files:
    st.info("👈 Please upload your Monthly Revenue Reports and ISIC files in the sidebar to begin.")
