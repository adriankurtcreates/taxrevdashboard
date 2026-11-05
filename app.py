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
    .chart-container { background-color: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .narrative-box { background-color: #f8f9fa; border-left: 5px solid #00B4D8; padding: 15px; margin-bottom: 20px; border-radius: 5px; font-size: 16px; }
    .snapshot-badge { display: inline-block; background-color: #00B4D8; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .snapshot-badge-warn { display: inline-block; background-color: #F59E0B; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 8px; }
    .debug-box { background-color: #fff3cd; border: 1px solid #ffc107; padding: 10px; border-radius: 5px; font-size: 13px; }
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
    "3151 009": "Passenger Facility Fee", "3422 008": "Collection Fees - Towns & Villages"
}
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

def month_order(m):
    return FISCAL_MONTHS.index(m) if m in FISCAL_MONTHS else 99

# ==============================================================================
# 3. PARSERS
# ==============================================================================
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
            for i in range(5, min(35, len(df))):
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
            diagnostics.append({"file": filename, "month": month, "rows": len([t for t in all_tax if t["Month"] == month])})
        except Exception as e:
            diagnostics.append({"file": filename, "month": None, "rows": 0, "error": str(e)})
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
# 4. SNAPSHOT PRIORITY ENGINE
# ==============================================================================
def select_authoritative_snapshot(df_tax, selected_months, analysis_mode):
    filtered = df_tax[df_tax['Month'].isin(selected_months)]
    if filtered.empty:
        return None, None
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
# 5. SIDEBAR (with Select All / Clear All)
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    
    if main_files:
        st.markdown("### 🎯 Analysis Mode")
        analysis_mode = st.radio(
            "Which snapshot should drive YTD?",
            ["YTD (Latest Summary)", "Quarter-End Snapshot", "Month Snapshot"]
        )
        st.markdown("---")
        st.markdown("### 🎛️ Filters")
        
        # --- MONTH SLICER with Select All / Clear All ---
        st.markdown("**Month**")
        mc1, mc2 = st.columns(2)
        if mc1.button("All", key="m_all"): st.session_state["month_sel"] = FISCAL_MONTHS
        if mc2.button("None", key="m_none"): st.session_state["month_sel"] = []
        if "month_sel" not in st.session_state: st.session_state["month_sel"] = FISCAL_MONTHS
        selected_months = st.multiselect("Months", FISCAL_MONTHS, key="month_sel", label_visibility="collapsed")
        
        # --- QUARTER SLICER (dynamic) ---
        all_quarters = list(set([get_quarter(m) for m in FISCAL_MONTHS]))
        all_quarters = [q for q in all_quarters if q != "Unknown"]
        st.markdown("**Quarter**")
        qc1, qc2 = st.columns(2)
        if qc1.button("All", key="q_all"): st.session_state["qtr_sel"] = all_quarters
        if qc2.button("None", key="q_none"): st.session_state["qtr_sel"] = []
        if "qtr_sel" not in st.session_state: st.session_state["qtr_sel"] = all_quarters
        selected_quarters = st.multiselect("Quarters", all_quarters, key="qtr_sel", label_visibility="collapsed")
        
        # --- TAX HEAD SLICER (dynamic from actual data) ---
        st.markdown("**Tax Head**")
        tax_options = list(TAX_HEADS.values())
        tc1, tc2 = st.columns(2)
        if tc1.button("All", key="t_all"): st.session_state["tax_sel"] = tax_options
        if tc2.button("None", key="t_none"): st.session_state["tax_sel"] = []
        if "tax_sel" not in st.session_state: st.session_state["tax_sel"] = tax_options
        selected_taxes = st.multiselect("Tax Heads", tax_options, key="tax_sel", label_visibility="collapsed")
        
        # --- CASH BASIS ---
        st.markdown("**Cash Basis**")
        cash_basis = st.radio("Focus", ["All", "Current Only", "Arrears Only"], horizontal=True, label_visibility="collapsed")

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

    # --- DIAGNOSTIC PANEL (always visible) ---
    with st.expander("🔍 Data Diagnostics", expanded=df_tax.empty):
        st.dataframe(diag_main, use_container_width=True)
        if not df_tax.empty:
            col_a, col_b = st.columns(2)
            with col_a:
                st.write("**Months in data:**", sorted(df_tax['Month'].unique().tolist()))
                st.write("**Quarters in data:**", sorted(df_tax['Quarter'].unique().tolist()))
            with col_b:
                st.write("**Codes in data:**", sorted(df_tax['Code'].unique().tolist()))
                st.write("**Total rows parsed:**", len(df_tax))
            st.write("**Selected Months:**", selected_months)
            st.write("**Selected Quarters:**", selected_quarters)
            st.write("**Selected Taxes (first 5):**", selected_taxes[:5])

    if not df_tax.empty:
        # --- STRICT FILTER CASCADE WITH EXPLICIT CHECK ---
        if not selected_months:
            st.warning("⚠️ No months selected. Click 'All' in the Month filter to restore.")
        elif not selected_quarters:
            st.warning("⚠️ No quarters selected. Click 'All' in the Quarter filter to restore.")
        elif not selected_taxes:
            st.warning("⚠️ No tax heads selected. Click 'All' in the Tax Head filter to restore.")
        else:
            # Apply filters
            df_f = df_tax[df_tax['Month'].isin(selected_months)].copy()
            df_f = df_f[df_f['Quarter'].isin(selected_quarters)].copy()
            selected_codes = [code for code, name in TAX_HEADS.items() if name in selected_taxes]
            df_f = df_f[df_f['Code'].isin(selected_codes)].copy()
            
            if df_f.empty:
                st.error("❌ Slicers are selecting valid options but no data rows remain. "
                         "Check the Data Diagnostics panel to see if your filters are too narrow.")
                st.write("**Debugging:**")
                st.write(f"- Rows after month filter: {len(df_tax[df_tax['Month'].isin(selected_months)])}")
                st.write(f"- Rows after quarter filter: {len(df_tax[df_tax['Month'].isin(selected_months) & df_tax['Quarter'].isin(selected_quarters)])}")
                st.write(f"- Rows after tax filter: {len(df_f)}")
                st.write(f"- Selected codes: {selected_codes}")
            else:
                # Snapshot selection
                snapshot_df, snapshot_month = select_authoritative_snapshot(df_f, selected_months, analysis_mode)
                
                if snapshot_df is None or snapshot_df.empty:
                    st.error(f"No {analysis_mode} snapshot available from your uploaded files. "
                             f"Try uploading a {', '.join(QUARTER_END_MONTHS.keys())} file, or switch analysis mode.")
                else:
                    months_elapsed = month_order(snapshot_month) + 1
                    ytd = snapshot_df.groupby(['Code', 'Account']).agg(
                        YTD_Actual=('YTD_Actual', 'first'),
                        Approved_Annual=('Approved_Annual', 'first')
                    ).reset_index()
                    ytd['Pro_Rata'] = ytd['Approved_Annual'] * (months_elapsed / 12)
                    ytd['Variance'] = ytd['YTD_Actual'] - ytd['Pro_Rata']
                    
                    total_ytd = ytd['YTD_Actual'].sum()
                    total_pro_rata = ytd['Pro_Rata'].sum()
                    total_var = total_ytd - total_pro_rata
                    
                    top_perf = ytd.nlargest(1, 'Variance').iloc[0] if not ytd.empty else None
                    top_short = ytd.nsmallest(1, 'Variance').iloc[0] if not ytd.empty else None
                    
                    is_quarter_end = snapshot_month in QUARTER_END_MONTHS
                    badge_class = "snapshot-badge" if is_quarter_end else "snapshot-badge-warn"
                    badge_text = f"✓ Quarter-End: {snapshot_month}" if is_quarter_end else f"Latest: {snapshot_month}"
                    
                    st.markdown(f"""
                    <div class="narrative-box">
                        <b>📊 Executive Summary</b>
                        <span class="{badge_class}">{badge_text}</span><br><br>
                        <b>Data Source:</b> <i>{snapshot_month} 2026 Revenue Summary</i> (Mode: {analysis_mode})<br>
                        <b>YTD Revenue:</b> ${total_ytd:,.0f} vs <b>Pro-Rata Target:</b> ${total_pro_rata:,.0f}<br>
                        <b>Net {'Overage' if total_var > 0 else 'Shortage'}:</b> ${abs(total_var):,.0f} ({(total_var/total_pro_rata*100):.1f}%)<br><br>
                        <b>Top Driver:</b> {top_perf['Account']} (+${top_perf['Variance']:,.0f})<br>
                        <b>Main Drag:</b> {top_short['Account']} (${top_short['Variance']:,.0f})
                    </div>
                    """, unsafe_allow_html=True)
                    
                    c1, c2, c3 = st.columns(3)
                    c1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Revenue</div><div class="kpi-value">${total_ytd:,.0f}</div><div class="kpi-delta">Source: {snapshot_month}</div></div>', unsafe_allow_html=True)
                    c2.markdown(f'<div class="kpi-card"><div class="kpi-title">Pro-Rata Target</div><div class="kpi-value">${total_pro_rata:,.0f}</div><div class="kpi-delta">{months_elapsed}/12 Months</div></div>', unsafe_allow_html=True)
                    c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_var > 0 else "#DC3545"}">${total_var:,.0f}</div></div>', unsafe_allow_html=True)
                    
                    # --- TABS ---
                    tab1, tab2, tab3, tab4 = st.tabs(["📈 Macro Trends", "📅 Daily & Office", "🏭 Sector (ISIC)", "📝 Reports"])
                    
                    with tab1:
                        st.markdown("### Tax Head Composition Over Time")
                        mc = df_f.groupby(['Month', 'Account'])['Monthly_Actual'].sum().reset_index()
                        mc['Order'] = mc['Month'].map(month_order)
                        mc = mc.sort_values('Order')
                        if not mc.empty:
                            st.plotly_chart(px.area(mc, x='Month', y='Monthly_Actual', color='Account', groupnorm='fraction'), use_container_width=True)
                    
                    with tab2:
                        if not df_daily.empty:
                            df_d_f = df_daily[df_daily['Month'].isin(selected_months)]
                            if not df_d_f.empty:
                                heatmap = df_d_f.pivot_table(index='Month', columns='Day', values='Total', aggfunc='sum').fillna(0)
                                st.plotly_chart(px.imshow(heatmap, aspect="auto", color_continuous_scale="Blues"), use_container_width=True)
                        if not df_offices.empty:
                            dfo = df_offices[df_offices['Month'].isin(selected_months)]
                            if not dfo.empty:
                                s = dfo.groupby('Office')['Total'].sum().reset_index().sort_values('Total', ascending=False)
                                st.plotly_chart(px.bar(s, x='Office', y='Total', color='Total'), use_container_width=True)
                    
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
                        st.markdown(f"### Variance Notes — {snapshot_month} 2026")
                        if st.button(f"Generate Report for {snapshot_month}"):
                            doc = Document()
                            doc.add_heading('SAINT LUCIA - DEPARTMENT OF FINANCE', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                            doc.add_heading(f'Variance Notes — {snapshot_month} 2026', 1)
                            doc.add_paragraph(f"Total YTD: ${total_ytd:,.2f} | Variance: ${total_var:,.2f}")
                            for _, r in ytd.sort_values('Code').iterrows():
                                doc.add_heading(f"{r['Code']} - {r['Account']}", level=2)
                                doc.add_paragraph(f"Actual: ${r['YTD_Actual']:,.2f} | Pro-Rata: ${r['Pro_Rata']:,.2f} | Variance: ${r['Variance']:,.2f}")
                            bio = _io.BytesIO()
                            doc.save(bio)
                            st.download_button("📥 Download", bio.getvalue(), f"Variance_Notes_{snapshot_month}.docx")
