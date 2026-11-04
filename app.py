import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io as _io
from docx import Document
from docx.shared import Pt

# ==============================================================================
# 1. PAGE CONFIG & CUSTOM CSS
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
    .delta-pos { color: #2E8B57; }
    .delta-neg { color: #DC3545; }
    .chart-container { background-color: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .narrative-box { background-color: #f8f9fa; border-left: 5px solid #00B4D8; padding: 15px; margin-bottom: 20px; border-radius: 5px; font-size: 16px; }
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

# ==============================================================================
# 2. DATA PARSING ENGINES
# ==============================================================================
@st.cache_data
def process_main_reports(file_payload):
    all_tax, all_office, diagnostics = [], [], []
    for filename, file_bytes in file_payload:
        diag = {"file": filename, "status": "❌ Failed", "rows_parsed": 0, "month": None, "error": None}
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            
            revenue_sheet = next((s for s in xls.sheet_names if "actual monthly revenue" in s.lower()), None)
            if not revenue_sheet:
                diag["error"] = "No revenue sheet found."
                diagnostics.append(diag); continue

            month = find_month_in_text(filename) or find_month_in_text(revenue_sheet)
            if not month:
                diag["error"] = "Could not detect month."
                diagnostics.append(diag); continue
            diag["month"] = month

            df = pd.read_excel(file_obj, sheet_name=revenue_sheet, header=None)
            
            # Find header row
            header_row = 2
            for r in range(15):
                row_str = " ".join(str(v) for v in df.iloc[r].values if pd.notna(v)).upper()
                if "APPROVED" in row_str and "ACTUAL" in row_str:
                    header_row = r; break

            parsed = 0
            for i in range(header_row + 3, min(header_row + 35, len(df))):
                code = df.iloc[i, 0]
                name = df.iloc[i, 1]
                if pd.isna(code) or pd.isna(name) or "total" in str(name).lower(): continue
                code_str = str(code).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code_str): continue

                approved = safe_number(df.iloc[i, 2])
                
                # Find the monthly actual (first non-zero column among indices 4-14)
                monthly_actual = 0.0
                for col_idx in range(4, 15):
                    val = safe_number(df.iloc[i, col_idx])
                    if val != 0:
                        monthly_actual = val
                        break
                
                # YTD is column P (index 15)
                ytd_actual = safe_number(df.iloc[i, 15])
                if ytd_actual == 0:
                    ytd_actual = sum(safe_number(df.iloc[i, c]) for c in range(4, 15))

                all_tax.append({
                    "Month": month, "Code": code_str, "Account": str(name).strip(),
                    "Approved_Annual": approved, "Monthly_Actual": monthly_actual, "YTD_Actual": ytd_actual
                })
                parsed += 1

            diag["rows_parsed"] = parsed
            if parsed > 0: diag["status"] = "✅ OK"
        except Exception as e:
            diag["error"] = str(e)
        diagnostics.append(diag)
    return pd.DataFrame(all_tax), pd.DataFrame(all_office), pd.DataFrame(diagnostics)

@st.cache_data
def process_isic_deep(file_payload):
    """
    Loops through specific ISIC sheets (PIT, CIT, VAT, HAT, etc.) to extract
    sector-level revenue data. Looks for columns with 'ISIC Description' and 'Current'.
    """
    all_isic = []
    target_sheets = ["PIT", "CIT", "VAT", "HAT", "Travel Tax", "HCSL", "Insurance Premium", "Stamp Duty"]
    
    for filename, file_bytes in file_payload:
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)
            month = find_month_in_text(filename)
            if not month: continue

            for sheet in target_sheets:
                # Case-insensitive find
                actual_sheet = next((s for s in xls.sheet_names if sheet.lower() in s.lower()), None)
                if not actual_sheet: continue

                df = pd.read_excel(file_obj, sheet_name=actual_sheet, header=None)
                
                # Find the header row containing 'ISIC Description'
                header_row = None
                isic_col = None
                current_col = None
                
                for r in range(15): # Search top 15 rows for header
                    row_vals = [str(v).lower() for v in df.iloc[r].values if pd.notna(v)]
                    if any("isic description" in val for val in row_vals):
                        header_row = r
                        for c, val in enumerate(df.iloc[r].values):
                            val_str = str(val).lower()
                            if "isic description" in val_str: isic_col = c
                            if "current" in val_str and current_col is None: current_col = c
                        break
                
                if header_row is None or isic_col is None: continue

                # Parse data rows
                for i in range(header_row + 1, len(df)):
                    isic_name = df.iloc[i, isic_col]
                    # Check if it's a grand total or blank
                    if pd.isna(isic_name) or "grand total" in str(isic_name).lower(): break
                    
                    # Get the value from the 'Current' column or the column immediately after ISIC
                    amount = 0.0
                    if current_col is not None:
                        amount = safe_number(df.iloc[i, current_col])
                    if amount == 0: # Fallback to scan nearby columns
                        for c in range(isic_col + 1, isic_col + 5):
                            if c < len(df.columns):
                                amount += safe_number(df.iloc[i, c])
                    
                    if amount > 0:
                        all_isic.append({
                            "Month": month,
                            "Tax_Type": actual_sheet,
                            "ISIC_Sector": str(isic_name).strip(),
                            "Amount": amount
                        })
        except Exception as e:
            pass # Silently skip if a specific sheet fails, to keep the main app running
            
    return pd.DataFrame(all_isic)

# ==============================================================================
# 3. SIDEBAR
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    st.markdown("### 📂 Data Ingestion")
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    st.markdown("---")
    selected_months = st.multiselect("Filter by Month", FISCAL_MONTHS, default=FISCAL_MONTHS)

# ==============================================================================
# 4. MAIN DASHBOARD
# ==============================================================================
st.title("🇱🇨 Revenue Analysis Dashboard")

if main_files:
    main_payload = [(f.name, f.getvalue()) for f in main_files]
    df_main, _, diag_main = process_main_reports(main_payload)

    if not df_main.empty:
        df_filtered = df_main[df_main['Month'].isin(selected_months)]
        
        if not df_filtered.empty:
            month_order = {m: i for i, m in enumerate(FISCAL_MONTHS)}
            available_months = sorted(df_filtered['Month'].unique(), key=lambda m: month_order.get(m, 99))
            latest_month = available_months[-1]
            months_elapsed = month_order.get(latest_month, 0) + 1
            
            # --- YTD CALCULATION ---
            ytd_data = df_filtered.groupby(['Code', 'Account']).agg(
                YTD_Actual=('Monthly_Actual', 'sum'),
                Approved_Annual=('Approved_Annual', 'first')
            ).reset_index()
            ytd_data['Pro_Rata_Target'] = ytd_data['Approved_Annual'] * (months_elapsed / 12)
            ytd_data['YTD_Variance'] = ytd_data['YTD_Actual'] - ytd_data['Pro_Rata_Target']

            total_ytd = ytd_data['YTD_Actual'].sum()
            total_pro_rata = ytd_data['Pro_Rata_Target'].sum()
            total_variance = total_ytd - total_pro_rata

            # --- NARRATIVE CONTEXT ENGINE ---
            top_performer = ytd_data.nlargest(1, 'YTD_Variance').iloc[0]
            top_shortfall = ytd_data.nsmallest(1, 'YTD_Variance').iloc[0]
            
            st.markdown(f"""
            <div class="narrative-box">
                <b>📊 Executive Summary ({latest_month} YTD):</b><br>
                Total revenue stands at <b>${total_ytd:,.0f}</b> against a pro-rata target of <b>${total_pro_rata:,.0f}</b>. 
                This represents a net {'overage' if total_variance > 0 else 'shortage'} of <b>${abs(total_variance):,.0f}</b>.<br><br>
                The primary driver of positive variance is <b>{top_performer['Account']}</b>, which is overperforming by <b>${top_performer['YTD_Variance']:,.0f}</b>. 
                The most significant drag on revenue is <b>{top_shortfall['Account']}</b>, which is underperforming by <b>${abs(top_shortfall['YTD_Variance']):,.0f}</b>.
            </div>
            """, unsafe_allow_html=True)

            # --- KPI CARDS ---
            col1, col2, col3 = st.columns(3)
            col1.markdown(f'<div class="kpi-card"><div class="kpi-title">YTD Revenue</div><div class="kpi-value">${total_ytd:,.0f}</div></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="kpi-card"><div class="kpi-title">Pro-Rata Target</div><div class="kpi-value">${total_pro_rata:,.0f}</div></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="kpi-card"><div class="kpi-title">Net Variance</div><div class="kpi-value" style="color:{"#2E8B57" if total_variance > 0 else "#DC3545"}">${total_variance:,.0f}</div></div>', unsafe_allow_html=True)

            # --- CHARTS ---
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Top YTD Over-Performers")
                top5 = ytd_data.nlargest(5, 'YTD_Variance')
                st.plotly_chart(px.bar(top5, x='YTD_Variance', y='Account', orientation='h', color_discrete_sequence=['#00B4D8']), use_container_width=True)
            with c2:
                st.subheader("Top YTD Shortfalls")
                short5 = ytd_data.nsmallest(5, 'YTD_Variance')
                st.plotly_chart(px.bar(short5, x='YTD_Variance', y='Account', orientation='h', color_discrete_sequence=['#DC3545']), use_container_width=True)

            # --- REPORT GENERATOR (Word) ---
            st.markdown("### 📝 Generate Reports")
            rep_col1, rep_col2 = st.columns(2)
            
            with rep_col1:
                if st.button("Generate Variance Notes Brief"):
                    doc = Document()
                    doc.add_heading(f'Variance Notes Brief - {latest_month} 2026', 0)
                    doc.add_paragraph(f"Total YTD Revenue: ${total_ytd:,.2f}")
                    doc.add_paragraph(f"Net Variance: ${total_variance:,.2f}")
                    doc.add_heading('Account Level Details', level=1)
                    for _, row in ytd_data.iterrows():
                        doc.add_paragraph(f"{row['Account']}: Actual ${row['YTD_Actual']:,.2f}, Target ${row['Pro_Rata_Target']:,.2f}, Variance ${row['YTD_Variance']:,.2f}")
                    
                    bio = _io.BytesIO()
                    doc.save(bio)
                    st.download_button("📥 Download Word Brief", bio.getvalue(), f"Variance_Notes_{latest_month}.docx")
            
            with rep_col2:
                if st.button("Generate Quarterly Performance Report"):
                    doc = Document()
                    doc.add_heading(f'Quarterly Performance Report', 0)
                    doc.add_paragraph(f"Performance Summary through {latest_month}.")
                    doc.add_heading('Key Drivers', level=1)
                    doc.add_paragraph(f"Top Performer: {top_performer['Account']}")
                    doc.add_paragraph(f"Top Shortfall: {top_shortfall['Account']}")
                    
                    bio = _io.BytesIO()
                    doc.save(bio)
                    st.download_button("📥 Download Quarterly Report", bio.getvalue(), "Quarterly_Performance.docx")

# ==============================================================================
# 5. DEEP ISIC ANALYTICS
# ==============================================================================
if isic_files:
    isic_payload = [(f.name, f.getvalue()) for f in isic_files]
    with st.spinner("Parsing deep ISIC data from tax sheets..."):
        df_isic = process_isic_deep(isic_payload)
        
    if not df_isic.empty:
        st.markdown("---")
        st.header("🏭 Deep ISIC Sector Analysis")
        st.markdown("Revenue breakdown by industry sector, aggregated across all uploaded ISIC reports.")
        
        df_isic_f = df_isic[df_isic['Month'].isin(selected_months)]
        
        if not df_isic_f.empty:
            # Aggregate all sectors across all months
            sector_summary = df_isic_f.groupby('ISIC_Sector')['Amount'].sum().reset_index().sort_values('Amount', ascending=False)
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.subheader("Total Collection by ISIC Sector")
                # Show top 15 sectors
                top_sectors = sector_summary.head(15)
                st.plotly_chart(px.bar(top_sectors, x='Amount', y='ISIC_Sector', orientation='h', color='Amount', color_continuous_scale='Blues'), use_container_width=True)
                
            with c2:
                st.subheader("Revenue Composition by Tax Head")
                tax_summary = df_isic_f.groupby('Tax_Type')['Amount'].sum().reset_index()
                st.plotly_chart(px.pie(tax_summary, names='Tax_Type', values='Amount', hole=0.5), use_container_width=True)
                
            st.markdown("### 📋 Detailed Sector Breakdown Table")
            st.dataframe(sector_summary.style.format({"Amount": "${:,.2f}"}), use_container_width=True)
