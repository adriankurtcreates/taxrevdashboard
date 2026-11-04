import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

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
</style>
""", unsafe_allow_html=True)

FISCAL_MONTHS = ["APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER", "JANUARY", "FEBRUARY", "MARCH"]
MONTH_ALIASES = {
    "APR": "April", "APRIL": "April",
    "MAY": "May",
    "JUN": "June", "JUNE": "June",
    "JUL": "July", "JULY": "July",
    "AUG": "August", "AUGUST": "August",
    "SEP": "September", "SEPT": "September", "SEPTEMBER": "September",
    "OCT": "October", "OCTOBER": "October",
    "NOV": "November", "NOVEMBER": "November",
    "DEC": "December", "DECEMBER": "December",
    "JAN": "January", "JANUARY": "January",
    "FEB": "February", "FEBRUARY": "February",
    "MAR": "March", "MARCH": "March"
}

def extract_month_from_text(text):
    """Try to extract a month name from any string (filename or sheet name)."""
    if not text:
        return None
    text_upper = str(text).upper()
    for alias, canonical in MONTH_ALIASES.items():
        # Look for the month as a whole word
        if re.search(r'\b' + alias + r'\b', text_upper):
            return canonical
    return None

def safe_number(val):
    """Convert to number safely, return 0.0 if not possible."""
    try:
        if pd.isna(val):
            return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# ==============================================================================
# 2. DATA PARSING ENGINE
# ==============================================================================
@st.cache_data
def process_main_reports(file_names_and_bytes):
    """
    Takes a list of (filename, bytes) tuples so caching works.
    Returns: (df_tax, df_office, diagnostics)
    """
    all_tax_data = []
    all_office_data = []
    diagnostics = []
    
    for filename, file_bytes in file_names_and_bytes:
        diag = {"file": filename, "status": "❌ Failed", "rows_parsed": 0, "month": None, "error": None}
        try:
            file_obj = pd.io.common.BytesIO(file_bytes) if not hasattr(file_bytes, 'read') else file_bytes
            xls = pd.ExcelFile(file_obj)
            sheet_names = xls.sheet_names
            
            # --- 1. Find the revenue sheet ---
            revenue_sheet = None
            for s in sheet_names:
                if "actual monthly revenue" in s.lower():
                    revenue_sheet = s
                    break
            
            if not revenue_sheet:
                diag["error"] = f"No 'Actual Monthly Revenue' sheet found. Sheets: {sheet_names}"
                diagnostics.append(diag)
                continue
            
            # --- 2. Determine the month ---
            month = extract_month_from_text(filename)
            if not month:
                month = extract_month_from_text(revenue_sheet)
            if not month:
                diag["error"] = "Could not detect month from filename or sheet name"
                diagnostics.append(diag)
                continue
            diag["month"] = month
            
            # --- 3. Parse the revenue sheet ---
            df = pd.read_excel(file_obj, sheet_name=revenue_sheet, header=None)
            
            rows_parsed = 0
            for i in range(5, min(35, len(df))):
                code = df.iloc[i, 0]
                name = df.iloc[i, 1]
                
                # Skip blanks and total rows
                if pd.isna(code) or pd.isna(name):
                    continue
                if "total" in str(name).lower():
                    continue
                # Tax codes look like "3111   001"
                code_str = str(code).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code_str):
                    continue
                
                approved = safe_number(df.iloc[i, 2])
                monthly_actual = safe_number(df.iloc[i, 9])   # Column J
                ytd_total = safe_number(df.iloc[i, 15])       # Column P
                variance = safe_number(df.iloc[i, 16])        # Column Q
                
                all_tax_data.append({
                    "Month": month,
                    "Code": code_str,
                    "Account": str(name).strip(),
                    "Approved_Annual": approved,
                    "Monthly_Actual": monthly_actual,
                    "YTD_Actual": ytd_total,
                    "YTD_Variance": variance
                })
                rows_parsed += 1
            
            diag["rows_parsed"] = rows_parsed
            
            # --- 4. Parse the office breakdown ---
            daily_sheet = None
            for s in sheet_names:
                if "daily collection" in s.lower():
                    daily_sheet = s
                    break
            
            if daily_sheet:
                df_daily = pd.read_excel(file_obj, sheet_name=daily_sheet, header=None)
                for i in range(30, min(55, len(df_daily))):
                    office_name = df_daily.iloc[i, 0]
                    office_total = df_daily.iloc[i, 20]
                    if pd.isna(office_name) or "sub" in str(office_name).lower():
                        continue
                    office_str = str(office_name).strip()
                    # Only count office rows (they typically have text like "Vieux-Fort Office")
                    if office_str and not office_str.startswith("20") and len(office_str) > 2:
                        all_office_data.append({
                            "Month": month,
                            "Office": office_str,
                            "Total_Collected": safe_number(office_total)
                        })
            
            if rows_parsed > 0:
                diag["status"] = "✅ OK"
                
        except Exception as e:
            diag["error"] = str(e)
        
        diagnostics.append(diag)
    
    df_tax = pd.DataFrame(all_tax_data)
    df_office = pd.DataFrame(all_office_data)
    return df_tax, df_office, pd.DataFrame(diagnostics)


@st.cache_data
def process_isic_reports(file_names_and_bytes):
    """Parse multiple ISIC files. Returns (df_isic, diagnostics)."""
    all_isic_data = []
    diagnostics = []
    
    for filename, file_bytes in file_names_and_bytes:
        diag = {"file": filename, "status": "❌ Failed", "rows_parsed": 0, "month": None, "error": None}
        try:
            file_obj = pd.io.common.BytesIO(file_bytes) if not hasattr(file_bytes, 'read') else file_bytes
            xls = pd.ExcelFile(file_obj)
            
            proj_sheet = None
            for s in xls.sheet_names:
                if "projected" in s.lower() and "actual" in s.lower():
                    proj_sheet = s
                    break
            
            if not proj_sheet:
                diag["error"] = f"No 'Projected v Actual' sheet. Sheets: {xls.sheet_names}"
                diagnostics.append(diag)
                continue
            
            month = extract_month_from_text(filename) or extract_month_from_text(proj_sheet)
            if not month:
                diag["error"] = "Could not detect month"
                diagnostics.append(diag)
                continue
            diag["month"] = month
            
            df = pd.read_excel(file_obj, sheet_name=proj_sheet, header=None)
            
            rows_parsed = 0
            for i in range(5, min(35, len(df))):
                code = df.iloc[i, 1]  # Column B
                name = df.iloc[i, 2]  # Column C
                if pd.isna(code) or pd.isna(name):
                    continue
                code_str = str(code).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code_str):
                    continue
                
                projection = safe_number(df.iloc[i, 3])
                collection = safe_number(df.iloc[i, 4])
                variance = safe_number(df.iloc[i, 7])
                
                all_isic_data.append({
                    "Month": month,
                    "Code": code_str,
                    "Account": str(name).strip(),
                    "Projection": projection,
                    "Collection": collection,
                    "Variance": variance
                })
                rows_parsed += 1
            
            diag["rows_parsed"] = rows_parsed
            if rows_parsed > 0:
                diag["status"] = "✅ OK"
                
        except Exception as e:
            diag["error"] = str(e)
        
        diagnostics.append(diag)
    
    return pd.DataFrame(all_isic_data), pd.DataFrame(diagnostics)

# ==============================================================================
# 3. SIDEBAR
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    st.markdown("### 📂 Data Ingestion")
    st.info("Upload April 2026 → March 2027 for full YTD view.")
    
    main_files = st.file_uploader("Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    
    st.markdown("---")
    selected_months = st.multiselect("Filter by Month", FISCAL_MONTHS, default=FISCAL_MONTHS)
    
    st.caption("Saint Lucia IRD — Strategic Research Terminal")

# ==============================================================================
# 4. MAIN DASHBOARD
# ==============================================================================
st.title("🇱🇨 Revenue Analysis Dashboard")

if main_files:
    # Convert uploads to bytes so caching works
    main_payload = [(f.name, f.getvalue()) for f in main_files]
    df_main, df_daily, diag_main = process_main_reports(main_payload)
    
    # --- DIAGNOSTIC PANEL ---
    with st.expander("🔍 Upload Diagnostics (open if charts are empty)", expanded=(df_main.empty)):
        st.write("**Main Reports:**")
        st.dataframe(diag_main, use_container_width=True)
        if df_main.empty:
            st.error("No rows parsed from any file. Check the 'error' column above to see why.")
            st.write("**File preview (first sheet, first 40 rows):**")
            f = main_files[0]
            preview = pd.read_excel(f, sheet_name=0, header=None).head(40)
            st.dataframe(preview, use_container_width=True)
    
    if not df_main.empty:
        df_filtered = df_main[df_main['Month'].isin(selected_months)]
        
        if df_filtered.empty:
            st.warning(f"Data was parsed (months: {df_main['Month'].unique().tolist()}) but none match your filter. Adjust the sidebar filter.")
        else:
            # Latest month for YTD KPIs (order by fiscal month position)
            month_order = {m: i for i, m in enumerate(FISCAL_MONTHS)}
            available_months = sorted(df_filtered['Month'].unique(), key=lambda m: month_order.get(m, 99))
            latest_month = available_months[-1]
            latest = df_filtered[df_filtered['Month'] == latest_month]
            
            total_ytd_revenue = latest['YTD_Actual'].sum()
            total_annual_target = latest['Approved_Annual'].sum()
            total_ytd_variance = latest['YTD_Variance'].sum()
            months_elapsed = month_order.get(latest_month, 0) + 1
            pro_rata_target = total_annual_target * (months_elapsed / 12)
            
            monthly_rev = df_filtered.groupby('Month')['Monthly_Actual'].sum().reset_index()
            monthly_rev['Order'] = monthly_rev['Month'].map(month_order)
            monthly_rev = monthly_rev.sort_values('Order')
            
            st.caption(f"Showing **{len(available_months)} month(s)** — Latest: **{latest_month}**")
            
            # --- KPI CARDS ---
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">💰 YTD Revenue</div>
                    <div class="kpi-value">${total_ytd_revenue:,.0f}</div>
                    <div class="kpi-delta delta-pos">Through {latest_month}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">🎯 YTD Variance</div>
                    <div class="kpi-value">${abs(total_ytd_variance):,.0f}</div>
                    <div class="kpi-delta {'delta-pos' if total_ytd_variance > 0 else 'delta-neg'}">
                        {'Overage' if total_ytd_variance > 0 else 'Shortage'}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">📈 Pro-Rata Target</div>
                    <div class="kpi-value">${pro_rata_target:,.0f}</div>
                    <div class="kpi-delta">{months_elapsed}/12 Months</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                conv = (total_ytd_revenue / pro_rata_target * 100) if pro_rata_target > 0 else 0
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">✅ Convergence</div>
                    <div class="kpi-value">{conv:.1f}%</div>
                    <div class="kpi-delta {'delta-pos' if conv >= 100 else 'delta-neg'}">
                        {'On Track' if conv >= 100 else 'Behind'}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with col5:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">📊 Accounts</div>
                    <div class="kpi-value">{latest['Code'].nunique()}</div>
                    <div class="kpi-delta">Active Tax Lines</div>
                </div>
                """, unsafe_allow_html=True)
            
            # --- CHARTS ROW 1 ---
            col_chart1, col_chart2, col_chart3 = st.columns([1, 1, 1])
            
            with col_chart1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Status vs Pro-Rata")
                status_df = latest.groupby(['Code', 'Account']).agg(
                    YTD=('YTD_Actual', 'sum'),
                    Target=('Approved_Annual', 'first')
                ).reset_index()
                status_df['Status'] = status_df.apply(
                    lambda x: 'On Track' if x['YTD'] >= (x['Target'] * months_elapsed / 12) else 'Behind',
                    axis=1
                )
                counts = status_df['Status'].value_counts().reset_index()
                counts.columns = ['Status', 'Count']
                fig = px.pie(counts, names='Status', values='Count', hole=0.6,
                             color='Status', color_discrete_map={'On Track': '#00B4D8', 'Behind': '#1A3B47'})
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_chart2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Overall Convergence")
                fig = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=conv,
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={
                        'axis': {'range': [None, 150]},
                        'bar': {'color': "#00B4D8"},
                        'steps': [
                            {'range': [0, 100], 'color': "#E0E0E0"},
                            {'range': [100, 150], 'color': "#F0F0F0"}
                        ],
                    }
                ))
                fig.update_layout(margin=dict(t=30, b=0, l=10, r=10))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_chart3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top 5 YTD Revenue")
                top5 = latest.groupby('Account')['YTD_Actual'].sum().nlargest(5).reset_index()
                fig = px.bar(top5, x='YTD_Actual', y='Account', orientation='h',
                             color='YTD_Actual', color_continuous_scale=['#00B4D8', '#1A3B47'])
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False,
                                  yaxis_title=None, xaxis_title=None)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            # --- CHARTS ROW 2 ---
            col_chart4, col_chart5 = st.columns([2, 1])
            
            with col_chart4:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Monthly Revenue Progression")
                fig = px.line(monthly_rev, x='Month', y='Monthly_Actual', markers=True)
                fig.update_traces(line_color='#00B4D8', marker_size=10)
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0),
                                  xaxis_title=None, yaxis_title="Revenue (XCD)")
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_chart5:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top 10 Shortfalls (YTD)")
                shortfalls = latest.groupby('Account')['YTD_Variance'].sum().nsmallest(10).reset_index()
                if not shortfalls.empty and shortfalls['YTD_Variance'].sum() < 0:
                    fig = px.bar(shortfalls, x='YTD_Variance', y='Account', orientation='h',
                                 color='YTD_Variance', color_continuous_scale='Reds')
                    fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("No negative YTD variances 🎉")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # --- OFFICE BREAKDOWN ---
            if not df_daily.empty:
                st.markdown("### 🏢 Revenue by Collection Office (YTD)")
                df_office_f = df_daily[df_daily['Month'].isin(selected_months)]
                office_summary = df_office_f.groupby('Office')['Total_Collected'].sum().reset_index()
                office_summary = office_summary[office_summary['Total_Collected'] > 0]
                office_summary = office_summary.sort_values('Total_Collected', ascending=False)
                
                if not office_summary.empty:
                    fig = px.bar(office_summary, x='Office', y='Total_Collected',
                                 color='Total_Collected', color_continuous_scale='Blues')
                    st.plotly_chart(fig, use_container_width=True)
            
            # --- DATA TABLE ---
            st.subheader("📋 Comprehensive YTD Inflow Matrix")
            table = latest[['Code', 'Account', 'Approved_Annual', 'YTD_Actual', 'YTD_Variance']].copy()
            st.dataframe(table.style.format({
                "Approved_Annual": "${:,.2f}",
                "YTD_Actual": "${:,.2f}",
                "YTD_Variance": "${:,.2f}"
            }), use_container_width=True, hide_index=True)

# ==============================================================================
# 5. ISIC SECTION
# ==============================================================================
if isic_files:
    isic_payload = [(f.name, f.getvalue()) for f in isic_files]
    df_isic, diag_isic = process_isic_reports(isic_payload)
    
    with st.expander("🔍 ISIC Upload Diagnostics", expanded=(df_isic.empty)):
        st.dataframe(diag_isic, use_container_width=True)
    
    if not df_isic.empty:
        st.markdown("---")
        st.header("🏭 ISIC Sector Deep-Dive")
        
        df_isic_f = df_isic[df_isic['Month'].isin(selected_months)]
        
        if df_isic_f.empty:
            st.warning("ISIC data parsed but no rows match the month filter.")
        else:
            col_isic1, col_isic2 = st.columns([1, 1])
            
            with col_isic1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Sector Revenue Contribution")
                summary = df_isic_f.groupby('Account')['Collection'].sum().reset_index()
                fig = px.treemap(summary, path=['Account'], values='Collection',
                                 color='Collection', color_continuous_scale='Blues')
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_isic2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top Sector Variances")
                var_sum = df_isic_f.groupby('Account')['Variance'].sum().reset_index()
                var_sum = var_sum.sort_values('Variance').head(10)
                fig = px.bar(var_sum, x='Variance', y='Account', orientation='h',
                             color='Variance', color_continuous_scale='RdYlGn')
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

elif not main_files:
    st.info("👈 Upload monthly reports in the sidebar to get started.")
