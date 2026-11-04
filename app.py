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

# ==============================================================================
# 2. DATA PARSING ENGINE (Multi-file support)
# ==============================================================================
def extract_month_from_filename(filename):
    """Extracts the month and year from a filename like 'Monthly Revenue Collections September 2026.xlsx'"""
    for month in FISCAL_MONTHS:
        if month.lower() in filename.lower():
            return month.capitalize()
    return "Unknown"

@st.cache_data
def process_main_reports(files):
    """Loops through multiple monthly revenue files and aggregates them."""
    all_tax_data = []
    all_office_data = []
    
    for file in files:
        month = extract_month_from_filename(file.name)
        
        try:
            # --- 1. Parse Tax Accounts (from 'Actual Monthly Revenue' sheet) ---
            # Find the sheet that contains "Actual Monthly Revenue"
            xls = pd.ExcelFile(file)
            revenue_sheet = next((s for s in xls.sheet_names if "Actual Monthly Revenue" in s), None)
            
            if revenue_sheet:
                df = pd.read_excel(file, sheet_name=revenue_sheet, header=None)
                
                for i in range(5, 32): # Rows 6 to 32 contain tax data
                    code = df.iloc[i, 0]
                    name = df.iloc[i, 1]
                    
                    if pd.notna(code) and pd.notna(name) and "Total" not in str(name):
                        # Columns (0-indexed): 0=Code, 1=Name, 2=Approved, 9=Current Month Actual, 15=YTD Total, 16=Variance
                        # Note: We use column 9 for the specific month's actual, regardless of fiscal month position
                        monthly_actual = pd.to_numeric(df.iloc[i, 9], errors='coerce')
                        ytd_total = pd.to_numeric(df.iloc[i, 15], errors='coerce')
                        variance = pd.to_numeric(df.iloc[i, 16], errors='coerce')
                        approved = pd.to_numeric(df.iloc[i, 2], errors='coerce')
                        
                        all_tax_data.append({
                            "Month": month,
                            "Code": str(code).strip(),
                            "Account": str(name).strip(),
                            "Approved_Annual": approved,
                            "Monthly_Actual": monthly_actual if pd.notna(monthly_actual) else 0,
                            "YTD_Actual": ytd_total if pd.notna(ytd_total) else 0,
                            "YTD_Variance": variance if pd.notna(variance) else 0
                        })
            
            # --- 2. Parse Office Breakdown (from 'Daily Collection' sheet) ---
            daily_sheet = next((s for s in xls.sheet_names if "Daily Collection" in s), None)
            if daily_sheet:
                df_daily = pd.read_excel(file, sheet_name=daily_sheet, header=None)
                for i in range(32, 50): # Office rows
                    office_name = df_daily.iloc[i, 0]
                    office_total = df_daily.iloc[i, 20] # Column U
                    if pd.notna(office_name) and pd.notna(office_total) and "Sub" not in str(office_name):
                        all_office_data.append({
                            "Month": month,
                            "Office": str(office_name).strip(),
                            "Total_Collected": pd.to_numeric(office_total, errors='coerce')
                        })
                        
        except Exception as e:
            st.warning(f"Could not fully parse {file.name}: {e}")
            
    return pd.DataFrame(all_tax_data), pd.DataFrame(all_office_data)

@st.cache_data
def process_isic_reports(files):
    """Loops through multiple ISIC files and aggregates the sector data."""
    all_isic_data = []
    
    for file in files:
        month = extract_month_from_filename(file.name)
        try:
            xls = pd.ExcelFile(file)
            # The ISIC data is scattered, but 'Projected v Actual' gives the top-level summary
            proj_sheet = next((s for s in xls.sheet_names if "Projected v Actual" in s), None)
            
            if proj_sheet:
                df = pd.read_excel(file, sheet_name=proj_sheet, header=None)
                for i in range(6, 30):
                    code = df.iloc[i, 1]
                    name = df.iloc[i, 2]
                    if pd.notna(code) and pd.notna(name):
                        projection = pd.to_numeric(df.iloc[i, 3], errors='coerce')
                        collection = pd.to_numeric(df.iloc[i, 4], errors='coerce')
                        variance = pd.to_numeric(df.iloc[i, 7], errors='coerce')
                        
                        all_isic_data.append({
                            "Month": month,
                            "Code": str(code).strip(),
                            "Account": str(name).strip(),
                            "Projection": projection if pd.notna(projection) else 0,
                            "Collection": collection if pd.notna(collection) else 0,
                            "Variance": variance if pd.notna(variance) else 0
                        })
        except Exception as e:
            st.warning(f"Could not parse ISIC file {file.name}: {e}")
            
    return pd.DataFrame(all_isic_data)

# ==============================================================================
# 3. SIDEBAR (Multi-file Uploaders)
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    st.markdown("### 📂 Data Ingestion")
    st.info("Upload multiple months at once to see YTD trends.")
    
    main_files = st.file_uploader("Upload Monthly Revenue Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    isic_files = st.file_uploader("Upload ISIC Variance Reports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    
    st.markdown("---")
    
    # Dynamic Filter populated after data is loaded
    selected_months = st.multiselect("Filter by Month", FISCAL_MONTHS, default=FISCAL_MONTHS)
    
    st.caption("Saint Lucia IRD — Strategic Research Terminal")

# ==============================================================================
# 4. MAIN DASHBOARD
# ==============================================================================
st.title("🇱🇨 Revenue Analysis Dashboard")

if main_files:
    with st.spinner("Processing monthly reports..."):
        df_main, df_daily = process_main_reports(main_files)
        
    if not df_main.empty:
        # Apply month filter
        df_filtered = df_main[df_main['Month'].isin(selected_months)]
        
        if df_filtered.empty:
            st.warning("No data available for the selected months.")
        else:
            # Calculate KPIs based on the latest month in the filtered dataset
            # Note: YTD_Actual is cumulative, so we take the max YTD value (which will be the latest month)
            latest_month_data = df_filtered[df_filtered['Month'] == df_filtered['Month'].iloc[-1]]
            
            total_ytd_revenue = latest_month_data['YTD_Actual'].sum()
            total_annual_target = latest_month_data['Approved_Annual'].sum()
            total_ytd_variance = latest_month_data['YTD_Variance'].sum()
            monthly_revenue_sum = df_filtered.groupby('Month')['Monthly_Actual'].sum().reset_index()

            # --- KPI CARDS ---
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">💰 YTD Revenue</div>
                    <div class="kpi-value">${total_ytd_revenue:,.0f}</div>
                    <div class="kpi-delta delta-pos">Fiscal Year Total</div>
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
                    <div class="kpi-title">📈 Annual Target</div>
                    <div class="kpi-value">${total_annual_target:,.0f}</div>
                    <div class="kpi-delta">Approved Budget</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col4:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">📅 Months Uploaded</div>
                    <div class="kpi-value">{len(selected_months)}</div>
                    <div class="kpi-delta">Out of 12 Fiscal Months</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col5:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">📊 Active Accounts</div>
                    <div class="kpi-value">{len(df_filtered['Code'].unique())}</div>
                    <div class="kpi-delta">Unique Tax Lines</div>
                </div>
                """, unsafe_allow_html=True)

            # --- CHARTS ROW 1 ---
            col_chart1, col_chart2, col_chart3 = st.columns([1, 1, 1])
            
            with col_chart1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Target Status (YTD)")
                # Aggregate across all months to see how many accounts hit their YTD targets
                df_status = df_filtered.groupby(['Code', 'Account']).agg({'YTD_Actual': 'max', 'Approved_Annual': 'first'}).reset_index()
                df_status['Status'] = df_status.apply(lambda x: 'On Track' if x['YTD_Actual'] >= (x['Approved_Annual'] * (len(selected_months)/12)) else 'Behind', axis=1)
                
                status_counts = df_status['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.6, 
                                    color='Status', color_discrete_map={'On Track': '#00B4D8', 'Behind': '#1A3B47'})
                fig_status.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
                st.plotly_chart(fig_status, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_chart2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Overall Convergence %")
                convergence_pct = (total_ytd_revenue / (total_annual_target * (len(selected_months)/12))) * 100 if total_annual_target > 0 else 0
                
                fig_gauge = go.Figure(go.Indicator(
                    mode = "gauge+number",
                    value = convergence_pct,
                    domain = {'x': [0, 1], 'y': [0, 1]},
                    gauge = {
                        'axis': {'range': [None, 150]},
                        'bar': {'color': "#00B4D8"},
                        'steps': [
                            {'range': [0, 100], 'color': "#E0E0E0"},
                            {'range': [100, 150], 'color': "#F0F0F0"}
                        ],
                    }
                ))
                fig_gauge.update_layout(margin=dict(t=30, b=0, l=10, r=10))
                st.plotly_chart(fig_gauge, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_chart3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top 5 YTD Revenue Generators")
                df_top = df_filtered.groupby('Account')['YTD_Actual'].max().nlargest(5).reset_index()
                fig_bar = px.bar(df_top, x='YTD_Actual', y='Account', orientation='h',
                                 color='YTD_Actual', color_continuous_scale=['#00B4D8', '#1A3B47'])
                fig_bar.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False, yaxis_title=None, xaxis_title=None)
                st.plotly_chart(fig_bar, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # --- CHARTS ROW 2 ---
            col_chart4, col_chart5 = st.columns([2, 1])
            
            with col_chart4:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Revenue Progression (Monthly)")
                
                # Sort months chronologically based on the fiscal year
                month_order = {m: i for i, m in enumerate(FISCAL_MONTHS)}
                monthly_revenue_sum['Order'] = monthly_revenue_sum['Month'].map(month_order)
                monthly_revenue_sum = monthly_revenue_sum.sort_values('Order')
                
                fig_line = px.line(monthly_revenue_sum, x='Month', y='Monthly_Actual', markers=True, 
                                   title="Monthly Actuals")
                fig_line.update_traces(line_color='#00B4D8', marker_size=10)
                fig_line.update_layout(margin=dict(t=30, b=0, l=0, r=0), showlegend=False, xaxis_title=None, yaxis_title="Revenue")
                st.plotly_chart(fig_line, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_chart5:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("YTD Variance by Account")
                df_neg = df_filtered.groupby('Account')['YTD_Variance'].max().nsmallest(10).reset_index()
                if not df_neg.empty:
                    fig_var = px.bar(df_neg, x='YTD_Variance', y='Account', orientation='h', color='YTD_Variance', color_continuous_scale='Reds')
                    st.plotly_chart(fig_var, use_container_width=True)
                else:
                    st.success("No significant negative variances detected.")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- OFFICE BREAKDOWN ---
            if not df_daily.empty:
                st.markdown("### 🏢 Revenue by Collection Office (YTD)")
                df_office_filtered = df_daily[df_daily['Month'].isin(selected_months)]
                office_summary = df_office_filtered.groupby('Office')['Total_Collected'].sum().reset_index().sort_values('Total_Collected', ascending=False)
                
                fig_office = px.bar(office_summary, x='Office', y='Total_Collected', color='Total_Collected', color_continuous_scale='Blues')
                st.plotly_chart(fig_office, use_container_width=True)

            # --- DATA TABLE ---
            st.subheader("📋 Comprehensive YTD Inflow Matrix")
            # Get the latest month's YTD data for the table
            latest_table_data = df_filtered[df_filtered['Month'] == df_filtered['Month'].iloc[-1]][['Code', 'Account', 'Approved_Annual', 'YTD_Actual', 'YTD_Variance']]
            
            st.dataframe(latest_table_data.style.format({
                "Approved_Annual": "${:,.2f}",
                "YTD_Actual": "${:,.2f}",
                "YTD_Variance": "${:,.2f}"
            }), use_container_width=True, hide_index=True)

# ==============================================================================
# 5. ISIC ANALYTICS SECTION
# ==============================================================================
if isic_files:
    with st.spinner("Processing ISIC reports..."):
        df_isic = process_isic_reports(isic_files)
        
    if not df_isic.empty:
        st.markdown("---")
        st.header("🏭 ISIC Sector Deep-Dive")
        
        # Filter ISIC data by selected months
        df_isic_filtered = df_isic[df_isic['Month'].isin(selected_months)]
        
        col_isic1, col_isic2 = st.columns([1, 1])
        
        with col_isic1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Sector Revenue Contribution")
            # Aggregate by Account across all selected months
            df_isic_summary = df_isic_filtered.groupby('Account')['Collection'].sum().reset_index()
            fig_treemap = px.treemap(df_isic_summary, path=['Account'], values='Collection', color='Collection', color_continuous_scale='Blues')
            fig_treemap.update_layout(margin=dict(t=30, b=0, l=0, r=0))
            st.plotly_chart(fig_treemap, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with col_isic2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Top Sector Variances")
            df_isic_var = df_isic_filtered.groupby('Account')['Variance'].sum().reset_index()
            df_isic_var = df_isic_var.sort_values('Variance', ascending=True).head(10)
            
            fig_isic_bar = px.bar(df_isic_var, x='Variance', y='Account', orientation='h', color='Variance', color_continuous_scale='RdYlGn')
            fig_isic_bar.update_layout(margin=dict(t=30, b=0, l=0, r=0))
            st.plotly_chart(fig_isic_bar, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("👈 Please upload Monthly Revenue Reports (and optionally ISIC Reports) in the sidebar to view the dashboard.")
