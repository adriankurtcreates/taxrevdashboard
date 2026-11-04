import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime

# ==============================================================================
# 1. PAGE CONFIG & CUSTOM CSS (To match the provided UI image)
# ==============================================================================
st.set_page_config(page_title="Saint Lucia Revenue Dashboard", layout="wide", page_icon="🇱🇨")

# Custom CSS for the sidebar and KPI cards
st.markdown("""
<style>
    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #1A3B47; /* Dark teal from the image */
        color: white;
    }
    [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label {
        color: white !important;
    }
    
    /* KPI Card Styling */
    .kpi-card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        border-top: 5px solid #00B4D8; /* Accent color */
        margin-bottom: 20px;
    }
    .kpi-title {
        color: #6C757D;
        font-size: 14px;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 10px;
    }
    .kpi-value {
        color: #1A3B47;
        font-size: 28px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .kpi-delta {
        font-size: 14px;
        font-weight: 500;
    }
    .delta-pos { color: #2E8B57; }
    .delta-neg { color: #DC3545; }
    
    /* Chart Container */
    .chart-container {
        background-color: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. DATA PARSING ENGINE
# ==============================================================================
# This function reads your specific Excel file layout and cleans it up.
@st.cache_data
def load_and_clean_data(file):
    if file is None:
        return None, None, None, None
        
    # Load the primary sheets based on your file structure
    try:
        # Sheet 1: Main Revenue Summary
        df_main = pd.read_excel(file, sheet_name="Actual Monthly Revenue Sep 2026", skiprows=4)
        
        # Sheet 2: Daily Collection (useful for trend lines)
        df_daily = pd.read_excel(file, sheet_name="Daily Collection - Sep 2026", skiprows=5)
        
        # Sheet 3: ISIC Sector Breakdown (useful for tree maps)
        df_isic = pd.read_excel(file, sheet_name="Projected v Actual", skiprows=4)
        
        # Sheet 4: Collection & Projection
        df_proj = pd.read_excel(file, sheet_name="Collection & Projection Report", skiprows=5)
        
        return df_main, df_daily, df_isic, df_proj
    except Exception as e:
        st.error(f"Error reading file: {e}. Please ensure the format matches the standard monthly report.")
        return None, None, None, None

# ==============================================================================
# 3. SIDEBAR (Navigation & Filters)
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    # File Uploader
    uploaded_file = st.file_uploader("Upload Monthly Revenue Report (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    st.subheader("Filters")
    
    # These will be populated after the file is loaded
    available_offices = ["All"]
    available_sectors = ["All"]
    
    selected_office = st.selectbox("Collection Office", available_offices)
    selected_sector = st.selectbox("ISIC Sector", available_sectors)
    
    st.markdown("---")
    st.caption("Saint Lucia IRD — Strategic Research Terminal")

# ==============================================================================
# 4. MAIN DASHBOARD LOGIC
# ==============================================================================
st.title("🇱🇨 Revenue Analysis Dashboard")
st.markdown("Power BI Style Overview — Monthly Performance")

if uploaded_file is not None:
    df_main, df_daily, df_isic, df_proj = load_and_clean_data(uploaded_file)
    
    if df_main is not None:
        # --- DATA PROCESSING ---
        # Clean up the main dataframe (drop empty rows, rename columns)
        df_main = df_main.dropna(subset=['ACCOUNT DESCRIPTION'])
        # Filter out 'Total' rows to avoid double counting
        df_main = df_main[~df_main['ACCOUNT DESCRIPTION'].str.contains('Total', na=False)]
        
        # Convert numeric columns
        numeric_cols = ['APPROVED REVENUE', 'ACTUAL REVENUE', 'TOTAL', 'VARIANCE']
        for col in numeric_cols:
            if col in df_main.columns:
                df_main[col] = pd.to_numeric(df_main[col], errors='coerce').fillna(0)
        
        # Calculate Totals
        total_revenue = df_main['ACTUAL REVENUE'].sum()
        total_target = df_main['APPROVED REVENUE'].sum()
        total_variance = total_revenue - total_target
        
        # --- 4.1 KPI CARDS (Top Row) ---
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">💰 Revenue</div>
                <div class="kpi-value">${total_revenue:,.0f}</div>
                <div class="kpi-delta delta-pos">↑ Target: ${total_target:,.0f}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">🎯 Variance</div>
                <div class="kpi-value">${abs(total_variance):,.0f}</div>
                <div class="kpi-delta {'delta-pos' if total_variance > 0 else 'delta-neg'}">
                    {'Overage' if total_variance > 0 else 'Shortage'}
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        with col3:
            # Example: Profit Margin placeholder (needs Cost data)
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">📈 Performance</div>
                <div class="kpi-value">{(total_revenue/total_target*100):.1f}%</div>
                <div class="kpi-delta">Convergence Rate</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col4:
            # Placeholder for collection rate or similar
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">🏢 Offices</div>
                <div class="kpi-value">5</div>
                <div class="kpi-delta">Active Locations</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col5:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">📊 Accounts</div>
                <div class="kpi-value">{len(df_main)}</div>
                <div class="kpi-delta">Active Tax Lines</div>
            </div>
            """, unsafe_allow_html=True)

        # --- 4.2 CHARTS ROW 1 ---
        col_chart1, col_chart2, col_chart3 = st.columns([1, 1, 1])
        
        with col_chart1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Total Count by Status")
            # Simulate Status based on Variance
            df_main['Status'] = df_main['VARIANCE'].apply(lambda x: 'Met' if x >= 0 else 'Missed')
            status_counts = df_main['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.6, 
                                color='Status', color_discrete_map={'Met': '#00B4D8', 'Missed': '#1A3B47'})
            fig_status.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig_status, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_chart2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Profit Margin %")
            # Gauge chart placeholder
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = (total_revenue/total_target*100),
                domain = {'x': [0, 1], 'y': [0, 1]},
                gauge = {
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "#00B4D8"},
                    'steps': [
                        {'range': [0, 50], 'color': "#E0E0E0"},
                        {'range': [50, 100], 'color': "#F0F0F0"}
                    ],
                }
            ))
            fig_gauge.update_layout(margin=dict(t=30, b=0, l=10, r=10))
            st.plotly_chart(fig_gauge, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_chart3:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Performance by Tax Type")
            # Top 5 taxes by actual revenue
            df_top = df_main.nlargest(5, 'ACTUAL REVENUE')
            fig_bar = px.bar(df_top, x='ACTUAL REVENUE', y='ACCOUNT DESCRIPTION', orientation='h',
                             color='ACTUAL REVENUE', color_continuous_scale=['#00B4D8', '#1A3B47'])
            fig_bar.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False, yaxis_title=None, xaxis_title=None)
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- 4.3 CHARTS ROW 2 ---
        col_chart4, col_chart5 = st.columns([2, 1])
        
        with col_chart4:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Target and Revenue by Region/Office")
            # Assuming we can extract office data from the Daily Collection sheet
            # For now, let's create a simulated breakdown from the main data
            # In a real scenario, you'd group by 'PAYMENT_LOC_DESC' from the daily sheet
            
            # Simulating office data
            office_data = pd.DataFrame({
                'Office': ['Castries', 'Vieux Fort', 'Soufriere', 'Gros Islet', 'Online'],
                'Target': [total_target*0.4, total_target*0.2, total_target*0.1, total_target*0.15, total_target*0.15],
                'Revenue': [total_revenue*0.45, total_revenue*0.18, total_revenue*0.08, total_revenue*0.14, total_revenue*0.15]
            })
            
            fig_region = go.Figure(data=[
                go.Bar(name='Target', x=office_data['Office'], y=office_data['Target'], marker_color='#1A3B47'),
                go.Bar(name='Revenue', x=office_data['Office'], y=office_data['Revenue'], marker_color='#00B4D8')
            ])
            fig_region.update_layout(barmode='group', margin=dict(t=30, b=0, l=0, r=0))
            st.plotly_chart(fig_region, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_chart5:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Variance Heatmap (Top 10)")
            # Heatmap of variance
            df_heat = df_main.nlargest(10, 'ACTUAL REVENUE')[['ACCOUNT DESCRIPTION', 'VARIANCE']]
            df_heat['Color'] = df_heat['VARIANCE'].apply(lambda x: '#2E8B57' if x > 0 else '#DC3545')
            
            fig_heat = px.bar(df_heat, x='VARIANCE', y='ACCOUNT DESCRIPTION', orientation='h',
                              color='VARIANCE', color_continuous_scale=['#DC3545', '#2E8B57'])
            fig_heat.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig_heat, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- 4.4 DATA TABLE ---
        st.markdown("### Detailed Inflow Summary Matrix")
        st.dataframe(df_main[['ACCOUNT DESCRIPTION', 'APPROVED REVENUE', 'ACTUAL REVENUE', 'VARIANCE']].style.format({
            'APPROVED REVENUE': "${:,.2f}",
            'ACTUAL REVENUE': "${:,.2f}",
            'VARIANCE': "${:,.2f}"
        }), use_container_width=True)

else:
    st.info("👈 Please upload the Monthly Revenue Report (.xlsx) in the sidebar to view the dashboard.")
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=200)
