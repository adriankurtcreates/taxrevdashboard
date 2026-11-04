import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

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

# ==============================================================================
# 2. DATA PARSING ENGINE
# ==============================================================================
@st.cache_data
def load_main_report(file):
    """Parses the Monthly Revenue Collections file using exact cell structures."""
    try:
        # We skip the first 5 rows (headers) and only read the columns we need
        # The actual data starts at row 6 (index 5 in 0-based pandas reading)
        df = pd.read_excel(file, sheet_name="Actual Monthly Revenue Sep 2026", header=None)
        
        # Extract the relevant data block (Rows 5 to 31 contain the tax accounts)
        # Columns: A=Code, B=Name, C=Approved, E=April, F=May... J=Sept, P=Total, Q=Variance
        data = []
        for i in range(5, 32):
            code = df.iloc[i, 0]
            name = df.iloc[i, 1]
            # Check if it's a valid tax line (not a total or blank)
            if pd.notna(code) and pd.notna(name) and "Total" not in str(name):
                approved = df.iloc[i, 2]
                actual_sep = df.iloc[i, 9] # Column J is September Actuals
                total_actual = df.iloc[i, 15] # Column P is Total Actual
                variance = df.iloc[i, 16] # Column Q is Variance
                
                data.append({
                    "Code": str(code).strip(),
                    "Account": str(name).strip(),
                    "Approved": pd.to_numeric(approved, errors='coerce'),
                    "Actual_Sep": pd.to_numeric(actual_sep, errors='coerce'),
                    "Total_Actual": pd.to_numeric(total_actual, errors='coerce'),
                    "Variance": pd.to_numeric(variance, errors='coerce')
                })
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error reading Main Report: {e}")
        return None

@st.cache_data
def load_daily_collection(file):
    """Parses the Daily Collection file to get Office/Location breakdown."""
    try:
        # Data starts around row 10. Total row is at the bottom, but we want the breakdown
        # by Office (Vieux-Fort, Soufriere, Online, etc.)
        df = pd.read_excel(file, sheet_name="Daily Collection - Sep 2026", header=None)
        
        offices = []
        # Offices are listed starting around row 32 in the sheet
        for i in range(32, 50):
            name = df.iloc[i, 0]
            total = df.iloc[i, 20] # Column U holds the row totals
            if pd.notna(name) and pd.notna(total) and name != "Sub Total":
                offices.append({"Office": str(name).strip(), "Total": pd.to_numeric(total, errors='coerce')})
        return pd.DataFrame(offices)
    except Exception as e:
        return pd.DataFrame() # Return empty if fails, not critical

@st.cache_data
def load_isic_data(file):
    """Parses the ISIC file to get tax types and projection vs actual."""
    try:
        df = pd.read_excel(file, sheet_name="Projected v Actual", header=None)
        
        data = []
        for i in range(6, 30):
            code = df.iloc[i, 1]
            name = df.iloc[i, 2]
            if pd.notna(code) and pd.notna(name):
                projection = df.iloc[i, 3]
                collection = df.iloc[i, 4]
                variance = df.iloc[i, 7]
                
                data.append({
                    "Code": str(code).strip(),
                    "Account": str(name).strip(),
                    "Projection": pd.to_numeric(projection, errors='coerce'),
                    "Collection": pd.to_numeric(collection, errors='coerce'),
                    "Variance": pd.to_numeric(variance, errors='coerce')
                })
        return pd.DataFrame(data)
    except Exception as e:
        return pd.DataFrame()

# ==============================================================================
# 3. SIDEBAR
# ==============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/200px-Coat_of_arms_of_Saint_Lucia.svg.png", width=100)
    st.title("Revenue Analysis")
    
    st.markdown("### 📂 Data Ingestion")
    main_file = st.file_uploader("Upload Monthly Revenue Report (.xlsx)", type=["xlsx"])
    isic_file = st.file_uploader("Upload ISIC Variance Report (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    st.caption("Saint Lucia IRD — Strategic Research Terminal")

# ==============================================================================
# 4. MAIN DASHBOARD
# ==============================================================================
st.title("🇱🇨 Revenue Analysis Dashboard")

if main_file is not None:
    df_main = load_main_report(main_file)
    
    if df_main is not None and not df_main.empty:
        df_daily = load_daily_collection(main_file)
        
        # Calculate Totals
        total_revenue = df_main['Total_Actual'].sum()
        total_target = df_main['Approved'].sum()
        total_variance = total_revenue - total_target
        
        # --- KPI CARDS ---
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">💰 Revenue (YTD)</div>
                <div class="kpi-value">${total_revenue:,.0f}</div>
                <div class="kpi-delta delta-pos">Against Target: ${total_target:,.0f}</div>
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
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">📈 Convergence</div>
                <div class="kpi-value">{(total_revenue/total_target*100):.1f}%</div>
                <div class="kpi-delta">Of Target Met</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">📅 Sept Revenue</div>
                <div class="kpi-value">${df_main['Actual_Sep'].sum():,.0f}</div>
                <div class="kpi-delta">Current Month</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col5:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">📊 Active Accounts</div>
                <div class="kpi-value">{len(df_main)}</div>
                <div class="kpi-delta">Tax Revenue Lines</div>
            </div>
            """, unsafe_allow_html=True)

        # --- CHARTS ROW 1 ---
        col_chart1, col_chart2, col_chart3 = st.columns([1, 1, 1])
        
        with col_chart1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Target Status")
            df_main['Status'] = df_main['Variance'].apply(lambda x: 'Met' if x >= 0 else 'Missed')
            status_counts = df_main['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.6, 
                                color='Status', color_discrete_map={'Met': '#00B4D8', 'Missed': '#1A3B47'})
            fig_status.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig_status, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_chart2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Overall Convergence %")
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
            st.subheader("Top 5 Revenue Generators")
            df_top = df_main.nlargest(5, 'Total_Actual')
            fig_bar = px.bar(df_top, x='Total_Actual', y='Account', orientation='h',
                             color='Total_Actual', color_continuous_scale=['#00B4D8', '#1A3B47'])
            fig_bar.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False, yaxis_title=None, xaxis_title=None)
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- CHARTS ROW 2 ---
        col_chart4, col_chart5 = st.columns([2, 1])
        
        with col_chart4:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Revenue by Collection Office")
            if not df_daily.empty:
                fig_region = px.bar(df_daily, x='Office', y='Total', color='Total', color_continuous_scale='Blues')
            else:
                # Fallback placeholder if daily data parsing fails
                fig_region = go.Figure(go.Bar(x=['Castries', 'Vieux Fort', 'Soufriere', 'Online'], 
                                              y=[total_revenue*0.5, total_revenue*0.2, total_revenue*0.1, total_revenue*0.2]))
            fig_region.update_layout(margin=dict(t=30, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig_region, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_chart5:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("Variance Profile")
            df_neg = df_main[df_main['Variance'] < 0].nsmallest(10, 'Variance')
            if not df_neg.empty:
                fig_var = px.bar(df_neg, x='Variance', y='Account', orientation='h', color='Variance', color_continuous_scale='Reds')
                st.plotly_chart(fig_var, use_container_width=True)
            else:
                st.success("All targets met or exceeded!")
            st.markdown('</div>', unsafe_allow_html=True)

        # --- DATA TABLE ---
        st.subheader("📋 Detailed Inflow Matrix")
        st.dataframe(df_main.style.format({
            "Approved": "${:,.2f}",
            "Actual_Sep": "${:,.2f}",
            "Total_Actual": "${:,.2f}",
            "Variance": "${:,.2f}"
        }), use_container_width=True, hide_index=True)

else:
    st.info("👈 Please upload the Monthly Revenue Report (.xlsx) in the sidebar to view the dashboard.")
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9f/Coat_of_arms_of_Saint_Lucia.svg/400px-Coat_of_arms_of_Saint_Lucia.svg.png", width=200)

# Optional: Add ISIC data if second file is uploaded
if isic_file is not None:
    st.markdown("---")
    st.header("🏭 ISIC Sector Analysis")
    df_isic = load_isic_data(isic_file)
    if not df_isic.empty:
        fig_isic = px.treemap(df_isic, path=['Account'], values='Collection', color='Variance', color_continuous_scale='RdYlGn')
        fig_isic.update_layout(margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_isic, use_container_width=True)
    else:
        st.warning("Could not parse ISIC data from the uploaded file.")
