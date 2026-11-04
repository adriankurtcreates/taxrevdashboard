import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io as _io

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

FISCAL_MONTHS = ["April", "May", "June", "July", "August", "September",
                 "October", "November", "December", "January", "February", "March"]

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
    "MAR": "March", "MARCH": "March",
}

def find_month_in_text(text):
    if not text:
        return None
    text_upper = str(text).upper()
    for alias in sorted(MONTH_ALIASES.keys(), key=len, reverse=True):
        if re.search(r'\b' + alias + r'\b', text_upper):
            return MONTH_ALIASES[alias]
    return None

def safe_number(val):
    try:
        if pd.isna(val):
            return 0.0
        s = str(val).replace(",", "").replace("$", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return 0.0

# ==============================================================================
# 2. DATA PARSING ENGINE — header-driven, not position-driven
# ==============================================================================
@st.cache_data
def process_main_reports(file_payload):
    all_tax, all_office, diagnostics = [], [], []
    for filename, file_bytes in file_payload:
        diag = {
            "file": filename, "status": "❌ Failed", "rows_parsed": 0,
            "month": None, "detection_source": None,
            "monthly_sum": 0.0, "ytd_sum": 0.0, "error": None,
        }
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)

            revenue_sheet = next(
                (s for s in xls.sheet_names if "actual monthly revenue" in s.lower()), None
            )
            if not revenue_sheet:
                revenue_sheet = next(
                    (s for s in xls.sheet_names
                     if "revenue" in s.lower() and "monthly" in s.lower()), None
                )
            if not revenue_sheet:
                diag["error"] = f"No revenue sheet. Sheets: {xls.sheet_names}"
                diagnostics.append(diag)
                continue

            # --- Month detection ---
            month = find_month_in_text(filename)
            source = "filename"
            if not month:
                month = find_month_in_text(revenue_sheet)
                source = f"sheet name ({revenue_sheet})"
            if not month:
                for s in xls.sheet_names:
                    month = find_month_in_text(s)
                    if month:
                        source = f"other sheet ({s})"
                        break
            if not month:
                df_probe = pd.read_excel(file_obj, sheet_name=revenue_sheet, header=None, nrows=10)
                for val in df_probe.values.flatten():
                    month = find_month_in_text(val)
                    if month:
                        source = f"cell ('{val}')"
                        break
            if not month:
                diag["error"] = "Could not detect month"
                diagnostics.append(diag)
                continue
            diag["month"] = month
            diag["detection_source"] = source

            # --- Read the sheet WITHOUT header, keep raw values ---
            df = pd.read_excel(file_obj, sheet_name=revenue_sheet, header=None)

            # --- Locate header row by finding the row containing 'ACTUAL' + 'TOTAL' ---
            header_row = None
            for r in range(0, min(15, len(df))):
                row_str = " ".join(str(v) for v in df.iloc[r].values if pd.notna(v)).upper()
                if "APPROVED" in row_str and "ACTUAL" in row_str and "TOTAL" in row_str:
                    header_row = r
                    break

            # Fallback: known layout — headers in row 2 (index 2), months in row 3
            if header_row is None:
                header_row = 2

            # --- Build a column index map from the header row ---
            # Row at header_row contains: APPROVED, REVISED, ACTUAL, ACTUAL, ..., TOTAL, VARIANCE
            # We want: first ACTUAL (Sept), and the TOTAL column
            header_values = [str(v).strip().upper() if pd.notna(v) else "" for v in df.iloc[header_row].values]

            # Find first "ACTUAL" column AFTER "REVISED" (skips the 'ACTUAL REVENUE' label)
            approved_col = None
            actual_col = None
            total_col = None
            variance_col = None

            for idx, val in enumerate(header_values):
                if val == "APPROVED" and approved_col is None:
                    approved_col = idx
                if val == "ACTUAL" and actual_col is None and idx > (approved_col or 0) + 2:
                    # The 'ACTUAL' at the top header usually spans multiple months
                    # We want the FIRST one after REVISED; but the actual month columns
                    # are labeled 'ACTUAL REVENUE' in row 3, so use that.
                    pass
                if val == "TOTAL" and total_col is None:
                    total_col = idx
                if val == "VARIANCE" and variance_col is None:
                    variance_col = idx

            # Better: find ACTUAL columns from row header_row + 1
            sub_header_values = [str(v).strip().upper() if pd.notna(v) else "" for v in df.iloc[header_row + 1].values]
            actual_indices = [i for i, v in enumerate(sub_header_values) if v == "ACTUAL"]
            # The September actual is the LAST actual (since the sheet is April → March and only April-September have data)
            # But we want the one where data exists — so we'll check each one and pick the first with a value.

            # --- Parse data rows ---
            parsed = 0
            monthly_sum = 0.0
            ytd_sum = 0.0
            for i in range(header_row + 3, min(header_row + 35, len(df))):
                code = df.iloc[i, 0]
                name = df.iloc[i, 1]
                if pd.isna(code) or pd.isna(name):
                    continue
                if "total" in str(name).lower():
                    continue
                code_str = str(code).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code_str):
                    continue

                # Approved = first numeric column with APPROVED header (fallback col 2)
                approved = safe_number(df.iloc[i, approved_col if approved_col is not None else 2])

                # Monthly actual: scan ALL 'ACTUAL' sub-columns for the FIRST one with a non-zero value
                monthly_actual = 0.0
                for ac in actual_indices:
                    val = safe_number(df.iloc[i, ac])
                    if val != 0:
                        monthly_actual = val
                        break  # take the first non-zero (usually September in a September file)

                # YTD: prefer the TOTAL column, fall back to summing monthlies later
                ytd_actual = safe_number(df.iloc[i, total_col]) if total_col is not None else 0.0
                ytd_variance = safe_number(df.iloc[i, variance_col]) if variance_col is not None else 0.0

                # If total_col returned 0 for a real row, reconstruct YTD from monthlies
                if ytd_actual == 0:
                    monthly_sum_this_row = sum(
                        safe_number(df.iloc[i, ac]) for ac in actual_indices
                    )
                    ytd_actual = monthly_sum_this_row
                if ytd_variance == 0 and approved > 0:
                    ytd_variance = ytd_actual - approved

                all_tax.append({
                    "Month": month,
                    "Code": code_str,
                    "Account": str(name).strip(),
                    "Approved_Annual": approved,
                    "Monthly_Actual": monthly_actual,
                    "YTD_Actual": ytd_actual,
                    "YTD_Variance": ytd_variance,
                })
                monthly_sum += monthly_actual
                ytd_sum += ytd_actual
                parsed += 1

            diag["rows_parsed"] = parsed
            diag["monthly_sum"] = monthly_sum
            diag["ytd_sum"] = ytd_sum

            # --- Daily collection sheet ---
            daily_sheet = next(
                (s for s in xls.sheet_names if "daily collection" in s.lower()), None
            )
            if daily_sheet:
                df_daily = pd.read_excel(file_obj, sheet_name=daily_sheet, header=None)
                for i in range(30, min(55, len(df_daily))):
                    office = df_daily.iloc[i, 0]
                    total = df_daily.iloc[i, 20]
                    if pd.isna(office) or "sub" in str(office).lower():
                        continue
                    office_str = str(office).strip()
                    if office_str and not office_str.startswith("20") and len(office_str) > 2:
                        all_office.append({
                            "Month": month,
                            "Office": office_str,
                            "Total_Collected": safe_number(total),
                        })

            if parsed > 0 and monthly_sum > 0:
                diag["status"] = "✅ OK"
            elif parsed > 0:
                diag["error"] = "Rows parsed but all monthly values are 0 — check column indices"

        except Exception as e:
            diag["error"] = str(e)
        diagnostics.append(diag)

    return pd.DataFrame(all_tax), pd.DataFrame(all_office), pd.DataFrame(diagnostics)


@st.cache_data
def process_isic_reports(file_payload):
    all_isic, diagnostics = [], []
    for filename, file_bytes in file_payload:
        diag = {"file": filename, "status": "❌ Failed", "rows_parsed": 0,
                "month": None, "detection_source": None, "error": None}
        try:
            file_obj = _io.BytesIO(file_bytes)
            xls = pd.ExcelFile(file_obj)

            proj_sheet = next(
                (s for s in xls.sheet_names
                 if "projected" in s.lower() and "actual" in s.lower()), None
            )
            if not proj_sheet:
                diag["error"] = f"No 'Projected v Actual' sheet. Sheets: {xls.sheet_names}"
                diagnostics.append(diag)
                continue

            month = find_month_in_text(filename)
            source = "filename"
            if not month:
                month = find_month_in_text(proj_sheet)
                source = f"sheet name ({proj_sheet})"
            if not month:
                for s in xls.sheet_names:
                    month = find_month_in_text(s)
                    if month:
                        source = f"other sheet ({s})"
                        break
            if not month:
                diag["error"] = "Could not detect month"
                diagnostics.append(diag)
                continue
            diag["month"] = month
            diag["detection_source"] = source

            df = pd.read_excel(file_obj, sheet_name=proj_sheet, header=None)
            parsed = 0
            for i in range(5, min(35, len(df))):
                code = df.iloc[i, 1]
                name = df.iloc[i, 2]
                if pd.isna(code) or pd.isna(name):
                    continue
                code_str = str(code).strip()
                if not re.match(r'^\d{4}\s*\d{3}$', code_str):
                    continue

                all_isic.append({
                    "Month": month,
                    "Code": code_str,
                    "Account": str(name).strip(),
                    "Projection": safe_number(df.iloc[i, 3]),
                    "Collection": safe_number(df.iloc[i, 4]),
                    "Variance": safe_number(df.iloc[i, 7]),
                })
                parsed += 1
            diag["rows_parsed"] = parsed
            if parsed > 0:
                diag["status"] = "✅ OK"
        except Exception as e:
            diag["error"] = str(e)
        diagnostics.append(diag)
    return pd.DataFrame(all_isic), pd.DataFrame(diagnostics)

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
    main_payload = [(f.name, f.getvalue()) for f in main_files]
    df_main, df_daily, diag_main = process_main_reports(main_payload)

    with st.expander("🔍 Main Report Diagnostics", expanded=df_main.empty):
        st.dataframe(diag_main, use_container_width=True)

        # Show a preview of the first file so you can see what the parser sees
        if not df_main.empty:
            st.write("**Sample parsed rows (first 5):**")
            st.dataframe(df_main.head(), use_container_width=True)

    if not df_main.empty:
        df_filtered = df_main[df_main['Month'].isin(selected_months)]

        if df_filtered.empty:
            st.warning(f"Parsed months: {sorted(df_main['Month'].unique())}. "
                       f"Filter: {selected_months}. Adjust the sidebar filter.")
        else:
            month_order = {m: i for i, m in enumerate(FISCAL_MONTHS)}
            available_months = sorted(df_filtered['Month'].unique(),
                                      key=lambda m: month_order.get(m, 99))
            latest_month = available_months[-1]
            latest = df_filtered[df_filtered['Month'] == latest_month]

            # ---- RECOMPUTED YTD (robust) ----
            # YTD = sum of Monthly_Actual across all months up to and including latest_month
            months_up_to_latest = [m for m in FISCAL_MONTHS
                                   if month_order.get(m, 99) <= month_order.get(latest_month, 99)]
            ytd_data = df_filtered[df_filtered['Month'].isin(months_up_to_latest)]
            ytd_per_account = ytd_data.groupby(['Code', 'Account']).agg(
                YTD_Actual=('Monthly_Actual', 'sum'),
                Approved_Annual=('Approved_Annual', 'first'),
            ).reset_index()
            ytd_per_account['YTD_Variance'] = ytd_per_account['YTD_Actual'] - ytd_per_account['Approved_Annual'] * (month_order.get(latest_month, 0) + 1) / 12

            total_ytd_revenue = ytd_per_account['YTD_Actual'].sum()
            total_annual_target = ytd_per_account['Approved_Annual'].sum()
            total_ytd_variance = ytd_per_account['YTD_Variance'].sum()
            months_elapsed = month_order.get(latest_month, 0) + 1
            pro_rata_target = total_annual_target * (months_elapsed / 12)

            monthly_rev = df_filtered.groupby('Month')['Monthly_Actual'].sum().reset_index()
            monthly_rev['Order'] = monthly_rev['Month'].map(month_order)
            monthly_rev = monthly_rev.sort_values('Order')

            st.caption(f"Showing **{len(available_months)} month(s)** — Latest: **{latest_month}** — "
                       f"Months elapsed: **{months_elapsed}/12**")

            # --- KPI CARDS ---
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.markdown(f"""<div class="kpi-card"><div class="kpi-title">💰 YTD Revenue</div>
                <div class="kpi-value">${total_ytd_revenue:,.0f}</div>
                <div class="kpi-delta delta-pos">Sum of monthlies</div></div>""", unsafe_allow_html=True)
            with col2:
                st.markdown(f"""<div class="kpi-card"><div class="kpi-title">🎯 YTD Variance</div>
                <div class="kpi-value">${abs(total_ytd_variance):,.0f}</div>
                <div class="kpi-delta {'delta-pos' if total_ytd_variance > 0 else 'delta-neg'}">
                {'Overage' if total_ytd_variance > 0 else 'Shortage'}</div></div>""", unsafe_allow_html=True)
            with col3:
                st.markdown(f"""<div class="kpi-card"><div class="kpi-title">📈 Pro-Rata Target</div>
                <div class="kpi-value">${pro_rata_target:,.0f}</div>
                <div class="kpi-delta">{months_elapsed}/12 Months</div></div>""", unsafe_allow_html=True)
            with col4:
                conv = (total_ytd_revenue / pro_rata_target * 100) if pro_rata_target > 0 else 0
                st.markdown(f"""<div class="kpi-card"><div class="kpi-title">✅ Convergence</div>
                <div class="kpi-value">{conv:.1f}%</div>
                <div class="kpi-delta {'delta-pos' if conv >= 100 else 'delta-neg'}">
                {'On Track' if conv >= 100 else 'Behind'}</div></div>""", unsafe_allow_html=True)
            with col5:
                st.markdown(f"""<div class="kpi-card"><div class="kpi-title">📊 Accounts</div>
                <div class="kpi-value">{ytd_per_account['Code'].nunique()}</div>
                <div class="kpi-delta">Active Tax Lines</div></div>""", unsafe_allow_html=True)

            # --- CHARTS ROW 1 ---
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Status vs Pro-Rata")
                st_df = ytd_per_account.copy()
                st_df['Status'] = st_df.apply(
                    lambda x: 'On Track' if x['YTD_Actual'] >= (x['Approved_Annual'] * months_elapsed / 12) else 'Behind', axis=1)
                counts = st_df['Status'].value_counts().reset_index()
                counts.columns = ['Status', 'Count']
                fig = px.pie(counts, names='Status', values='Count', hole=0.6,
                             color='Status', color_discrete_map={'On Track': '#00B4D8', 'Behind': '#1A3B47'})
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Overall Convergence")
                fig = go.Figure(go.Indicator(mode="gauge+number", value=conv,
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={'axis': {'range': [None, 150]}, 'bar': {'color': "#00B4D8"},
                           'steps': [{'range': [0, 100], 'color': "#E0E0E0"},
                                     {'range': [100, 150], 'color': "#F0F0F0"}]}))
                fig.update_layout(margin=dict(t=30, b=0, l=10, r=10))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            with c3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top 5 YTD Revenue")
                top5 = ytd_per_account.nlargest(5, 'YTD_Actual')[['Account', 'YTD_Actual']]
                fig = px.bar(top5, x='YTD_Actual', y='Account', orientation='h',
                             color='YTD_Actual', color_continuous_scale=['#00B4D8', '#1A3B47'])
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False,
                                  yaxis_title=None, xaxis_title=None)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # --- CHARTS ROW 2 ---
            c4, c5 = st.columns([2, 1])
            with c4:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Monthly Revenue Progression")
                fig = px.line(monthly_rev, x='Month', y='Monthly_Actual', markers=True)
                fig.update_traces(line_color='#00B4D8', marker_size=10)
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0),
                                  xaxis_title=None, yaxis_title="Revenue (XCD)")
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            with c5:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top 10 Shortfalls (YTD)")
                short = ytd_per_account.nsmallest(10, 'YTD_Variance')[['Account', 'YTD_Variance']]
                if not short.empty and short['YTD_Variance'].sum() < 0:
                    fig = px.bar(short, x='YTD_Variance', y='Account', orientation='h',
                                 color='YTD_Variance', color_continuous_scale='Reds')
                    fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("No negative YTD variances 🎉")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- OFFICE BREAKDOWN ---
            if not df_daily.empty:
                st.markdown("### 🏢 Revenue by Collection Office (YTD)")
                office_f = df_daily[df_daily['Month'].isin(selected_months)]
                office_sum = office_f.groupby('Office')['Total_Collected'].sum().reset_index()
                office_sum = office_sum[office_sum['Total_Collected'] > 0].sort_values('Total_Collected', ascending=False)
                if not office_sum.empty:
                    fig = px.bar(office_sum, x='Office', y='Total_Collected',
                                 color='Total_Collected', color_continuous_scale='Blues')
                    st.plotly_chart(fig, use_container_width=True)

            # --- DATA TABLE ---
            st.subheader("📋 Comprehensive YTD Inflow Matrix")
            table = ytd_per_account[['Code', 'Account', 'Approved_Annual', 'YTD_Actual', 'YTD_Variance']].copy()
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

    with st.expander("🔍 ISIC Upload Diagnostics", expanded=df_isic.empty):
        st.dataframe(diag_isic, use_container_width=True)

    if not df_isic.empty:
        st.markdown("---")
        st.header("🏭 ISIC Sector Deep-Dive")
        df_isic_f = df_isic[df_isic['Month'].isin(selected_months)]

        if df_isic_f.empty:
            st.warning(f"ISIC months parsed: {sorted(df_isic['Month'].unique())}. "
                       f"Filter: {selected_months}.")
        else:
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Sector Revenue Contribution")
                summary = df_isic_f.groupby('Account')['Collection'].sum().reset_index()
                fig = px.treemap(summary, path=['Account'], values='Collection',
                                 color='Collection', color_continuous_scale='Blues')
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("Top Sector Variances")
                var_sum = df_isic_f.groupby('Account')['Variance'].sum().reset_index().sort_values('Variance').head(10)
                fig = px.bar(var_sum, x='Variance', y='Account', orientation='h',
                             color='Variance', color_continuous_scale='RdYlGn')
                fig.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

elif not main_files:
    st.info("👈 Upload monthly reports in the sidebar to get started.")
