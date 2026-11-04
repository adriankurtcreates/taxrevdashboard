import streamlit as st
import pandas as pd

# Set up page configurations for IMF presentation standards
st.set_page_config(page_title="Saint Lucia IRD Revenue Dashboard", layout="wide")

# Embedded statutory rules metadata base
TAX_META = {
    "3111 001": {"name": "Personal Income Tax", "ctx": "Predictable baseline employment payroll contributions across important local workforce engines."},
    "3111 002": {"name": "Arrears on Personal Income Tax", "ctx": "Reflect streams facing operational back-tax recovery deficits across target registries."},
    "3112 001": {"name": "Corporate Income Tax", "ctx": "Aligned cleanly with scheduled corporate installment base filing timelines."},
    "3112 002": {"name": "Arrears on Corporate Income Tax", "ctx": "Outperformed expectations driven by macro clearance transitions."},
    "3141 001": {"name": "Value Added Tax (VAT)", "ctx": "Points to stable domestic transaction health and resilient consumption cycles."},
    "3141 006": {"name": "Health and Citizen Security Levy", "ctx": "Primary consumption collections anchored across communication networks."}
}

MONTHS = ["APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER", "JANUARY", "FEBRUARY", "MARCH"]

# Populating vectors with your exact data arrays from the running summary logs
data_actuals = {
    "3111 001": [12368176.08, 12089845.47, 11668394.84, 11777207.93, 10501463.70, 11847649.68, 0, 0, 0, 0, 0, 0],
    "3111 002": [3312220.57, 2860023.48, 2225352.86, 2570577.57, 1504049.91, 953968.67, 0, 0, 0, 0, 0, 0],
    "3112 001": [1415048.69, 1463606.60, 41024880.87, 3327019.16, 1059014.49, 40605384.68, 0, 0, 0, 0, 0, 0],
    "3112 002": [3120385.57, 4873630.09, 1304925.80, 793404.97, 653602.57, 2385422.93, 0, 0, 0, 0, 0, 0],
    "3141 001": [23733957.36, 27230858.61, 21012716.16, 17833795.54, 19234062.01, 18328871.14, 0, 0, 0, 0, 0, 0],
    "3141 006": [1420269.94, 1619359.49, 1306093.91, 1351250.65, 1303207.95, 1385801.94, 0, 0, 0, 0, 0, 0]
}

data_projections = {
    "3111 001": [12368176.00, 10703669.00, 10357647.50, 10566566.28, 10676735.82, 11098530.86, 10678454.36, 10393058.27, 9877114.54, 13762617.20, 11580072.33, 20443516.72],
    "3111 002": [3312220.57, 2425908.75, 1790196.00, 1971125.47, 1456688.20, 1942250.94, 2427813.67, 2791985.72, 2913376.41, 2398939.14, 2679869.00, 3541720.51],
    "3112 001": [1415048.69, 2221323.00, 40460814.00, 2279191.25, 1244187.00, 40555120.00, 6215906.30, 2896048.57, 11003110.00, 1646215.00, 1894724.00, 55516410.00],
    "3112 002": [3120385.57, 4966397.84, 623151.00, 830868.00, 1246302.01, 1661736.01, 2077170.01, 2378745.51, 2492604.01, 2908038.01, 2125755.01, 3115755.01],
    "3141 001": [23733957.36, 15303771.99, 17879005.47, 18073391.02, 18460495.57, 17343832.40, 17220571.45, 15910403.66, 18701154.99, 26877321.57, 23995657.67, 23626045.85],
    "3141 006": [1420269.94, 882046.69, 1755997.81, 2135499.97, 2030704.94, 4108135.39, 1735075.22, 1994095.06, 1979916.84, 2102206.32, 2282915.44, 2248478.38]
}

# Sector mappings from ledger logs
data_sectors = {
    "3111 001": [["5510", "Short term accommodation activities", 1962959.97], ["6419", "Other monetary intermediation", 390596.88], ["5229", "Other transportation support activities", 379939.78]],
    "3111 002": [["5229", "Other transportation support activities", 57758.39], ["6110", "Wired telecommunications activities", 31532.46], ["5510", "Short term accommodation activities", 20579.97]],
    "3112 001": [["6419", "Other monetary intermediation", 5120497.78], ["4721", "Retail sale of food in specialized stores", 5062340.77], ["6110", "Wired telecommunications activities", 3128450.77]],
    "3112 002": [["4690", "Non-specialized wholesale trade", 1651684.58], ["6419", "Other monetary intermediation", 1090209.16], ["4741", "Retail sale of computers in specialized stores", 651418.74]],
    "3141 001": [["5510", "Short term accommodation activities", 3716151.57], ["6120", "Wireless telecommunications activities", 1394599.30], ["4721", "Retail sale of food in specialized stores", 1273633.13]]
}

# APP INTERFACE
st.title("St. Lucia IRD — Revenue Intelligence Workspace")
st.caption("Strategic Research Terminal — Optimized for IMF Chapter IV Review Cycles")

# Sidebar Controller Options
st.sidebar.header("Execution Controls")
scope = st.sidebar.selectbox("Select Timeline Scope", ["September 2026", "Quarter 1 (Apr-Jun)", "Quarter 2 (Jul-Sep)", "Full Fiscal Year-to-Date"])

# Map timeline choices to respective indices
if scope == "September 2026":
    month_indices = [5]
elif scope == "Quarter 1 (Apr-Jun)":
    month_indices = [0, 1, 2]
elif scope == "Quarter 2 (Jul-Sep)":
    month_indices = [3, 4, 5]
else:
    month_indices = [0, 1, 2, 3, 4, 5]

# Process and aggregate rows dynamically
matrix_data = []
total_act, total_proj = 0, 0

for code, meta in TAX_META.items():
    act_val = sum(data_actuals[code][i] for i in month_indices)
    proj_val = sum(data_projections[code][i] for i in month_indices)
    var_val = act_val - proj_val
    perf_pct = (act_val / proj_val * 100) if proj_val > 0 else 0.0
    
    total_act += act_val
    total_proj += proj_val
    
    matrix_data.append({
        "Account Code": code,
        "Description": meta["name"],
        "Projected Target": proj_val,
        "Actual Collected": act_val,
        "Net Variance": var_val,
        "Convergence %": f"{perf_pct:.2f}%"
    })

# High-Level Metrics Layout Display Card Blocks
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Actual Revenues (XCD)", f"$ {total_act:,.2f}")
kpi2.metric("Target Projections (XCD)", f"$ {total_proj:,.2f}")
net_var = total_act - total_proj
kpi3.metric("Net Fiscal Variance", f"$ {net_var:,.2f}", delta=f"{net_var:,.2f}")
kpi4.metric("Convergence Efficiency", f"{(total_act/total_proj*100):.2f}%")

# Main Interactive Data Grid Matrix View
st.subheader(f"Detailed Inflow Summary Matrix — Scope: {scope}")
df = pd.DataFrame(matrix_data)
st.dataframe(df.style.format({
    "Projected Target": "${:,.2f}",
    "Actual Collected": "${:,.2f}",
    "Net Variance": "${:,.2f}"
}), use_container_width=True)

# --------------------------------===================
# PROGRAMMATIC VARIANCE NARRATIVE TEXT COMPILED LIVE
# --------------------------------===================
st.subheader("📝 Automated Variance Notes Briefing Generator")

report_text = f"""SAINT LUCIA\nDEPARTMENT OF FINANCE\nINLAND REVENUE DEPARTMENT\n
Revenue Collection Report — Scope: {scope}
Evaluation Framework for the IMF Article IV Delegation Brief

OVERVIEW
For the targeted timeline period tracking {scope}, the Inland Revenue Department projected gross aggregate receipts of $ {total_proj:,.2f}. Actual collections validation models calculated total realized volume at $ {total_act:,.2f}, indicating a net variation outturn of $ {net_var:,.2f}. 

Note: Data feeds remain provisional matching internal oversight standards. All figures remain subject to final matching reconciliations against collections logs from the Accountant General's Department.

VARIANCE ANALYSIS BY TAX TYPE
"""

for item in matrix_data:
    code = item["Account Code"]
    meta = TAX_META[code]
    v_fmt = f"(${abs(item['Projected Target'] - item['Actual Collected']):,.2f})" if (item['Actual Collected'] < item['Projected Target']) else f"${item['Net Variance']:,.2f}"
    
    report_text += f"""
{code} — {item['Description']}
Projection: $ {item['Projected Target']:,.2f}
Actual: $ {item['Actual Collected']:,.2f}
Variance: {v_fmt} or {item['Convergence %']}

Analytical Assessment:
Collections under this tax head window track functional seasonal parameters. {meta['ctx']} Performance lines show clear baseline convergence matching targeted tracking models for the historical fiscal period.
"""
    if code in data_sectors and scope == "September 2026":
        report_text += "Top Performing Sectors:\n"
        for sec in data_sectors[code]:
            report_text += f"  * ISIC {sec[0]} — {sec[1]}: $ {sec[2]:,.2f}\n"
        report_text += "The performance under this sector profile was heavily anchored by stable corporate collections. The contributing taxpayer profile was driven directly by Large Taxpayers, Medium Taxpayers, and Small and Micro Taxpayers. Macroeconomic factors within these specific industry lines present possible conservative drivers for the underlying value variance. Further observation of these possibilities will have to be made.\n"

report_text += f"\nCONCLUSION\nThe overall performance framework produces an aggregate revenue target convergence rate of {(total_act/total_proj*100):.2f}% matching budgeted parameters. Close monitoring of direct and indirect resource streams will be employed as the Department continues to observe developments in key economic sectors to assess the sustainability of collection patterns.\nReport compiled successfully by the Revenue Research Division."

st.text_area("Live Document Transcript Preview", report_text, height=350)

# Instant File Export Control File Stream Ingestion Link
st.download_button(
    label="📥 Download Formal Brief Document (.txt)",
    data=report_text,
    file_name=f"SLU_IRD_IMF_Briefing_{scope.replace(' ', '_')}.txt",
    mime="text/plain"
)
