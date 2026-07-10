import streamlit as st
import pandas as pd
import io
import base64
import os
from PIL import Image
from openpyxl.styles import Font, numbers, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# --- Handle Logo Loading ---
logo_path = "/content/download.jfif"
try:
    with open(logo_path, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode('utf-8')
    logo_html = f'<img src="data:image/jpeg;base64,{logo_base64}" style="height: 40px; vertical-align: middle; margin-right: 15px; border-radius: 5px;">'
except Exception:
    logo_html = "🚗 "

# --- Page Configuration ---
st.set_page_config(
    page_title="Slice Parking Payouts",
    page_icon="🅿️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom Styling ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

    html, body, [class*='css'] {
        font-family: 'Inter', sans-serif;
    }
    .main-header {font-size: 2.4rem; font-weight: 700; color: #1E3A8A; margin-bottom: 5px; display: flex; align-items: center;}
    .sub-header {font-size: 1.1rem; font-weight: 500; color: #475569; margin-bottom: 25px;}
    .stButton>button {background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: 600; height: 50px; transition: all 0.3s ease; border: none; box-shadow: 0 4px 6px -1px rgba(30, 58, 138, 0.2);}
    .stButton>button:hover {background-color: #1E40AF; color: white; box-shadow: 0 10px 15px -3px rgba(30, 58, 138, 0.3); transform: translateY(-2px);}

    /* Modern Metric Cards */
    .metric-card {background: #FFFFFF; padding: 24px; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); text-align: center;}
    .metric-title {font-size: 0.95rem; font-weight: 600; color: #475569; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px;}
    .metric-value {font-size: 2rem; font-weight: 700; color: #065F46; margin: 0;}
    .metric-highlight {color: #047857;}

    .section-title {font-size: 1.5rem; font-weight: 700; color: #1E3A8A; margin-top: 30px; margin-bottom: 15px; border-bottom: 2px solid #E2E8F0; padding-bottom: 5px;}
    </style>
""", unsafe_allow_html=True)

# --- Main Dashboard ---
st.markdown(f"<div class='main-header'>{logo_html} Slice Parking Payout Generator</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Professional financial reconciliation & automated payout reporting.</div>", unsafe_allow_html=True)

st.markdown("<div class='section-title'>⚙️ Step 1: Configuration</div>", unsafe_allow_html=True)

# Initialize session state for sliders if not present
if 'slice_pct' not in st.session_state:
    st.session_state.slice_pct = 30
if 'client_pct' not in st.session_state:
    st.session_state.client_pct = 70

def update_client():
    st.session_state.client_pct = 100 - st.session_state.slice_pct

def update_slice():
    st.session_state.slice_pct = 100 - st.session_state.client_pct

col0, col1, col2, col3 = st.columns(4)
with col0:
    parking_code = st.selectbox("🅿️ Parking Code", ["BCW", "CP", "TAJ", "GH", "Gator", "HBH", "CHS", "BDCHP", "Billiards", "BlueWave", "Crain", "Dyson", "EastonDemo", "SunBeach", "UUMC"])
with col1:
    target_date = st.date_input("📅 Target Payout Date", value=pd.to_datetime('2026-07-03').date())
with col2:
    slice_pct = st.slider("🍕 Slice %", min_value=0, max_value=100, key="slice_pct", on_change=update_client)
with col3:
    client_pct = st.slider("🤝 Client %", min_value=0, max_value=100, key="client_pct", on_change=update_slice)

if slice_pct + client_pct != 100:
    st.warning("⚠️ Heads up: Your Slice and Client percentages do not add up to 100%.")

# --- Sidebar Inputs ---
with st.sidebar:
    st.markdown(f"<div style='text-align: center; margin-bottom: 20px; font-size: 50px;'>🅿️</div>", unsafe_allow_html=True)
    st.markdown("## 📂 Upload Data Files")

    if parking_code == "HBH":
        main_file = st.file_uploader("1. Main Transaction Data (HBH)", type=["xlsx", "xls"], key="hbh_main")
        enf_file = st.file_uploader("2. Enforcement Data (Payout)", type=["xlsx", "xls"], key="hbh_enf")
        tier_file = st.file_uploader("3. Tier Data (Parkpliant)", type=["xlsx", "xls"], key="hbh_tier")
        sh_file = st.file_uploader("4. Spot Hero Data (HBH Only)", type=["csv"], key="hbh_sh")
    elif parking_code in ["CHS", "BDCHP", "Billiards", "BlueWave", "Crain", "Dyson", "EastonDemo", "SunBeach", "UUMC"]:
        main_file = st.file_uploader("1. Main Transaction Data (Slice Parking)", type=["xlsx", "xls"], key="std_main")
        sh_file = None
        enf_file = None
        tier_file = None
    elif parking_code in ["BCW", "CP", "TAJ"]:
        main_file = st.file_uploader("1. Main Transaction Data (Slice Parking)", type=["xlsx", "xls"], key="std_main")
        enf_file = st.file_uploader("2. Enforcement Data (Payout)", type=["xlsx", "xls"], key="std_enf")
        tier_file = st.file_uploader("3. Tier Data (Parkpliant)", type=["xlsx", "xls"], key="std_tier")
        sh_file = None
    else: # GH, Gator
        main_file = st.file_uploader("1. Main Transaction Data (Slice Parking)", type=["xlsx", "xls"], key="std_main")
        enf_file = st.file_uploader("2. Enforcement Data (Payout)", type=["xlsx", "xls"], key="std_enf")
        tier_file = st.file_uploader("3. Tier Data (Parkpliant)", type=["xlsx", "xls"], key="std_tier")
        sh_file = st.file_uploader("4. Spot Hero Data (CSV)", type=["csv"], key="std_sh")

    st.markdown("---")
    st.markdown("**Need help?** Contact your Slice Parking admin.")

st.markdown("<div class='section-title'>🚀 Step 2: Generate Report</div>", unsafe_allow_html=True)
generate_btn = st.button("Generate Payout Report", use_container_width=True)

# --- Main Processing Logic ---
if generate_btn:
    if not main_file:
        st.error("🛑 Please upload at least the Main Transaction Data file.")
    else:
        with st.spinner("Crunching the numbers... Please wait..."):
            try:
                df = pd.read_excel(main_file)
                er = pd.read_excel(enf_file) if enf_file else pd.DataFrame()
                pp = pd.read_excel(tier_file) if tier_file else pd.DataFrame()
                sh = pd.read_csv(sh_file) if sh_file else pd.DataFrame()

                target_date_pd = pd.to_datetime(target_date).date()
                if 'Payout Date' in df.columns:
                    df['Payout Date'] = pd.to_datetime(df['Payout Date'], errors='coerce').dt.date

                # Global metric defaults
                Enforce, sh_sum, net_calc, total_net, payout, specific_sum_pp = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
                cpu_gross, cpu_service, cpu_credit, cpu_subtotal, cpu_dock, cpu_net, cpu_total_net, cpu_Enforce, cpu_specific_sum_pp = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
                cp_gross, cp_service, cp_credit, cp_subtotal, cp_dock, cp_net, cp_total_net, cp_Enforce, cp_specific_sum_pp = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

                if not pp.empty:
                    T1, T2, T3, Te, L1, L2 = 1.5, 2, 2, 1, 2, 2
                    pp_sum_series = (pp['Tier 1 Lookup']*T1 + pp['Tier 2 Lookup']*T2 +
                                     pp['Tier 3 Lookup']*T3 + pp['Text 1']*Te +
                                     pp['Letter 1']*L1 + pp['Letter 2']*L2)
                else:
                    pp_sum_series = pd.Series()

                df_main = df[(df['Parking Code'] == parking_code) & (df['Payout Date'] == target_date_pd)]
                gross = float(df_main['Total Amount'].sum()) if not df_main.empty else 0.0
                service = float(df_main['Service Fee'].sum()) if not df_main.empty else 0.0
                credit = float(df_main['Payment Gateway Fee'].sum()) if 'Payment Gateway Fee' in df_main.columns else 0.0
                tax = 0.0

                if parking_code == "BCW":
                    tax = float(df_main['City Tax'].sum()) if not df_main.empty else 0.0
                    subtotal = float(gross - service - credit - tax)
                elif parking_code == 'CP':
                    cp_gross = gross
                    cp_service = service
                    cp_credit = credit
                    cp_subtotal = float(cp_gross - cp_service - cp_credit)

                    df_cpu = df[(df['Parking Code'] == 'CPUpperLot') & (df['Payout Date'] == target_date_pd)]
                    if not df_cpu.empty:
                        cpu_gross = float(df_cpu['Total Amount'].sum())
                        cpu_service = float(df_cpu['Service Fee'].sum())
                        cpu_credit = float(df_cpu['Payment Gateway Fee'].sum()) if 'Payment Gateway Fee' in df_cpu.columns else 0.0
                        cpu_subtotal = float(cpu_gross - cpu_service - cpu_credit)
                    subtotal = cp_subtotal # UI metric focuses on the main code selected
                else:
                    subtotal = float(gross - service - credit)

                dock_pct = 0.05 if parking_code in ["HBH", "Billiards", "Dyson"] else 0.10
                dock = float(subtotal * dock_pct)
                net = float(subtotal - dock)

                if parking_code == 'CP':
                    cp_dock = float(cp_subtotal * dock_pct)
                    cp_net = float(cp_subtotal - cp_dock)
                    cpu_dock = float(cpu_subtotal * dock_pct)
                    cpu_net = float(cpu_subtotal - cpu_dock)

                if parking_code == "BCW":
                    filtered_er = er[er['Lot Code'] == 'Baltimore Clayworks']
                    sum_net_enf = float(filtered_er['Net'].sum()) if not filtered_er.empty else 0.0
                    specific_sum_pp = float(pp_sum_series.iloc[1]) if len(pp_sum_series) > 1 else 0.0
                    Enforce = float(sum_net_enf - specific_sum_pp)
                    bcws_all_rows = df[(df['Parking Code'] == 'BCWS') & (df['Payout Date'] == target_date_pd)]
                    total_calc = float(bcws_all_rows['Total Amount'].sum() - bcws_all_rows['Payment Gateway Fee'].sum()) if not bcws_all_rows.empty else 0.0
                    net_calc = float(total_calc * 0.90)
                    total_net = float(net + net_calc + Enforce)
                elif parking_code == "CP":
                    filtered_cp = er[er['Lot Code'] == 'CP']
                    cp_sum_net_enf = float(filtered_cp['Net'].sum()) if not filtered_cp.empty else 0.0
                    cp_specific_sum_pp = float(pp_sum_series.iloc[3]) if len(pp_sum_series) > 3 else 0.0
                    cp_Enforce = float(cp_sum_net_enf - cp_specific_sum_pp)
                    cp_total_net = float(cp_net + cp_Enforce)

                    filtered_cpu = er[er['Lot Code'] == 'CPUpperLot']
                    cpu_sum_net_enf = float(filtered_cpu['Net'].sum()) if not filtered_cpu.empty else 0.0
                    cpu_Enforce = float(cpu_sum_net_enf) # Assuming no parkpliant deductor for upper
                    cpu_total_net = float(cpu_net + cpu_Enforce)

                    total_net = float(cp_total_net + cpu_total_net)
                else:
                    total_net = net

                payout = float(total_net * (client_pct / 100))

                st.success(f"Report Generated for {parking_code} on {target_date}.")

                # UI Summary
                metrics = ['Gross Revenue', 'Service Fee', 'Payment Gateway Fee', 'Subtotal', 'Tech Fee', 'Net Revenue', 'Total Net Revenue', f'Total Payout ({client_pct}%)']
                values = [gross, service, credit, subtotal, dock, net, total_net, payout]
                st.table(pd.DataFrame({'Metric': metrics, 'Value': values}).style.format({"Value": "${:,.2f}"}))

                # Excel Export
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    columns_to_drop = ['Application Fee', 'Month', 'Time', 'Date', 'County Tax', 'Owner Payout', 'Client Payout', 'Operator Payout', 'Dock Revenue', 'Purpose', 'id', 'Total Refund', 'Refund Transaction ID', 'Customer ID', 'Payment Method ID', 'Transaction ID', 'Year', 'Weekday', 'Hour', 'Payout Date', 'Payout Difference', 'Difference', 'Failed Reason', 'Start Date', 'End Date', 'Rate', 'Extend Reminder Sent?', 'Extended Reservation', 'Validation Code', 'Email', 'Name', 'Space Number', 'City Tax']
                    df_main_export = df_main.drop(columns=[c for c in columns_to_drop if c in df_main.columns]).copy()
                    df_main_export.to_excel(writer, sheet_name='Report', index=False, startrow=0)

                    worksheet = writer.sheets['Report']
                    current_row = len(df_main_export) + 3

                    main_sum_df = pd.DataFrame({'Metric': metrics, 'Value': values})
                    main_sum_df.to_excel(writer, sheet_name='Report', index=False, startrow=current_row)
                    current_row += len(main_sum_df) + 4

                    if parking_code == 'CP' and not df_cpu.empty:
                        cpu_export = df_cpu.drop(columns=[c for c in columns_to_drop if c in df_cpu.columns]).copy()
                        cpu_export.to_excel(writer, sheet_name='Report', index=False, startrow=current_row)
                        current_row += len(cpu_export) + 2

                        cpu_metrics = ['CP_Upper_Lot Gross', 'Service Fee', 'Gateway Fee', 'Sub Total', 'Tech Fee', 'Net Revenue', 'Total Net']
                        cpu_values = [cpu_gross, cpu_service, cpu_credit, cpu_subtotal, cpu_dock, cpu_net, cpu_total_net]
                        cpu_sum_df = pd.DataFrame({'Metric': cpu_metrics, 'Value': cpu_values})
                        cpu_sum_df.to_excel(writer, sheet_name='Report', index=False, startrow=current_row)

                st.download_button(label="Download Excel Report", data=buffer.getvalue(), file_name=f"{parking_code}_Report.xlsx")

            except Exception as e:
                st.error(f"Error: {e}")
