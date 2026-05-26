import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import os
import io

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Rental Automation",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- TEMPLATE UTILITIES ---
def load_template(filename, directory="templates"):
    """Reads raw file contents from templates directory for dynamic HTML/CSS injection."""
    path = os.path.join(directory, filename)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return ""
    return ""

# Inject custom stylesheet from templates directory
style_css = load_template("style.css")
if style_css:
    st.markdown(f"<style>{style_css}</style>", unsafe_allow_html=True)


# --- DYNAMIC EXCEL PARSER ---
def parse_input_xlsx(file_bytes):
    """Parses user-uploaded input Excel sheet with extensive error resilience."""
    try:
        df = pd.read_excel(file_bytes, sheet_name="Main")
        df.columns = [c.strip() for c in df.columns]
        
        # Build dictionary from Name-Value pairs
        raw_params = {}
        for idx, row in df.iterrows():
            if 'Field Name' in df.columns and 'Sample Value' in df.columns:
                name = str(row['Field Name']).strip()
                val = row['Sample Value']
                if pd.isna(val):
                    val = None
                raw_params[name] = val
        
        # Helper to parse dates
        def parse_date(date_val, default=None):
            if date_val is None:
                return default
            if isinstance(date_val, (datetime.date, datetime.datetime)):
                return date_val.date() if isinstance(date_val, datetime.datetime) else date_val
            try:
                return pd.to_datetime(str(date_val).strip()).date()
            except Exception:
                return default

        # Clean/Normalize values
        params = {}
        params["Lease ID"] = raw_params.get("Lease ID", "LEASE-001")
        params["REU Name"] = raw_params.get("REU Name", "IN-CHEK2")
        params["Lease Type"] = raw_params.get("Lease Type", "Office")
        params["Building Name"] = raw_params.get("Building Name", "SKCL Tech Park")
        params["City"] = raw_params.get("City", "Chennai")
        params["Country"] = raw_params.get("Country", "India")
        params["Currency"] = raw_params.get("Currency", "INR")
        
        params["Chargeable Area Sqft"] = float(raw_params.get("Chargeable Area Sqft", 9515))
        params["Chargeable Area Sqm"] = params["Chargeable Area Sqft"] / 10.764
        params["Parking Slots"] = int(raw_params.get("Parking Slots", 20))
        
        start_date = parse_date(raw_params.get("Agreement Start Date"), datetime.date(2026, 4, 1))
        end_date = parse_date(raw_params.get("Agreement End Date"), datetime.date(2030, 3, 31))
        params["Agreement Start Date"] = start_date
        params["Agreement End Date"] = end_date
        params["Rent Start Date"] = parse_date(raw_params.get("Rent Start Date"), start_date)
        
        # Escalation rate
        params["Rent Per Sqft"] = float(raw_params.get("Rent Per Sqft", 120))
        
        esc_val = raw_params.get("Escalation %", 0.15)
        params["Escalation %"] = float(esc_val) if esc_val is not None else 0.15
        
        # Robust check for frequency e.g. 0.36 -> 36 months
        freq_val = raw_params.get("Escalation Frequency Months", 36)
        if freq_val is not None:
            freq_val = float(freq_val)
            if freq_val < 1.0:
                freq_val = int(round(freq_val * 100))
            else:
                freq_val = int(freq_val)
        else:
            freq_val = 36
        params["Escalation Frequency Months"] = freq_val
        
        params["Billing Frequency"] = raw_params.get("Billing Frequency", "Monthly")
        params["Security Deposit Months"] = float(raw_params.get("Security Deposit Months", 6))
        params["Security Deposit Amount"] = float(raw_params.get("Security Deposit Amount", 11418000))
        params["Refundable Deposit"] = raw_params.get("Refundable Deposit", "Yes")
        
        # Capex & Cost factors
        params["Fitout Cost"] = float(raw_params.get("Fitout Cost", 2000000))
        params["Useful Life Years"] = float(raw_params.get("Useful Life Years", 5))
        params["Residual Value"] = float(raw_params.get("Residual Value", 0))
        params["Discount Rate"] = float(raw_params.get("Discount Rate", 0.08))
        params["Incremental Borrowing Rate"] = float(raw_params.get("Incremental Borrowing Rate", 0.08))
        params["Cost of Capital"] = float(raw_params.get("Cost of Capital", 0.15))
        params["Addnl.Deposit -energy(Refundable)"] = float(raw_params.get("Addnl.Deposit -energy(Refundable)", 500000))
        
        # Handle CAM override (defaults to 15.48 per sqft if not found)
        cam_rate = None
        for k, v in raw_params.items():
            if "CAM" in k or "Maintenance" in k:
                try:
                    cam_rate = float(v)
                    break
                except Exception:
                    pass
        params["Quoted CAM"] = cam_rate if cam_rate is not None else 15.48
        params["CAM Escalation %"] = 0.05
        
        return params, None
    except Exception as e:
        return {}, str(e)


# --- WORKBOOK INJECTION GENERATOR ---
def generate_output_workbook(template_path, params):
    """Loads specimen template, updates input-driven parameters, and saves new workbook."""
    wb = openpyxl.load_workbook(template_path, data_only=False)
    
    # 1. Update Main sheet
    ws_main = wb["Main"]
    ws_main["B2"] = params["REU Name"]
    ws_main["B3"] = params["Chargeable Area Sqft"]
    ws_main["C3"] = "=B3/10.764"
    ws_main["B4"] = params["Agreement Start Date"]
    ws_main["B5"] = params["Agreement End Date"]
    ws_main["B6"] = "=(B5-B4)/365"
    ws_main["B7"] = params["Rent Per Sqft"]
    ws_main["B8"] = params["Quoted CAM"]
    ws_main["B9"] = params["Security Deposit Amount"]
    ws_main["B10"] = params["Fitout Cost"] * 1.2
    ws_main["B12"] = params["Cost of Capital"]
    ws_main["B13"] = params["Cost of Capital"] / 2.0
    
    # Text strings for escalation description
    rent_esc_freq_years = int(params["Escalation Frequency Months"] // 12)
    ws_main["B14"] = f"{int(params['Escalation %'] * 100)}% every {rent_esc_freq_years} years"
    ws_main["B15"] = f"{int(params['CAM Escalation %'] * 100)}% every years"
    ws_main["B17"] = params["Addnl.Deposit -energy(Refundable)"]
    
    # 2. Update Lease Rent sheet starting values and dates
    ws_rent = wb["Lease Rent"]
    
    # Schedule Start Date is Jan 1st of the starting year
    start_year = params["Agreement Start Date"].year
    schedule_start = datetime.date(start_year, 1, 1)
    
    ws_rent["A31"] = start_year
    ws_rent["B31"] = schedule_start
    ws_rent["H31"] = "=Main!B7"
    ws_rent["J31"] = "=Main!B8"
    
    # Update Years dynamically so financial summaries display correctly
    ws_rent["G15"] = "=A31"
    
    # Save to a memory stream for download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- PYTHON RENT SCHEDULE SIMULATOR ---
def simulate_schedule(params):
    """Simulates month-by-month rent/CAM escalation exactly like Excel formulas."""
    start_year = params["Agreement Start Date"].year
    start_date = datetime.date(start_year, 1, 1)
    
    rows = []
    current_from = start_date
    area_sqft = params["Chargeable Area Sqft"]
    base_rent = params["Rent Per Sqft"]
    base_cam = params["Quoted CAM"]
    
    for m in range(1, 73):
        # Fiscal Year classification
        if m <= 9:
            fy_year = start_year
        elif 10 <= m <= 21:
            fy_year = start_year + 1
        elif 22 <= m <= 33:
            fy_year = start_year + 2
        elif 34 <= m <= 45:
            fy_year = start_year + 3
        elif 46 <= m <= 57:
            fy_year = start_year + 4
        elif 58 <= m <= 69:
            fy_year = start_year + 5
        else:
            fy_year = start_year + 6
            
        # To Date calculation
        if m in [2, 14, 26, 38, 50, 62]:
            is_leap = (current_from.year % 4 == 0 and (current_from.year % 100 != 0 or current_from.year % 400 == 0))
            days = 29 if is_leap else 28
        elif m in [4, 6, 7, 16, 18, 20, 28, 30, 32, 40, 42, 44, 52, 54, 56, 64, 66, 68, 70, 72]:
            days = 30
        else:
            days = 31
            
        current_to = current_from + datetime.timedelta(days=days-1)
        
        # Rent rate (escalation rate every escalation frequency months)
        if m <= params["Escalation Frequency Months"]:
            rent_rate = base_rent
        else:
            rent_rate = base_rent * (1.0 + params["Escalation %"])
            
        # CAM rate (5% escalation in Month 14, 26, 37, 50, 62)
        if m <= 13:
            cam_rate = base_cam
        elif 14 <= m <= 25:
            cam_rate = base_cam * 1.05
        elif 26 <= m <= 36:
            cam_rate = base_cam * 1.05 * 1.05
        elif 37 <= m <= 49:
            cam_rate = base_cam * 1.05 * 1.05 * 1.05
        elif 50 <= m <= 61:
            cam_rate = base_cam * 1.05 * 1.05 * 1.05 * 1.05
        else:
            cam_rate = base_cam * 1.05 * 1.05 * 1.05 * 1.05 * 1.05
            
        rent_amt = rent_rate * area_sqft
        cam_amt = cam_rate * area_sqft
        total_amt = rent_amt + cam_amt
        
        rows.append({
            "Month #": m,
            "Fiscal Year": fy_year,
            "From Date": current_from.strftime("%Y-%m-%d"),
            "To Date": current_to.strftime("%Y-%m-%d"),
            "Area (Sqft)": area_sqft,
            "Rent Rate ($/Sqft)": round(rent_rate, 4),
            "Rent Amount": round(rent_amt, 2),
            "CAM Rate ($/Sqft)": round(cam_rate, 4),
            "CAM Amount": round(cam_amt, 2),
            "Total Billing": round(total_amt, 2)
        })
        
        # Prepare next month start
        current_from = current_to + datetime.timedelta(days=1)
        
    return pd.DataFrame(rows)


# --- SECURITY DEPOSIT & ARO SIMULATOR ---
def compute_deposit_working(params, schedule_df):
    """Simulates carrying costs and ARO parameters from inputs."""
    sd_amount = params["Security Deposit Amount"]
    rate_capital = params["Cost of Capital"]
    term_months = 72
    
    # Interest carrying costs calculations
    sd_interest = sd_amount * rate_capital / 12 * term_months
    energy_deposit = params["Addnl.Deposit -energy(Refundable)"]
    energy_interest = energy_deposit * rate_capital * 3
    
    fitout_interest = 0.0
    total_sd = sd_amount + energy_deposit
    total_interest = sd_interest + energy_interest
    
    carrying_cost_per_sqft = total_interest / (params["Chargeable Area Sqft"] * term_months)
    
    # ARO calculations
    area = params["Chargeable Area Sqft"]
    aro_rate = 82.6 if area < 50000 else 62.72
    total_aro_cost = area * aro_rate
    aro_per_month = total_aro_cost / term_months
    conversion_factor = aro_per_month / area
    
    results = {
        "Fitout Deposit": (0.0, fitout_interest),
        "Security Deposit": (sd_amount, sd_interest),
        "Energy Deposit": (energy_deposit, energy_interest),
        "Total sd": (total_sd, total_interest),
        "Carrying Cost per Sqft": carrying_cost_per_sqft,
        "ARO Rate": aro_rate,
        "Total ARO Cost": total_aro_cost,
        "ARO per Month": aro_per_month,
        "ARO Conversion Factor": conversion_factor
    }
    return results


# --- STREAMLIT UI LAYOUT & CONTROLS ---

# 1. Render custom premium header
header_html = load_template("header.html")
if header_html:
    st.markdown(header_html, unsafe_allow_html=True)
else:
    st.title("🏢 Rental Automation")

# 2. Sidebar Upload and Parameter Panel Configuration
st.sidebar.markdown('<p style="font-size: 1.3rem; font-weight:700; color:#f8fafc; margin-bottom: 15px;">📁 Data Ingestion</p>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Upload Lease Parameter Workbook (.xlsx)", type=["xlsx"])

# Default fallback values matching specimen
default_params = {
    "REU Name": "IN-CHEK2",
    "Building Name": "SKCL Tech Park",
    "City": "Chennai",
    "Country": "India",
    "Currency": "INR",
    "Chargeable Area Sqft": 9515.0,
    "Agreement Start Date": datetime.date(2026, 4, 1),
    "Agreement End Date": datetime.date(2030, 3, 31),
    "Rent Per Sqft": 120.0,
    "Quoted CAM": 15.48,
    "Escalation %": 0.15,
    "Escalation Frequency Months": 36,
    "CAM Escalation %": 0.05,
    "Security Deposit Amount": 11418000.0,
    "Addnl.Deposit -energy(Refundable)": 500000.0,
    "Fitout Cost": 2000000.0,
    "Cost of Capital": 0.15
}

params = default_params
parse_error_msg = None
upload_occurred = False

# Process upload if present
if uploaded_file is not None:
    upload_occurred = True
    parsed_vals, err = parse_input_xlsx(uploaded_file)
    if err:
        parse_error_msg = err
    else:
        params = parsed_vals

# Render sidebar override controls (dynamic inputs)
st.sidebar.markdown('<p style="font-size: 1.25rem; font-weight:600; color:#cbd5e1; margin-top: 20px;">⚙️ Parameter Configuration</p>', unsafe_allow_html=True)

reu_name = st.sidebar.text_input("Real Estate Unit (REU) Name", value=params["REU Name"])
bldg_name = st.sidebar.text_input("Building Name", value=params["Building Name"])
city = st.sidebar.text_input("City", value=params["City"])
area = st.sidebar.number_input("Chargeable Area (Sq.ft.)", value=int(params["Chargeable Area Sqft"]), step=100)

col_d1, col_d2 = st.sidebar.columns(2)
with col_d1:
    start_date = st.date_input("Start Date", value=params["Agreement Start Date"])
with col_d2:
    end_date = st.date_input("End Date", value=params["Agreement End Date"])

rent_sqft = st.sidebar.number_input("Quoted Rentals (per Sqft/mo)", value=float(params["Rent Per Sqft"]), step=1.0)
cam_sqft = st.sidebar.number_input("Quoted CAM (per Sqft/mo)", value=float(params["Quoted CAM"]), step=0.1)

rent_esc = st.sidebar.slider("Rent Escalation %", min_value=0.0, max_value=0.5, value=float(params["Escalation %"]), step=0.01)
rent_esc_freq = st.sidebar.number_input("Rent Escalation Frequency (Months)", value=int(params["Escalation Frequency Months"]), step=12)

cam_esc = st.sidebar.slider("CAM Escalation %", min_value=0.0, max_value=0.2, value=float(params["CAM Escalation %"]), step=0.01)

sec_deposit = st.sidebar.number_input("Security Deposit Amount", value=float(params["Security Deposit Amount"]), step=50000.0)
capital_rate = st.sidebar.slider("Cost of Capital %", min_value=0.0, max_value=0.3, value=float(params["Cost of Capital"]), step=0.01)
energy_dep = st.sidebar.number_input("Energy Deposit Amount", value=float(params["Addnl.Deposit -energy(Refundable)"]), step=10000.0)
fitout_cost = st.sidebar.number_input("Capex (Fitout cost)", value=float(params["Fitout Cost"]), step=50000.0)

# Pack active values into parameters dictionary
ui_params = {
    "REU Name": reu_name,
    "Building Name": bldg_name,
    "City": city,
    "Country": params["Country"],
    "Currency": params["Currency"],
    "Chargeable Area Sqft": float(area),
    "Agreement Start Date": start_date,
    "Agreement End Date": end_date,
    "Rent Per Sqft": float(rent_sqft),
    "Quoted CAM": float(cam_sqft),
    "Escalation %": float(rent_esc),
    "Escalation Frequency Months": int(rent_esc_freq),
    "CAM Escalation %": float(cam_esc),
    "Security Deposit Amount": float(sec_deposit),
    "Cost of Capital": float(capital_rate),
    "Addnl.Deposit -energy(Refundable)": float(energy_dep),
    "Fitout Cost": float(fitout_cost)
}

# Run live models / schedule simulations in memory
schedule_df = simulate_schedule(ui_params)
sd_results = compute_deposit_working(ui_params, schedule_df)

# Initialize session state for download status
if "download_clicked" not in st.session_state:
    st.session_state.download_clicked = False

# Signature of parameters to auto-hide previews if values are modified
current_sig = f"{reu_name}-{bldg_name}-{city}-{area}-{start_date}-{end_date}-{rent_sqft}-{cam_sqft}-{rent_esc}-{rent_esc_freq}-{cam_esc}-{sec_deposit}-{capital_rate}-{energy_dep}-{fitout_cost}"
if "last_params_sig" not in st.session_state:
    st.session_state.last_params_sig = current_sig

if st.session_state.last_params_sig != current_sig:
    st.session_state.download_clicked = False
    st.session_state.last_params_sig = current_sig

def handle_download():
    st.session_state.download_clicked = True

# --- DYNAMIC STATUS CARD SECTION ---
if parse_error_msg:
    error_tpl = load_template("error_card.html")
    if error_tpl:
        st.markdown(error_tpl.format(error_message=f"Failed to parse sheet: {parse_error_msg}"), unsafe_allow_html=True)
    else:
        st.error(f"Failed to parse sheet: {parse_error_msg}")
elif upload_occurred:
    success_tpl = load_template("success_card.html")
    if success_tpl:
        total_sd = ui_params["Security Deposit Amount"] + ui_params["Addnl.Deposit -energy(Refundable)"]
        duration_yrs = (ui_params["Agreement End Date"] - ui_params["Agreement Start Date"]).days / 365.0
        st.markdown(success_tpl.format(
            reu_name=ui_params["REU Name"],
            area=f"{int(ui_params['Chargeable Area Sqft']):,}",
            duration=f"{duration_yrs:.2f}",
            currency=ui_params["Currency"],
            total_deposit=f"{total_sd:,.2f}"
        ), unsafe_allow_html=True)
    else:
        st.success(f"Workbook parameters for {ui_params['REU Name']} compiled successfully!")
else:
    info_tpl = load_template("info_card.html")
    if info_tpl:
        st.markdown(info_tpl, unsafe_allow_html=True)

# --- LIVE METRIC DASHBOARD (CUSTOM BESPOKE HTML GRID) ---
total_rent_cam = ui_params["Rent Per Sqft"] + ui_params["Quoted CAM"]
duration_yrs = (ui_params["Agreement End Date"] - ui_params["Agreement Start Date"]).days / 365.0
total_sd_amount = ui_params["Security Deposit Amount"] + ui_params["Addnl.Deposit -energy(Refundable)"]

metrics_html = f"""
<div class="stats-grid">
    <div class="stat-cell">
        <span class="stat-label">Chargeable Area</span>
        <span class="stat-value">{int(ui_params['Chargeable Area Sqft']):,} sqft</span>
        <span class="stat-meta">≈ {ui_params['Chargeable Area Sqft']/10.764:,.2f} sq.m.</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Initial Rent + CAM Rate</span>
        <span class="stat-value">{ui_params['Currency']} {total_rent_cam:.2f}</span>
        <span class="stat-meta">Rent: {ui_params['Rent Per Sqft']:.1f} | CAM: {ui_params['Quoted CAM']:.2f}</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Total Deposits</span>
        <span class="stat-value">{ui_params['Currency']} {total_sd_amount:,.2f}</span>
        <span class="stat-meta">Capital cost rate: {ui_params['Cost of Capital']*100:.1f}%</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Lease Duration</span>
        <span class="stat-value">{duration_yrs:.2f} yrs</span>
        <span class="stat-meta">{ui_params['Agreement Start Date'].strftime('%b %d, %Y')} to {ui_params['Agreement End Date'].strftime('%b %d, %Y')}</span>
    </div>
</div>
"""
st.markdown(metrics_html, unsafe_allow_html=True)


# --- EXCEL WORKBOOK GENERATOR ---
template_file_path = r"c:\Users\z004df5r\Documents\rental_automation_v1\artifacts\rental-specimen.xlsx"

if os.path.exists(template_file_path):
    try:
        output_stream = generate_output_workbook(template_file_path, ui_params)
        
        st.download_button(
            label="💾 Generate and Download Excel Workbook",
            data=output_stream,
            file_name=f"rental_sheet_{ui_params['REU Name']}_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=handle_download
        )
    except Exception as e:
        st.error(f"Excel template compilation failed: {e}")
else:
    st.warning(f"Excel template specimen not found in path: {template_file_path}")


# --- SHEET CONTENT PREVIEWS (ELEGANT INTERACTIVE EXPANDERS) ---
if st.session_state.download_clicked:
    st.markdown("### 🔍 Live Excel Worksheet Previews")
    
    # Expander 1: Main Sheet Preview
    with st.expander("📂 Sheet 1: Main Summary Properties"):
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.markdown("<h4>🏢 Sheet 1 Variables Preview</h4>", unsafe_allow_html=True)
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown(f"**REU Name:** <code class='premium-code'>{ui_params['REU Name']}</code>", unsafe_allow_html=True)
            st.markdown(f"**Building Name:** <code class='premium-code'>{ui_params['Building Name']}</code>", unsafe_allow_html=True)
            st.markdown(f"**City & Country:** <code class='premium-code'>{ui_params['City']}, {ui_params['Country']}</code>", unsafe_allow_html=True)
            st.markdown(f"**Chargeable Area:** <code class='premium-code'>{ui_params['Chargeable Area Sqft']:,} Sq.ft.</code>", unsafe_allow_html=True)
            st.markdown(f"**Agreement Term:** <code class='premium-code'>{ui_params['Agreement Start Date'].strftime('%d-%b-%Y')} to {ui_params['Agreement End Date'].strftime('%d-%b-%Y')}</code>", unsafe_allow_html=True)
        with col_m2:
            st.markdown(f"**Rent Rate / Sqft / Mo:** <code class='premium-code'>{ui_params['Currency']} {ui_params['Rent Per Sqft']:.2f}</code>", unsafe_allow_html=True)
            st.markdown(f"**Quoted CAM / Sqft / Mo:** <code class='premium-code'>{ui_params['Currency']} {ui_params['Quoted CAM']:.2f}</code>", unsafe_allow_html=True)
            st.markdown(f"**Security Deposit:** <code class='premium-code'>{ui_params['Currency']} {ui_params['Security Deposit Amount']:,.2f}</code>", unsafe_allow_html=True)
            st.markdown(f"**Rent Escalation Clause:** <code class='premium-code'>{int(ui_params['Escalation %'] * 100)}% every {int(ui_params['Escalation Frequency Months'] // 12)} years</code>", unsafe_allow_html=True)
            st.markdown(f"**CAM Escalation Clause:** <code class='premium-code'>{int(ui_params['CAM Escalation %'] * 100)}% every year</code>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Expander 2: Lease Rent Schedule
    with st.expander("📈 Sheet 2: Monthly Lease Schedule Preview"):
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.markdown("<h4>📋 Rent Schedule Table View</h4>", unsafe_allow_html=True)
        
        # Beautifully styled formatted pandas dataframe
        display_df = schedule_df.copy()
        display_df["Rent Amount"] = display_df["Rent Amount"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
        display_df["CAM Amount"] = display_df["CAM Amount"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
        display_df["Total Billing"] = display_df["Total Billing"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
        
        st.dataframe(display_df, use_container_width=True, height=350)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Expander 3: Security Deposit Workings
    with st.expander("🛡️ Sheet 3: Security Deposit & ARO Working"):
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.markdown("<h4>🛡️ SD Imputed Interest & ARO Workings</h4>", unsafe_allow_html=True)
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("##### Carrying Cost Calculations")
            sd_data = [
                {"Deposit Component": "Fitout Deposit", "Amount": f"{ui_params['Currency']} {sd_results['Fitout Deposit'][0]:,.2f}", "Carrying Interest Cost": f"{ui_params['Currency']} {sd_results['Fitout Deposit'][1]:,.2f}"},
                {"Deposit Component": "Security Deposit", "Amount": f"{ui_params['Currency']} {sd_results['Security Deposit'][0]:,.2f}", "Carrying Interest Cost": f"{ui_params['Currency']} {sd_results['Security Deposit'][1]:,.2f}"},
                {"Deposit Component": "Energy Deposit", "Amount": f"{ui_params['Currency']} {sd_results['Energy Deposit'][0]:,.2f}", "Carrying Interest Cost": f"{ui_params['Currency']} {sd_results['Energy Deposit'][1]:,.2f}"},
                {"Deposit Component": "Total Deposits", "Amount": f"{ui_params['Currency']} {sd_results['Total sd'][0]:,.2f}", "Carrying Interest Cost": f"{ui_params['Currency']} {sd_results['Total sd'][1]:,.2f}"}
            ]
            st.table(pd.DataFrame(sd_data))
            st.markdown(f"**Carrying Cost Rate per Sqft:** `{sd_results['Carrying Cost per Sqft']:.6f} {ui_params['Currency']}/Sqft/mo`")
        
        with col_d2:
            st.markdown("##### ARO (Asset Retirement Obligation) Workings")
            aro_data = [
                {"Property Description": "Area (Sq.ft.)", "Value": f"{area:,} Sq.ft."},
                {"Property Description": "ARO Cost Rate per Sqft (as on 2017)", "Value": f"{sd_results['ARO Rate']:.2f}"},
                {"Property Description": "Total ARO Capital Cost Asset", "Value": f"{ui_params['Currency']} {sd_results['Total ARO Cost']:,.2f}"},
                {"Property Description": "Monthly ARO Amortization Cost", "Value": f"{ui_params['Currency']} {sd_results['ARO per Month']:,.2f}"},
                {"Property Description": "Conversion Factor", "Value": f"{sd_results['ARO Conversion Factor']:.6f}"}
            ]
            st.table(pd.DataFrame(aro_data))
        st.markdown('</div>', unsafe_allow_html=True)
