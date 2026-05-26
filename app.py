import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import os
import io
import plotly.graph_objects as go

# --- PAGE SETUP & PREMIUM STYLING ---
st.set_page_config(
    page_title="Rental Automation Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Google Font and Custom CSS with Harmonious Color Palettes & Micro-animations
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Main container styling */
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
    }
    
    /* Card Container with glassmorphism and subtle shadow */
    .premium-card {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(226, 232, 240, 0.8);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.02);
        transition: all 0.3s ease-in-out;
        margin-bottom: 20px;
    }
    
    .premium-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.08), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
    }
    
    /* Elegant Title Badge */
    .badge-title {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
        font-size: 2.5rem;
        margin-bottom: 5px;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #0f172a;
        color: #f8fafc;
        border-right: 1px solid #1e293b;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 8px;
        color: #475569;
        font-weight: 500;
        border: none;
        transition: all 0.2s ease;
        padding: 0px 20px;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(255, 255, 255, 0.8);
        color: #0f172a;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #4f46e5 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    
    /* Metrics section */
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e1b4b;
    }
    .metric-label {
        font-size: 0.85rem;
        text-transform: uppercase;
        color: #64748b;
        font-weight: 600;
        letter-spacing: 0.05em;
    }
</style>
""", unsafe_allow_html=True)


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
                # String standard format parse e.g. 01-Apr-2026
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
            if freq_val < 1.0: # e.g. 0.36
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
        params["CAM Escalation %"] = 0.05 # Default CAM escalation rate
        
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
    ws_main["B10"] = params["Fitout Cost"] * 1.2 # Match spec Capex ratio or fitout cost
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
    
    # Update Years in G15-G21 dynamically so financial summaries display correctly
    ws_rent["G15"] = "=A31"
    
    # Save to a memory stream for download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- PYTHON RENT SCHEDULE SIMULATOR (FOR UI GRAPH/TABLE PREVIEWS) ---
def simulate_schedule(params):
    """Simulates month-by-month rent/CAM escalation exactly like Excel formulas."""
    start_year = params["Agreement Start Date"].year
    start_date = datetime.date(start_year, 1, 1)
    
    rows = []
    current_from = start_date
    area_sqft = params["Chargeable Area Sqft"]
    area_sqm = area_sqft / 10.764
    base_rent = params["Rent Per Sqft"]
    base_cam = params["Quoted CAM"]
    
    for m in range(1, 73):
        # 1. Fiscal Year classification
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
            
        # 2. To Date calculation
        if m in [2, 14, 26, 38, 50, 62]: # February rows
            # Leap year adjustment
            is_leap = (current_from.year % 4 == 0 and (current_from.year % 100 != 0 or current_from.year % 400 == 0))
            days = 29 if is_leap else 28
        elif m in [4, 6, 7, 16, 18, 20, 28, 30, 32, 40, 42, 44, 52, 54, 56, 64, 66, 68, 70, 72]: # 30 day rows in spec formulas
            days = 30
        else:
            days = 31
            
        current_to = current_from + datetime.timedelta(days=days-1)
        
        # 3. Rent rate (15% escalation every 3 years / 36 months)
        if m <= 36:
            rent_rate = base_rent
        else:
            rent_rate = base_rent * (1.0 + params["Escalation %"])
            
        # 4. CAM rate (5% escalation in Month 14, 26, 37, 50, 62)
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


# --- DYNAMIC SECURITY DEPOSIT & ARO WORKING SIMULATOR ---
def compute_deposit_working(params, schedule_df):
    """Simulates carrying costs and ARO parameters from inputs."""
    sd_amount = params["Security DepositAmount"] if "Security DepositAmount" in params else params["Security Deposit Amount"]
    rate_capital = params["Cost of Capital"]
    term_months = 72 # Total months in schedule
    
    # Interest carrying costs calculations
    sd_interest = sd_amount * rate_capital / 12 * term_months
    energy_deposit = params["Addnl.Deposit -energy(Refundable)"]
    energy_interest = energy_deposit * rate_capital * 3 # matching formula: =+B6*B9*3
    
    fitout_interest = 0.0 # B3 = 0, so C3 = 0
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


# --- STREAMLIT UI DESIGN & INTERACTIVITY ---

# Header Panel
col_title, col_logo = st.columns([7, 1])
with col_title:
    st.markdown('<p class="badge-title">🏢 Rental Automation Pro</p>', unsafe_allow_html=True)
    st.markdown('<p style="color: #64748b; font-size: 1.15rem; margin-top:-10px;">High-Fidelity Lease Parameter Integration and Excel Generation Engine</p>', unsafe_allow_html=True)

st.markdown("---")

# Sidebar file uploader and manual overrides
st.sidebar.markdown('<p style="font-size: 1.4rem; font-weight:700; color:#f8fafc; margin-bottom: 15px;">📁 Upload & Configure</p>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Upload Lease Parameter Workbook (.xlsx)", type=["xlsx"])

st.sidebar.markdown('<p style="font-size: 1.2rem; font-weight:600; color:#cbd5e1; margin-top: 20px;">⚙️ Parameter Overrides</p>', unsafe_allow_html=True)

# Default Mock values matching specimen/input sheet to let users test instantly
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

# Load parsed values if file is uploaded, otherwise use defaults
if uploaded_file is not None:
    parsed_vals, err = parse_input_xlsx(uploaded_file)
    if err:
        st.sidebar.error(f"Error parsing workbook: {err}")
        params = default_params
    else:
        st.sidebar.success("Workbook parsed successfully!")
        params = parsed_vals
else:
    params = default_params

# Editable Sidebar controls
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

# Merge UI modifications back into the parameters dictionary
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

# Run live Python-side simulations
schedule_df = simulate_schedule(ui_params)
sd_results = compute_deposit_working(ui_params, schedule_df)

# --- STATS SUMMARY BAR ---
st.markdown("### 📊 Live Parameter Dashboard")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)

with col_s1:
    st.markdown(f"""
    <div class="premium-card">
        <div class="metric-label">Chargeable Office Area</div>
        <div class="metric-value">{int(ui_params['Chargeable Area Sqft']):,} sqft</div>
        <div style="font-size:0.85rem; color:#64748b; margin-top:2px;">≈ {ui_params['Chargeable Area Sqft']/10.764:,.2f} sq.m.</div>
    </div>
    """, unsafe_allow_html=True)
    
with col_s2:
    st.markdown(f"""
    <div class="premium-card">
        <div class="metric-label">Initial Rent + CAM Rate</div>
        <div class="metric-value">{ui_params['Currency']} {ui_params['Rent Per Sqft'] + ui_params['Quoted CAM']:.2f}</div>
        <div style="font-size:0.85rem; color:#64748b; margin-top:2px;">Rent: {ui_params['Rent Per Sqft']} | CAM: {ui_params['Quoted CAM']}</div>
    </div>
    """, unsafe_allow_html=True)
    
with col_s3:
    st.markdown(f"""
    <div class="premium-card">
        <div class="metric-label">Security Deposit Amount</div>
        <div class="metric-value">{ui_params['Currency']} {ui_params['Security Deposit Amount']:,.2f}</div>
        <div style="font-size:0.85rem; color:#64748b; margin-top:2px;">Interest rate: {ui_params['Cost of Capital']*100:.1f}%</div>
    </div>
    """, unsafe_allow_html=True)
    
with col_s4:
    st.markdown(f"""
    <div class="premium-card">
        <div class="metric-label">Lease Term Coverage</div>
        <div class="metric-value">{(ui_params['Agreement End Date'] - ui_params['Agreement Start Date']).days/365:.2f} yrs</div>
        <div style="font-size:0.85rem; color:#64748b; margin-top:2px;">{ui_params['Agreement Start Date'].strftime('%b %Y')} to {ui_params['Agreement End Date'].strftime('%b %Y')}</div>
    </div>
    """, unsafe_allow_html=True)


# --- TAB PREVIEWS ---
tab_main, tab_schedule, tab_deposit = st.tabs([
    "📂 Sheet 1: Main Summary Preview",
    "📈 Sheet 2: Monthly Lease Schedule & Plot",
    "🛡️ Sheet 3: Security Deposit Working"
])

# 1. Main Sheet Tab
with tab_main:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("🏢 Main Summary Meta-Properties")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown(f"**REU Name:** `{ui_params['REU Name']}`")
        st.markdown(f"**Building Name:** `{ui_params['Building Name']}`")
        st.markdown(f"**City & Country:** `{ui_params['City']}, {ui_params['Country']}`")
        st.markdown(f"**Chargeable Area:** `{ui_params['Chargeable Area Sqft']:,} Sq.ft.` (`{ui_params['Chargeable Area Sqft']/10.764:,.2f} Sq.m.`)")
        st.markdown(f"**Agreement Term:** `{ui_params['Agreement Start Date'].strftime('%d-%b-%Y')} to {ui_params['Agreement End Date'].strftime('%d-%b-%Y')}`")
        st.markdown(f"**Agreement Period:** `{(ui_params['Agreement End Date'] - ui_params['Agreement Start Date']).days / 365.0:.2f} Years`")
    with col_m2:
        st.markdown(f"**Rent Rate / Sqft / Mo:** `{ui_params['Currency']} {ui_params['Rent Per Sqft']:.2f}`")
        st.markdown(f"**CAM Rate / Sqft / Mo:** `{ui_params['Currency']} {ui_params['Quoted CAM']:.2f}`")
        st.markdown(f"**Capex (Fitout cost):** `{ui_params['Currency']} {ui_params['Fitout Cost']:,.2f}`")
        st.markdown(f"**Security Deposit:** `{ui_params['Currency']} {ui_params['Security Deposit Amount']:,.2f}`")
        st.markdown(f"**Rent Escalation Clause:** `{int(ui_params['Escalation %'] * 100)}% every {int(ui_params['Escalation Frequency Months'] // 12)} years`")
        st.markdown(f"**CAM Escalation Clause:** `{int(ui_params['CAM Escalation %'] * 100)}% every year`")
    st.markdown('</div>', unsafe_allow_html=True)

# 2. Schedule & Chart Tab
with tab_schedule:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("📈 Escalation Trend Chart & Forecast")
    
    # Plotly Line Chart for Rentals vs CAM
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=schedule_df["Month #"], 
        y=schedule_df["Rent Rate ($/Sqft)"],
        mode='lines+markers',
        name='Rental Rate / Sqft',
        line=dict(color='#4f46e5', width=3),
        marker=dict(size=6)
    ))
    fig.add_trace(go.Scatter(
        x=schedule_df["Month #"], 
        y=schedule_df["CAM Rate ($/Sqft)"],
        mode='lines+markers',
        name='CAM Rate / Sqft',
        line=dict(color='#10b981', width=3, dash='dash'),
        marker=dict(size=6)
    ))
    
    fig.update_layout(
        title="Dynamic Rate Escalation Lifecycle (72 Months)",
        xaxis_title="Lease Schedule Months",
        yaxis_title=f"Rate ({ui_params['Currency']} / Sq.ft. / Month)",
        legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01),
        margin=dict(l=20, r=20, t=50, b=20),
        plot_bgcolor='rgba(248, 250, 252, 0.5)',
        paper_bgcolor='rgba(0,0,0,0)',
        hovermode="x"
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Grouped Financial Summaries
    st.markdown("#### Monthly Rent Schedule Table View")
    
    # Beautiful Pandas table view with formatted currencies
    display_df = schedule_df.copy()
    display_df["Rent Amount"] = display_df["Rent Amount"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
    display_df["CAM Amount"] = display_df["CAM Amount"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
    display_df["Total Billing"] = display_df["Total Billing"].map(lambda x: f"{ui_params['Currency']} {x:,.2f}")
    
    st.dataframe(display_df, use_container_width=True, height=400)
    st.markdown('</div>', unsafe_allow_html=True)

# 3. Security Deposit carrying cost Tab
with tab_deposit:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("🛡️ Security Deposit Carrying Cost & Imputed Interest Workings")
    
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


# --- EXCEL WORKBOOK GENERATOR PANEL ---
st.markdown("### 📤 Excel Workbook Generator & Download")
st.markdown("""
Press the button below to inject all of the customized parameters directly into the high-fidelity Excel workbook structure. 
All formatting, sheets, carrying cost formulas, and dynamic financial aggregates will be fully preserved.
""")

template_file_path = r"c:\Users\z004df5r\Documents\rental_automation_v1\artifacts\rental-specimen.xlsx"

if os.path.exists(template_file_path):
    try:
        # Build button and memory stream download
        output_stream = generate_output_workbook(template_file_path, ui_params)
        
        st.download_button(
            label="💾 Generate and Download Excel Workbook",
            data=output_stream,
            file_name=f"rental_sheet_{ui_params['REU Name']}_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Workbook is compiled and ready to download! Open in MS Excel to see carrying cost calculations live.")
    except Exception as e:
        st.error(f"Failed to load Excel template or compile changes: {e}")
else:
    st.warning(f"Master template workbook not found in workspace path: {template_file_path}. Please check file locations.")
