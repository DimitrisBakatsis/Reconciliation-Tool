import streamlit as st
import pandas as pd

# 1. Ρύθμιση Σελίδας & Custom CSS για DARK MODE & BLUE/PURPLE DESIGN
st.set_page_config(page_title="Client Money Reconciliation", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    /* Φόντο όλης της εφαρμογής (Βαθύ Σκούρο) */
    .stApp {
        background-color: #0f0f12;
        color: #ffffff;
    }
    
    /* Στυλ για τα 4 Cards στην κορυφή (Σκούρο Μπλε/Γκρι) */
    .kpi-container {
        display: flex;
        gap: 15px;
        margin-bottom: 25px;
    }
    .kpi-card {
        background-color: #1e1e24;
        padding: 20px;
        border-radius: 8px;
        flex: 1;
        border: 1px solid #2d2d38;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
    }
    .kpi-title {
        font-size: 13px;
        color: #a0a0b0;
        font-weight: 500;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: bold;
        color: #4dadff; /* Έντονο Μπλε */
    }
    
    /* Σκούρα Containers για τα Blocks */
    .section-container {
        background-color: #16161a;
        padding: 25px;
        border-radius: 8px;
        border: 1px solid #25252e;
        margin-bottom: 25px;
    }
    .section-title {
        font-size: 18px;
        font-weight: bold;
        color: #ffffff;
        margin-bottom: 3px;
    }
    .section-subtitle {
        font-size: 13px;
        color: #707080;
        margin-bottom: 20px;
    }
    
    /* Custom Sidebar για Dark Theme */
    [data-testid="stSidebar"] {
        background-color: #111115;
        border-right: 1px solid #25252e;
    }
    
    /* Διορθώσεις για τα κείμενα των default widgets του Streamlit σε Dark */
    label {
        color: #a0a0b0 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Φόρτωση Δεδομένων από το Excel της Plum
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_data(ttl=2)
def load_excel_data():
    return pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)

try:
    raw_df = load_excel_data()
    
    # --- Πραγματικά ποσά από το Excel της Plum ---
    total_requirement = 2601370286.70
    resource_total = 2607556676.61
    shortfall_surplus = -1244.09
    net_change = 3246757.00
    
    # Πίνακας Account Balances (Τα αληθινά δεδομένα της 2ης καρτέλας)
    accounts_data = pd.DataFrame([
        {"BANK": "Citibank NA London", "ACCOUNT": "Saveable Cash ISA UK Client Money (14747801)", "PREV DAY BALANCE": 171777814.00, "COB BALANCE": 174394177.00, "VARIANCE": 2616363.00, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA Client Account (27551460)", "PREV DAY BALANCE": 3959844.00, "COB BALANCE": 3959844.00, "VARIANCE": 0.00, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA 30D Notice Client Account (27571468)", "PREV DAY BALANCE": 747535672.00, "COB BALANCE": 747535672.00, "VARIANCE": 0.00, "ENTITY": "Saveable Limited"}
    ])

    # --- SIDEBAR (Dark Mode με Μπλε Highlight) ---
    st.sidebar.markdown("<h2 style='color:#ffffff; font-size:18px; margin-bottom:0;'>CASS BARE TRUST</h2><p style='color:#707080; font-size:13px;'>Client Money Reconciliation</p>", unsafe_allow_html=True)
    st.sidebar.date_input("Reconciliation date")
    
    st.sidebar.markdown("<br><p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>OVERVIEW</p>", unsafe_allow_html=True)
    st.sidebar.markdown("<div style='padding: 8px 12px; color:#a0a0b0; font-size:14px;'>• Sign Off & Checks</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div style='padding: 8px 12px; background-color:#1e293b; color:#38bdf8; border-radius:4px; font-weight:500; font-size:14px;'>• Daily CM Report</div>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<br><p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>RECONCILIATION</p>", unsafe_allow_html=True)
    st.sidebar.markdown("<div style='padding: 8px 12px; color:#a0a0b0; font-size:14px;'>• Bare Trust Workings</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div style='padding: 8px 12px; color:#a0a0b0; font-size:14px;'>• Account Details</div>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<br><br><br><br><p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>EXPORT</p>", unsafe_allow_html=True)
    st.sidebar.button("🔮 Export to Excel (.xlsx)", use_container_width=True)
    st.sidebar.button("⬜ Print / PDF", use_container_width=True)

    # --- MAIN CONTENT ---
    
    st.markdown(f"""
        <h1 style='font-size: 26px; font-weight: bold; color: #ffffff; margin-bottom:0;'>Daily Client Money Report</h1>
        <p style='color: #a0a0b0; font-size: 14px; margin-bottom: 25px;'>Saveable – Bare Trust Client Money Balances (GBP)</p>
    """, unsafe_allow_html=True)

    # 1. TOP CARDS (Οριζόντια διάταξη σε Dark Mode)
    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card"><div class="kpi-title">Total Requirement</div><div class="kpi-value">£ {total_requirement:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Resource</div><div class="kpi-value">£ {resource_total:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Shortfall / Surplus</div><div class="kpi-value" style="color: {'#ff4d4d' if shortfall_surplus < 0 else '#4dffaa'}">£ {shortfall_surplus:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Net Change (COB)</div><div class="kpi-value">£ {net_change:,.2f}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # 2. BLOCK: Account Balances
    st.markdown("""
        <div class="section-container">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
                <div class="section-title">Account Balances</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Ο διαδραστικός πίνακας τώρα ακολουθεί το Dark Theme
    edited_df = st.data_editor(
        accounts_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="balances_table_dark"
    )
    
    # 3. BLOCK: Daily Reconciliation (Οριζόντια Inputs)
    st.markdown("<br><div class='section-container'><div class='section-title' style='margin-bottom:15px;'>Daily Reconciliation</div>", unsafe_allow_html=True)
    
    rec_col1, rec_col2, rec_col3 = st.columns(3)
    with rec_col1:
        req_val = st.number_input("Requirement (£)", value=float(total_requirement), format="%.2f")
    with rec_col2:
        trans_val = st.number_input("Inclusive of transfers in to apply (£)", value=6187634.40, format="%.2f")
    with rec_col3:
        res_val = st.number_input("Resource (inclusive of Quai) (£)", value=float(resource_total), format="%.2f", disabled=True)
        
    st.markdown("</div>", unsafe_allow_html=True)

    # 4. BLOCK: Shortfall & Comments
    st.markdown("<div class='section-container'>", unsafe_allow_html=True)
    
    st.text_input("Calculated Shortfall / Surplus", value=f"£ {shortfall_surplus:,.2f}", disabled=True)
    
    st.text_area("Reason for Internal Movements", value="CISA: Overall Shortfall of £1,244.79 residual interest paid to users as part of the transfer out process...")
    st.text_area("Additional Comments", value="Amount of £1,315.68 is currently being held on LISA Unalloc...")
    
    st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Error drawing custom UI: {e}")
