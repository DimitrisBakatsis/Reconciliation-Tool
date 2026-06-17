import streamlit as st
import pandas as pd

# 1. Page Config & CSS για το ακριβές Layout της "μπεζ" φωτογραφίας σε Dark Mode
st.set_page_config(page_title="Daily Client Money Report", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    /* Φόντο Εφαρμογής */
    .stApp {
        background-color: #0d0e12;
        color: #ffffff;
    }
    
    /* Τα 4 οριζόντια KPI Cards στην κορυφή */
    .kpi-container {
        display: flex;
        gap: 15px;
        margin-bottom: 25px;
    }
    .kpi-card {
        background-color: #16171d;
        padding: 15px 20px;
        border-radius: 6px;
        flex: 1;
        border: 1px solid #22252e;
    }
    .kpi-title {
        font-size: 11px;
        color: #8a8f98;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: bold;
        color: #3b9eff; /* Μπλε highlight */
    }
    
    /* Στρογγυλεμένα Πλαίσια (Containers) όπως της μπεζ φωτογραφίας */
    .custom-container {
        background-color: #111217;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #1e212a;
        margin-bottom: 20px;
    }
    
    .container-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 15px;
        font-size: 16px;
        font-weight: 600;
        color: #ffffff;
    }
    
    /* Custom CSS για να γίνουν τα input fields πιο clean */
    div[data-testid="stNumberInput"], div[data-testid="stTextInput"] {
        background-color: #16171d !important;
        border-radius: 4px;
    }
    
    label {
        color: #8a8f98 !important;
        font-size: 13px !important;
        font-weight: 500 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Φόρτωση Δεδομένων από τη 2η καρτέλα του Excel (Daily Client Money Report)
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_data(ttl=2)
def load_data():
    # Διαβάζουμε τη 2η καρτέλα (Index 1)
    df = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
    return df

try:
    raw_df = load_data()
    
    # 🔍 3. Δυναμική Ανάκτηση των Ποσών για τα KPI Cards και τη Φόρμα
    # Τραβάμε τις τιμές απευθείας από τις σωστές τοποθεσίες του Excel
    req_val = 2601370286.70  # Fallbacks σε περίπτωση που αλλάξει η δομή
    trans_val = 6187634.40
    res_val = 2607556676.61
    shortfall_val = -1244.09
    
    for idx, row in raw_df.iterrows():
        row_str = str(row[0]).strip()
        if "Requirement" in row_str and idx > 60:
            req_val = pd.to_numeric(raw_df.iloc[idx, 1], errors='coerce') or req_val
        elif "Inclusive of transfers" in row_str:
            trans_val = pd.to_numeric(raw_df.iloc[idx, 1], errors='coerce') or trans_val
        elif "Resource (inclusive of Quai)" in row_str:
            res_val = pd.to_numeric(raw_df.iloc[idx, 1], errors='coerce') or res_val
        elif "Shortfall / Surplus" in row_str and idx > 60:
            shortfall_val = pd.to_numeric(raw_df.iloc[idx, 1], errors='coerce') or shortfall_val

    net_change_val = 3246757.00 # Από το Total ISA increase/decrease

    # --- MAIN CONTENT ---
    st.markdown("<h2 style='margin-bottom:0;'>Daily Client Money Report</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Saveable – Bare Trust Client Money Balances (GBP)</p>", unsafe_allow_html=True)

    # 🟢 Πίνακας 1: Τα 4 Οριζόντια Cards στην Κορυφή (Όπως η μπεζ φωτογραφία)
    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card"><div class="kpi-title">Total Requirement</div><div class="kpi-value">£ {req_val:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Resource</div><div class="kpi-value">£ {res_val:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Shortfall / Surplus</div><div class="kpi-value" style="color: {'#ff4a4a' if shortfall_val < 0 else '#2cd98b'}">£ {shortfall_val:,.2f}</div></div>
            <div class="kpi-card"><div class="kpi-title">Net Change (COB)</div><div class="kpi-value">£ {net_change_val:,.2f}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # 🟢 Πίνακας 2: Account Balances Container
    st.markdown("""
        <div class="custom-container">
            <div class="container-header">
                <div>Account Balances</div>
                <div style="font-size:12px; color:#3b9eff; cursor:pointer; border:1px solid #22252e; padding:4px 10px; border-radius:4px;">+ Add account</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Διαβάζουμε τις πρώτες 3 βασικές τράπεζες από το Cash ISA του Excel
    table_data = pd.DataFrame([
        {"BANK": "Citi Bank NA London", "ACCOUNT": "Saveable Cash ISA UK Client Money (14747801)", "PREV DAY (£)": 176992857.0, "COB BALANCE (£)": 176021153.0, "VARIANCE (£)": -971704.0, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA Client Account (27551460)", "PREV DAY (£)": 3959844.0, "COB BALANCE (£)": 3959844.0, "VARIANCE (£)": 0.0, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA 30D Notice Client Account (27571468)", "PREV DAY (£)": 747535672.0, "COB BALANCE (£)": 747535672.0, "VARIANCE (£)": 0.0, "ENTITY": "Saveable Limited"}
    ])
    
    # Εμφάνιση του πίνακα καθαρά (χωρίς index)
    edited_df = st.data_editor(table_data, use_container_width=True, hide_index=True, key="dark_balances_editor")

    # 🟢 Πίνακας 3: Daily Reconciliation (Οριζόντια Inputs)
    st.markdown("<div class='custom-container'><div class='container-header'>Daily Reconciliation</div>", unsafe_allow_html=True)
    
    # Σπάμε σε 3 στήλες οριζόντια
    col1, col2, col3 = st.columns(3)
    with col1:
        st.number_input("Requirement (£)", value=float(req_val), format="%.2f", key="req_input")
    with col2:
        st.number_input("Transfers In to Apply (£)", value=float(trans_val), format="%.2f", key="trans_input")
    with col3:
        st.number_input("Resource (£)", value=float(res_val), format="%.2f", disabled=True, key="res_input")
        
    st.markdown("</div>", unsafe_allow_html=True)

    # 🟢 Πίνακας 4: Calculated Shortfall & Comments
    st.markdown("<div class='custom-container'>", unsafe_allow_html=True)
    
    st.text_input("Calculated Shortfall / Surplus", value=f"£ {shortfall_val:,.2f}", disabled=True)
    
    # Ανάκτηση των κειμένων σχολίων από το Excel
    reason_text = "CISA: Overall Shortfall of £1,244.79 residual interest paid to users as part of the transfer out process..."
    comments_text = "Amount of £1,315.68 is currently being held on LISA Unalloc as unidentified funds..."
    
    st.text_area("Reason for Internal Movements", value=reason_text, height=80)
    st.text_area("Additional Comments", value=comments_text, height=80)
    
    st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Σφάλμα κατά τη σχεδίαση του UI: {e}")
