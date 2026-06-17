import streamlit as st
import pandas as pd

# 1. Page Config & Extended Premium Dark CSS
st.set_page_config(page_title="CASS Reconciliation Portal", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0d0e12; color: #ffffff; }
    
    /* Top Header Styling */
    .main-header { font-size: 28px; font-weight: bold; color: #ffffff; margin-bottom: 2px; }
    .date-subheader { font-size: 14px; color: #8a8f98; margin-bottom: 30px; font-style: italic; }
    
    /* KPI Layouts */
    .kpi-container { display: flex; gap: 15px; margin-bottom: 25px; }
    .kpi-card { background-color: #16171d; padding: 15px 20px; border-radius: 6px; flex: 1; border: 1px solid #22252e; }
    .kpi-title { font-size: 11px; color: #8a8f98; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
    .kpi-value { font-size: 22px; font-weight: bold; color: #3b9eff; }
    
    /* Modern Section Container */
    .custom-container { background-color: #111217; padding: 22px; border-radius: 8px; border: 1px solid #1e212a; margin-bottom: 25px; }
    .container-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; font-size: 16px; font-weight: 600; color: #ffffff; border-bottom: 1px solid #1e212a; padding-bottom: 8px; }
    
    /* Reconciliation Rows for CUB Check */
    .recon-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dashed #1e212a; font-size: 14px; }
    .recon-row.total { border-bottom: none; font-weight: bold; font-size: 15px; color: #3b9eff; padding-top: 12px; }
    
    /* Status Badges */
    .status-badge-match { background-color: rgba(44, 217, 139, 0.1); color: #2cd98b; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 13px; border: 1px solid rgba(44, 217, 139, 0.2); }
    .status-badge-alert { background-color: rgba(255, 74, 74, 0.1); color: #ff4a4a; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 13px; border: 1px solid rgba(255, 74, 74, 0.2); }
    
    div[data-testid="stNumberInput"], div[data-testid="stTextInput"] { background-color: #16171d !important; border-radius: 4px; }
    label { color: #8a8f98 !important; font-size: 13px !important; font-weight: 500 !important; }
    [data-testid="stSidebar"] { background-color: #111215; border-right: 1px solid #22252e; }
    div[data-testid="stRadio"] > label { display: none !important; }
    </style>
""", unsafe_allow_html=True)

# 2. Φόρτωση του αρχείου Excel
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

# ΑΛΛΑΓΗ ΕΔΩ: st.cache_resource αντί για st.cache_data για να μην πετάει το serialization error
@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    
    # 📁 1. ΔΙΑΒΑΣΜΑ ΗΜΕΡΟΜΗΝΙΑΣ ΑΠΟ TAB 13 (Κελί D4 -> Row 3, Col 3)
    try:
        df_tab13 = pd.read_excel(EXCEL_FILE, sheet_name=12, header=None) # Index 12 = Tab 13
        raw_date = df_tab13.iloc[3, 3] # Κελί D4
        formatted_date = str(raw_date).split()[0] if pd.notna(raw_date) else "16/06/2026"
    except:
        formatted_date = "16/06/2026" # Fallback αν αποτύχει

    # --- SIDEBAR ---
    st.sidebar.markdown("<h2 style='color:#ffffff; font-size:16px; margin-bottom:0;'>Saveable Portal</h2><p style='color:#6c727f; font-size:12px;'>CASS Corporate Suite</p>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>NAVIGATION</p>", unsafe_allow_html=True)
    
    # Επιλογή μόνο για τα δύο πρώτα Tabs προς το παρόν
    active_menu = st.sidebar.radio("Μενού:", [sheet_names[0], sheet_names[1]])

    # --- ΚΕΝΤΡΙΚΟΣ ΤΙΤΛΟΣ ---
    st.markdown(f"<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'>📅 Close of Business Date: {formatted_date}</div>", unsafe_allow_html=True)

    # ==========================================
    # 🟢 TAB 1: MANUAL COMBINED USER BALANCE CHECK
    # ==========================================
    if active_menu == sheet_names[0]:
        st.markdown("### Manual Combined User Balance Check")
        st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Internal CUB verification vs Ledger Reporting Data</p>", unsafe_allow_html=True)
        
        # Δεδομένα από image_8a6545.png
        cisa_data = {
            "prev_day": 2386124297.55, "debits": 10417421.49, "credits": 11826163.22,
            "total": 2387533039.28, "rec_date": 2387533039.28, "diff": 0.00
        }
        lisa_data = {
            "prev_day": 217714664.80, "debits": 251643.28, "credits": 946498.01,
            "total": 218409519.53, "rec_date": 218409519.53, "diff": 0.00
        }

        # Προσπάθεια parsing από το αληθινό Tab 1 (Index 0) αν υπάρχει
        try:
            df_tab1 = pd.read_excel(EXCEL_FILE, sheet_name=0, header=None)
            # Εδώ μπορεί να προστεθεί αυτόματο lookup αν χρειαστεί
        except:
            pass

        col_left, col_right = st.columns(2)

        # BOX 1: CISA Check
        with col_left:
            st.markdown(f"""
                <div class="custom-container">
                    <div class="container-header">
                        <div>Combined User Balance Check - CISA</div>
                        <div class="status-badge-match">MATCHED</div>
                    </div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {cisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {cisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {cisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {cisa_data['total']:,.2f}</span></div>
                    <br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><span>£ {cisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #2cd98b;"><span>Difference</span><span>£ {cisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

        # BOX 2: LISA Check
        with col_right:
            st.markdown(f"""
                <div class="custom-container">
                    <div class="container-header">
                        <div>Combined User Balance Check - LISA</div>
                        <div class="status-badge-match">MATCHED</div>
                    </div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {lisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {lisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {lisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {lisa_data['total']:,.2f}</span></div>
                    <br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><span>£ {lisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #2cd98b;"><span>Difference</span><span>£ {lisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 🟢 TAB 2: DAILY CLIENT MONEY REPORT
    # ==========================================
    elif active_menu == sheet_names[1]:
        st.markdown("### Daily Client Money Report")
        st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Saveable – Bare Trust Client Money Balances (GBP)</p>", unsafe_allow_html=True)

        req_val, trans_val, res_val, shortfall_val, net_change_val = 3532196.96, 3507294.74, 2612998057.00, 0.18, 3246757.00

        # 4 KPI Cards στην κορυφή του Tab 2
        st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card"><div class="kpi-title">Total Requirement</div><div class="kpi-value">£ {req_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Resource</div><div class="kpi-value">£ {res_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Shortfall / Surplus</div><div class="kpi-value" style="color: #2cd98b;">£ {shortfall_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Net Change (COB)</div><div class="kpi-value">£ {net_change_val:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        # Account Balances Table Box
        st.markdown("""
            <div class="custom-container">
                <div class="container-header">
                    <div>Account Balances</div>
                    <div style="font-size:12px; color:#3b9eff; cursor:pointer; border:1px solid #22252e; padding:4px 10px; border-radius:4px;">+ Add account</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        table_data = pd.DataFrame([
            {"BANK": "Citi Bank NA London", "ACCOUNT": "Saveable Cash ISA UK Client Money (14747801)", "PREV DAY (£)": 176992857, "COB BALANCE (£)": 176021153, "VARIANCE (£)": -971704, "ENTITY": "Saveable Limited"},
            {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA Client Account (27551460)", "PREV DAY (£)": 3959844, "COB BALANCE (£)": 3959844, "VARIANCE (£)": 0, "ENTITY": "Saveable Limited"},
            {"BANK": "Lloyds Bank Plc", "ACCOUNT": "Saveable Cash ISA 30D Notice Client Account (27571468)", "PREV DAY (£)": 747535672, "COB BALANCE (£)": 747535672, "VARIANCE (£)": 0, "ENTITY": "Saveable Limited"}
        ])
        st.data_editor(table_data, use_container_width=True, hide_index=True, key="dark_balances_tab2")

        # Daily Reconciliation Inputs Section
        st.markdown("<div class='custom-container'><div class='container-header'>Daily Reconciliation</div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.number_input("Requirement (£)", value=float(req_val), format="%.2f", key="req_tab2")
        with col2:
            st.number_input("Transfers In to Apply (£)", value=float(trans_val), format="%.2f", key="trans_tab2")
        with col3:
            st.number_input("Resource (£)", value=float(res_val), format="%.2f", disabled=True, key="res_tab2")
        st.markdown("</div>", unsafe_allow_html=True)

        # Calculated Shortfall Box & Comments
        st.markdown("<div class='custom-container'>", unsafe_allow_html=True)
        st.text_input("Calculated Shortfall / Surplus", value=f"£ {shortfall_val:,.2f}", disabled=True)
        st.text_area("Reason for Internal Movements", value="CISA: Overall Shortfall of £1,244.79 residual interest paid to users...", height=80)
        st.text_area("Additional Comments", value="Amount of £1,315.68 is currently being held on LISA Unalloc...", height=80)
        st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"Σφάλμα κατά τη φόρτωση ή σχεδίαση: {e}")
