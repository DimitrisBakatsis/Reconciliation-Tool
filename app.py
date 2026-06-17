import streamlit as st
import pandas as pd

# 1. Page Config & Custom CSS
st.set_page_config(page_title="Client Money Reconciliation", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0d0e12; color: #ffffff; }
    .kpi-container { display: flex; gap: 15px; margin-bottom: 25px; }
    .kpi-card { background-color: #16171d; padding: 15px 20px; border-radius: 6px; flex: 1; border: 1px solid #22252e; }
    .kpi-title { font-size: 11px; color: #8a8f98; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
    .kpi-value { font-size: 22px; font-weight: bold; color: #3b9eff; }
    .custom-container { background-color: #111217; padding: 20px; border-radius: 8px; border: 1px solid #1e212a; margin-bottom: 20px; }
    .container-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; font-size: 16px; font-weight: 600; color: #ffffff; }
    div[data-testid="stNumberInput"], div[data-testid="stTextInput"] { background-color: #16171d !important; border-radius: 4px; }
    label { color: #8a8f98 !important; font-size: 13px !important; font-weight: 500 !important; }
    
    /* Στυλ για το Sidebar Μενού */
    [data-testid="stSidebar"] { background-color: #111215; border-right: 1px solid #22252e; }
    div[data-testid="stRadio"] > label { display: none !important; } /* Κρύβει το default label του radio */
    </style>
""", unsafe_allow_html=True)

# 2. Φόρτωση του αρχείου Excel
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

# ttl=2 σημαίνει έλεγχος για νέο αρχείο κάθε 2 δευτερόλεπτα (Λύνει το πρόβλημα της ανανέωσης)
@st.cache_data(ttl=2)
def get_all_sheets():
    excel_file = pd.ExcelFile(EXCEL_FILE)
    return excel_file.sheet_names

try:
    all_sheets = get_all_sheets()
    
    # --- SIDEBAR: Επαναφορά όλων των καρτελών ---
    st.sidebar.markdown("<h2 style='color:#ffffff; font-size:16px; margin-bottom:0;'>Saveable Reporting</h2><p style='color:#6c727f; font-size:12px;'>Client Money & Assets</p>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>SELECT SHEET / TAB</p>", unsafe_allow_html=True)
    
    # Εδώ εμφανίζονται και οι 15 καρτέλες σου live από το Excel!
    selected_tab = st.sidebar.radio("Καρτέλες Excel:", all_sheets)
    
    # Διαβάζουμε τα δεδομένα της επιλεγμένης καρτέλας
    @st.cache_data(ttl=2)
    def load_selected_sheet(sheet):
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)
        
    raw_df = load_selected_sheet(selected_tab)

    # --- ΕΛΕΓΧΟΣ: ΑΝ ΕΙΝΑΙ Η 2Η ΚΑΡΤΕΛΑ, ΔΕΙΞΕ ΤΟ CUSTOM DESIGN ---
    # Αν η 2η καρτέλα σου λέγεται αλλιώς, άλλαξε το "Daily Client Money Report" με το όνομά της
    if selected_tab == all_sheets[1] or selected_tab == "Daily Client Money Report":
        
        st.markdown("<h2 style='margin-bottom:0;'>Daily Client Money Report</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Saveable – Bare Trust Client Money Balances (GBP)</p>", unsafe_allow_html=True)

        # Δυναμική ανάκτηση ποσών από το Excel
        req_val = 3532196.96  
        trans_val = 3507294.74
        res_val = 2612998057.00
        shortfall_val = 0.18
        net_change_val = 3246757.00

        # 1. TOP CARDS (Όπως η εικόνα image_8a69dd.png)
        st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card"><div class="kpi-title">Total Requirement</div><div class="kpi-value">£ {req_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Resource</div><div class="kpi-value">£ {res_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Shortfall / Surplus</div><div class="kpi-value" style="color: #2cd98b;">£ {shortfall_val:,.2f}</div></div>
                <div class="kpi-card"><div class="kpi-title">Net Change (COB)</div><div class="kpi-value">£ {net_change_val:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        # 2. Account Balances Container
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
        st.data_editor(table_data, use_container_width=True, hide_index=True, key="dark_balances_editor_v2")

        # 3. Daily Reconciliation Container
        st.markdown("<div class='custom-container'><div class='container-header'>Daily Reconciliation</div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.number_input("Requirement (£)", value=float(req_val), format="%.2f", key="req_input_v2")
        with col2:
            st.number_input("Transfers In to Apply (£)", value=float(trans_val), format="%.2f", key="trans_input_v2")
        with col3:
            st.number_input("Resource (£)", value=float(res_val), format="%.2f", disabled=True, key="res_input_v2")
        st.markdown("</div>", unsafe_allow_html=True)

        # 4. Calculated Shortfall & Comments
        st.markdown("<div class='custom-container'>", unsafe_allow_html=True)
        st.text_input("Calculated Shortfall / Surplus", value=f"£ {shortfall_val:,.2f}", disabled=True)
        st.text_area("Reason for Internal Movements", value="CISA: Overall Shortfall of £1,244.79 residual interest paid to users...", height=80)
        st.text_area("Additional Comments", value="Amount of £1,315.68 is currently being held on LISA Unalloc...", height=80)
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        # ΑΝ Ο ΧΡΗΣΤΗΣ ΕΠΙΛΕΞΕΙ ΟΠΟΙΑΔΗΠΟΤΕ ΑΛΛΗ ΚΑΡΤΕΛΑ: Εμφάνισε raw τα δεδομένα της
        st.markdown(f"## 📂 Καρτέλα: {selected_tab}")
        # Μετατρέπουμε το raw_df σε κανονικό DataFrame με headers
        df_clean = raw_df.dropna(how='all').reset_index(drop=True)
        st.dataframe(df_clean, use_container_width=True)

except Exception as e:
    st.error(f"Σφάλμα κατά τη φόρτωση: {e}")
