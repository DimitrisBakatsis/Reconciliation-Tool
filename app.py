import streamlit as st
import pandas as pd

# 1. Page Config & Premium UI Styling (Pixel-Perfect για image_89ea59.png)
st.set_page_config(page_title="CASS Reconciliation Hub", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0b0c10; color: #e5e7eb; font-family: 'Inter', sans-serif; }
    
    /* Top Banner Styling */
    .main-header { font-size: 26px; font-weight: 700; color: #ffffff; letter-spacing: -0.5px; }
    .date-subheader { font-size: 13px; color: #9ca3af; margin-bottom: 25px; display: flex; align-items: center; gap: 6px; }
    
    /* Enterprise Workspace Cards (image_89ea59.png look) */
    .workspace-card { 
        background-color: #0d0f16; 
        border: 1px solid #1f2937; 
        border-radius: 8px; 
        padding: 24px; 
        margin-bottom: 20px; 
    }
    .workspace-header { 
        border-bottom: 1px solid #1f2937; 
        padding-bottom: 14px; 
        margin-bottom: 20px; 
    }
    .workspace-title { 
        font-size: 13px; 
        font-weight: 700; 
        color: #a78bfa; /* Μωβ/Μπλε τίτλος όπως στη φωτογραφία */
        text-transform: uppercase; 
        letter-spacing: 0.5px; 
    }
    
    /* Status Badges */
    .badge-matched { background-color: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; }
    .badge-alert { background-color: rgba(239, 68, 68, 0.12); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; }
    
    /* Rows for Grid Look */
    .recon-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #1f2937; font-size: 14px; color: #d1d5db; }
    .recon-row.total { border-bottom: none; font-weight: 700; font-size: 15px; color: #3b82f6; padding-top: 14px; }
    
    /* Custom Input Styles για να ταιριάζουν με το image_89ea59.png */
    div[data-testid="stSelectbox"] > label, div[data-testid="stTextArea"] > label {
        color: #ffffff !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        margin-bottom: 8px !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"], div[data-testid="stTextArea"] textarea {
        background-color: #1a1c24 !important;
        border: 1px solid #2d3748 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
    }
    
    /* Custom Buttons */
    .stButton > button {
        background-color: #1a1c24 !important;
        color: #ffffff !important;
        border: 1px solid #4a5568 !important;
        border-radius: 6px !important;
        padding: 6px 16px !important;
        font-size: 13px !important;
    }
    
    /* Sidebar Navigation Customization */
    [data-testid="stSidebar"] { background-color: #0d0f16; border-right: 1px solid #1f2937; }
    .sidebar-section-title { font-size: 11px; font-weight: 700; color: #4b5563; letter-spacing: 1px; margin-top: 20px; margin-bottom: 8px; text-transform: uppercase; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

# Αρχικοποίηση session states για τα logs
if "cisa_comments" not in st.session_state:
    st.session_state.cisa_comments = []
if "lisa_comments" not in st.session_state:
    st.session_state.lisa_comments = []

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    
    # 📁 Διάβασμα Ημερομηνίας (Tab 13 -> Κελί D4)
    try:
        df_tab13 = pd.read_excel(EXCEL_FILE, sheet_name=12, header=None)
        raw_date = df_tab13.iloc[3, 3]
        formatted_date = str(raw_date).split()[0] if pd.notna(raw_date) else "16/06/2026"
    except:
        formatted_date = "16/06/2026"

    # --- SIDEBAR: ΟΛΕΣ ΟΙ 15 ΚΑΡΤΕΛΕΣ ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-section-title'>Active Worksheets</div>", unsafe_allow_html=True)
    selected_tab = st.sidebar.radio("Καρτέλες:", sheet_names)

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
    }

    # ==========================================
    # 🟢 TAB 1: COMBINED USER BALANCE CHECK
    # ==========================================
    if selected_tab == sheet_names[0]:
        st.markdown("### Ledger vs CUB Verification Workspace")
        cisa_data = {"prev_day": 2386124297.55, "debits": 10417421.49, "credits": 11826163.22, "total": 2387533039.28, "rec_date": 2387533039.28, "diff": 0.00}
        lisa_data = {"prev_day": 217714664.80, "debits": 251643.28, "credits": 946498.01, "total": 218409519.53, "rec_date": 218409519.53, "diff": 0.00}

        col_left, col_right = st.columns(2)
        for col, title, data in zip([col_left, col_right], ["CISA Check", "LISA Check"], [cisa_data, lisa_data]):
            with col:
                st.markdown(f"""
                    <div class="workspace-card">
                        <div class="workspace-header"><div class="workspace-title" style="color:#ffffff;">{title}</div><div class="badge-matched">MATCHED</div></div>
                        <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {data['prev_day']:,.2f}</span></div>
                        <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {data['debits']:,.2f}</span></div>
                        <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {data['credits']:,.2f}</span></div>
                        <div class="recon-row total"><span>Total Expected CUB</span><span>£ {data['total']:,.2f}</span></div>
                        <div class="recon-row" style="margin-top: 15px; border-top: 1px solid #1f2937;"><span>Internal CUB from Rec Date</span><span>£ {data['rec_date']:,.2f}</span></div>
                        <div class="recon-row total" style="color: #10b981;"><span>Variance Difference</span><span>£ {data['diff']:,.2f}</span></div>
                    </div>
                """, unsafe_allow_html=True)

    # ==========================================
    # 🟢 TAB 2: DAILY CLIENT MONEY REPORT
    # ==========================================
    elif selected_tab == sheet_names[1]:
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        
        # 1. CASH ISA MAIN LEDGER
        st.markdown("""
            <div class="workspace-card" style="padding-bottom: 5px; margin-bottom: 5px;">
                <div class="workspace-header">
                    <div class="workspace-title" style="color: #a78bfa;">Cash ISA Client Money Balances</div>
                    <div class="badge-alert">CISA Net Change: -£971,704.00</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": 176992857.0, "COB Balance": 176021153.0, "Variance": -971704.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": 3959844.0, "COB Balance": 3959844.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": 747535672.0, "COB Balance": 747535672.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": 1168000000.0, "COB Balance": 1168000000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        # 2. LIFETIME ISA MAIN LEDGER
        st.markdown("""
            <div class="workspace-card" style="padding-bottom: 5px; margin-bottom: 5px; margin-top: 25px;">
                <div class="workspace-header">
                    <div class="workspace-title" style="color: #a78bfa;">Lifetime ISA Client Money Balances</div>
                    <div class="badge-matched">LISA Net Change: +£610,408.97</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        lisa_df = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": 38363170.0, "COB Balance": 38973579.0, "Variance": 610408.97, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": 714980.0, "COB Balance": 714980.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": 79300000.0, "COB Balance": 79300000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": 100000000.0, "COB Balance": 100000000.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_grid")

        # 3. AUDITING & COMMENTARY SUITE (ΑΚΡΙΒΕΣ LAYOUT IMAGE_89ea59.png)
        st.markdown("<br>", unsafe_allow_html=True)
        col_cisa, col_lisa = st.columns(2)
        
        # --- CASH ISA COLUMN ---
        with col_cisa:
            st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>CASH ISA - COMMENTARY ON VARIANCES</div></div>", unsafe_allow_html=True)
            cisa_account = st.selectbox("Select Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"], key="cisa_sel")
            cisa_text = st.text_area("Variance Explanation / Treasury Action", placeholder="Type manual movement or commentary here...", key="cisa_text_area", height=120)
            if st.button("Post CISA Commentary", key="btn_cisa"):
                if cisa_text:
                    st.session_state.cisa_comments.append({"Account": cisa_account, "Commentary": cisa_text})
                    st.toast("CASS CISA log updated!", icon="💜")
            
            if st.session_state.cisa_comments:
                st.markdown("<p style='font-size:12px; color:#8a8f98; margin-top:15px;'>Recorded Audit Logs:</p>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(st.session_state.cisa_comments), use_container_width=True, hide_index=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # --- LIFETIME ISA COLUMN ---
        with col_lisa:
            st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>LIFETIME ISA - COMMENTARY ON VARIANCES</div></div>", unsafe_allow_html=True)
            lisa_account = st.selectbox("Select Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"], key="lisa_sel")
            lisa_text = st.text_area("Variance Explanation / Treasury Action", placeholder="Type manual movement or commentary here...", key="lisa_text_area", height=120)
            if st.button("Post LISA Commentary", key="btn_lisa"):
                if lisa_text:
                    st.session_state.lisa_comments.append({"Account": lisa_account, "Commentary": lisa_text})
                    st.toast("CASS LISA log updated!", icon="💜")
                        
            if st.session_state.lisa_comments:
                st.markdown("<p style='font-size:12px; color:#8a8f98; margin-top:15px;'>Recorded Audit Logs:</p>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(st.session_state.lisa_comments), use_container_width=True, hide_index=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # 4. EXPANDABLE SUB-LEDGERS
        st.markdown("### 🔍 Secondary Portfolios & Trust Breakdowns")
        
        with st.expander("📊 Stocks / Shares ISA Ledger Breakdown"):
            stocks_df = pd.DataFrame([
                {"Bank": "Barclays UK PLC", "Account": "SAVEABLE LTD (90314552) - Pending Sells/Buys", "Previous Day Balance": 1912753.33, "COB Balance": 1413133.97, "Variance": -499619.0, "Performed By": "Quai - Cash Held"},
                {"Bank": "Winterfloods", "Account": "SAVEABLE LTD", "Previous Day Balance": 102290541.37, "COB Balance": 102342064.32, "Variance": 51522.95, "Performed By": "Quai - Units Held"},
                {"Bank": "Winterfloods", "Account": "SAVEABLE LTD Total Asset Line", "Previous Day Balance": 366903547.62, "COB Balance": 370473068.35, "Variance": 3569520.73, "Performed By": "Quai - Cash Held"}
            ])
            st.data_editor(stocks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="stocks_grid")
            
        with st.expander("🏛️ Quai Resource & Requirement Breakdown"):
            quai_data = {"Requirement": 3532196.96, "Resource": 3532197.14, "Shortfall / Surplus": 0.18}
            st.markdown(f"""
                <div class="recon-row"><span>Quai Total Platform Requirement</span><span>£ {quai_data['Requirement']:,.2f}</span></div>
                <div class="recon-row"><span>Quai Segregated Pool Resource</span><span>£ {quai_data['Resource']:,.2f}</span></div>
                <div class="recon-row total" style="color: #10b981;"><span>Calculated Residual Surplus</span><span>£ {quai_data['Shortfall / Surplus']:,.2f}</span></div>
            """, unsafe_allow_html=True)

        # 5. DAILY RECONCILIATION CALCULATION BLOCK
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>⚙️ CASS Corporate Daily Calculation Suite</div></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.number_input("Total Pure Requirement (£)", value=2601370286.70, format="%.2f", disabled=True)
        with col2:
            st.number_input("Inclusive of Transfers In to Apply (£)", value=6187634.40, format="%.2f", disabled=True)
        with col3:
            st.number_input("Total Eligible Resource Pool (£)", value=2607556676.61, format="%.2f", disabled=True)
        
        st.text_input("Net Shortfall / Surplus Result", value="-£1,244.09 (Action Required: Cover via Corporate Account)", disabled=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================
    # 🟢 ANY OTHER TAB (AUTOMATIC POWER-VIEW)
    # ==========================================
    else:
        st.markdown(f"### 📂 Sub-Ledger Archive View: {selected_tab}")
        df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
        st.dataframe(df_any.dropna(how='all').reset_index(drop=True), use_container_width=True)

except Exception as e:
    st.error(f"System Error: {e}")
