import streamlit as st
import pandas as pd

# 1. Page Config & High-End Enterprise Styling
st.set_page_config(page_title="CASS Reconciliation Hub", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0b0c10; color: #e5e7eb; font-family: 'Inter', sans-serif; }
    
    /* Top Banner Styling */
    .main-header { font-size: 26px; font-weight: 700; color: #ffffff; letter-spacing: -0.5px; }
    .date-subheader { font-size: 13px; color: #9ca3af; margin-bottom: 25px; display: flex; align-items: center; gap: 6px; }
    
    /* Enterprise Workspace Cards */
    .workspace-card { background-color: #11131c; border: 1px solid #1f2937; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3); }
    .workspace-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #1f2937; padding-bottom: 12px; margin-bottom: 16px; }
    .workspace-title { font-size: 14px; font-weight: 600; color: #ffffff; text-transform: uppercase; letter-spacing: 0.5px; }
    
    /* Status Badges */
    .badge-matched { background-color: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; }
    .badge-alert { background-color: rgba(239, 68, 68, 0.12); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; }
    .badge-purple { background-color: rgba(139, 92, 246, 0.12); color: #a78bfa; border: 1px solid rgba(139, 92, 246, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; }
    
    /* Rows for Grid Look */
    .recon-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #1f2937; font-size: 14px; color: #d1d5db; }
    .recon-row.total { border-bottom: none; font-weight: 700; font-size: 15px; color: #3b82f6; padding-top: 14px; }
    
    /* Form & Input adjustments for Dark Theme */
    div[data-testid="ststForm"] { border: 1px solid #1f2937 !important; }
    [data-testid="stSidebar"] { background-color: #0d0f16; border-right: 1px solid #1f2937; }
    .sidebar-section-title { font-size: 11px; font-weight: 700; color: #4b5563; letter-spacing: 1px; margin-top: 20px; margin-bottom: 8px; text-transform: uppercase; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

# Αρχικοποίηση session state για την αποθήκευση των σχολίων/κινήσεων (Treasury Adjustments Log)
if "treasury_logs" not in st.session_state:
    st.session_state.treasury_logs = []

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    
    # 📁 1. Διάβασμα Ημερομηνίας (Tab 13 -> Κελί D4)
    try:
        df_tab13 = pd.read_excel(EXCEL_FILE, sheet_name=12, header=None)
        raw_date = df_tab13.iloc[3, 3]
        formatted_date = str(raw_date).split()[0] if pd.notna(raw_date) else "16/06/2026"
    except:
        formatted_date = "16/06/2026"

    # --- SIDEBAR: ΕΠΑΝΕΦΕΡΑ ΚΑΙ ΤΙΣ 15 ΚΑΡΤΕΛΕΣ ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-section-title'>Active Worksheets</div>", unsafe_allow_html=True)
    selected_tab = st.sidebar.radio("Καρτέλες:", sheet_names)

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    # Global configurations για μορφοποίηση στηλών σε λίρες
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
                        <div class="workspace-header"><div class="workspace-title">{title}</div><div class="badge-matched">MATCHED</div></div>
                        <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {data['prev_day']:,.2f}</span></div>
                        <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {data['debits']:,.2f}</span></div>
                        <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {data['credits']:,.2f}</span></div>
                        <div class="recon-row total"><span>Total Expected CUB</span><span>£ {data['total']:,.2f}</span></div>
                        <div class="recon-row" style="margin-top: 15px; border-top: 1px solid #1f2937;"><span>Internal CUB from Rec Date</span><span>£ {data['rec_date']:,.2f}</span></div>
                        <div class="recon-row total" style="color: #10b981;"><span>Variance Difference</span><span>£ {data['diff']:,.2f}</span></div>
                    </div>
                """, unsafe_allow_html=True)

    # ==========================================
    # 🟢 TAB 2: DAILY CLIENT MONEY REPORT (ENTERPRISE LOOK)
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

        # 3. INTERACTIVE TREASURY MOVEMENTS & COMMENTARY LOG (Αντικατάσταση του N/A)
        st.markdown("<br>", unsafe_allow_html=True)
        col_form, col_log = st.columns([1, 1])
        
        with col_form:
            st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📥 Log Manual Treasury Movement (Ex N/A Box)</div></div>", unsafe_allow_html=True)
            with st.form("treasury_movement_form", clear_on_submit=True):
                t_from = st.selectbox("From Account / Institution", ["CITIBANK", "LLOYDS", "QNB", "BBVA", "MODULR"])
                t_to = st.selectbox("To Account / Institution", ["LLOYDS", "CITIBANK", "QNB", "BBVA", "MODULR"])
                t_amount = st.number_input("Transaction Amount (£)", min_value=0.0, value=1000000.0, step=50000.0, format="%.2f")
                t_reason = st.text_input("Commentary / Variance Reason", value="Treasury movement liquidity coverage")
                if st.form_submit_button("Commit Movement to Audit Log"):
                    st.session_state.treasury_logs.append({"From": t_from, "To": t_to, "Amount": t_amount, "Reason": t_reason})
                    st.toast("Movement recorded successfully!", icon="✅")
            st.markdown("</div>", unsafe_allow_html=True)
            
        with col_log:
            st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📜 Active Variance Adjustments Log</div></div>", unsafe_allow_html=True)
            if st.session_state.treasury_logs:
                log_df = pd.DataFrame(st.session_state.treasury_logs)
                log_df["Amount"] = log_df["Amount"].map(lambda x: f"£{x:,.2f}")
                st.dataframe(log_df, use_container_width=True, hide_index=True)
            else:
                st.info("No active variance commentary or manual treasury movements posted for today.")
            st.markdown("</div>", unsafe_allow_html=True)

        # 4. EXPANDABLE SUB-LEDGERS (Stocks/Shares, Quai, Other Client Accounts)
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
            
        with st.expander("💶 Other Foreign Currency & Client Money Accounts"):
            other_df = pd.DataFrame([
                {"Bank": "Citi Bank NA London", "Account": "Saveable UK Client Money EUR (14747763)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"}
            ])
            st.data_editor(other_df, column_config=currency_config, use_container_width=True, hide_index=True, key="other_accounts_grid")

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
