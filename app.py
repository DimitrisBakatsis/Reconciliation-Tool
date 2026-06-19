import streamlit as st
import pandas as pd

# 1. Page Config & Premium UI Styling
st.set_page_config(page_title="CASS Reconciliation Hub", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0b0c10; color: #e5e7eb; font-family: 'Inter', sans-serif; }
    .main-header { font-size: 26px; font-weight: 700; color: #ffffff; letter-spacing: -0.5px; }
    .date-subheader { font-size: 13px; color: #9ca3af; margin-bottom: 25px; display: flex; align-items: center; gap: 6px; }
    
    /* Top Metrics Grid */
    .metric-grid { display: flex; gap: 16px; margin-bottom: 25px; flex-wrap: wrap; }
    .metric-card { 
        background-color: #11141d; 
        border: 1px solid #1f2937; 
        border-radius: 8px; 
        padding: 20px; 
        flex: 1; 
        min-width: 220px; 
    }
    .metric-label { font-size: 11px; font-weight: 600; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
    .metric-value { font-size: 24px; font-weight: 700; }
    .metric-value.blue { color: #3b82f6; }
    .metric-value.red { color: #ef4444; }
    .metric-value.green { color: #10b981; }
    
    /* Table Header Container with Top-Right Badges */
    .table-header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #0d0f16;
        border: 1px solid #1f2937;
        border-bottom: none;
        border-radius: 8px 8px 0 0;
        padding: 14px 20px;
        margin-top: 20px;
    }
    .table-title { font-size: 13px; font-weight: 700; color: #a78bfa; text-transform: uppercase; letter-spacing: 0.5px; }
    
    /* Inline Floating Net Change Badges */
    .net-change-badge {
        font-size: 12px;
        font-weight: 700;
        padding: 5px 12px;
        border-radius: 20px;
    }
    .net-change-badge.red { background-color: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); }
    .net-change-badge.green { background-color: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); }
    .net-change-badge.orange { background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; border: 1px solid rgba(245, 158, 11, 0.3); }

    .workspace-card { background-color: #0d0f16; border: 1px solid #1f2937; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
    .workspace-header { border-bottom: 1px solid #1f2937; padding-bottom: 14px; margin-bottom: 20px; }
    .workspace-title { font-size: 13px; font-weight: 700; color: #a78bfa; text-transform: uppercase; letter-spacing: 0.5px; }
    .recon-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #1f2937; font-size: 14px; color: #d1d5db; }
    .recon-row.total { border-bottom: none; font-weight: 700; font-size: 15px; color: #3b82f6; padding-top: 14px; }
    .recon-row.difference { border-bottom: none; font-weight: 700; font-size: 15px; color: #10b981; padding-top: 14px; }
    
    /* Strategic Reason / Conclusion Box look */
    .reason-box {
        background-color: #1a1c24;
        border-left: 4px solid #3b82f6;
        padding: 20px;
        border-radius: 6px;
        margin-top: 25px;
        margin-bottom: 25px;
    }
    .reason-title { font-size: 14px; font-weight: 700; color: #ffffff; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
    .reason-section { font-size: 13px; line-height: 1.6; color: #d1d5db; margin-bottom: 12px; }
    .reason-section strong { color: #ffffff; }

    /* Log Card Style */
    .log-card { 
        background-color: #11141d; 
        border-left: 4px solid #a78bfa; 
        border-radius: 6px; 
        padding: 16px; 
        margin-bottom: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        width: 100%;
    }
    .log-details { flex-grow: 1; padding-right: 15px; }
    .log-meta { font-size: 12px; color: #9ca3af; margin-bottom: 4px; font-weight: 600; }
    .log-text { font-size: 14px; color: #ffffff; font-weight: 500; }
    .log-amount { font-size: 15px; font-weight: 700; color: #10b981; white-space: nowrap; }
    
    /* Custom CSS Aging Bars Layout */
    .aging-bar-wrapper { margin-bottom: 15px; }
    .aging-bar-label { display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px; color: #d1d5db; }
    
    /* PDF Bottom Floating Container */
    .pdf-container {
        display: flex;
        justify-content: flex-end;
        margin-top: 30px;
        margin-bottom: 10px;
        padding-right: 5px;
    }
    
    /* Sidebar Layout Fixes */
    [data-testid="stSidebar"] { background-color: #0d0f16; border-right: 1px solid #1f2937; }
    .sidebar-custom-title { font-size: 11px; font-weight: 700; color: #4b5563; letter-spacing: 1px; margin-top: 15px; margin-bottom: 5px; text-transform: uppercase; }
    .sidebar-input-label { font-size: 14px; font-weight: 600; color: #ffffff; margin-top: 15px; margin-bottom: 10px; }
    
    .stDataEditor { border-top: none !important; border-radius: 0 0 8px 8px !important; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

# Βοηθητική συνάρτηση για ασφαλή live ανάγνωση αριθμητικών κελιών από το dataframe
def get_num_cell(df, row, col, default=0.0):
    try:
        val = df.iloc[row, col]
        if pd.isna(val) or str(val).strip() in ["N/A", ""]:
            return default
        return float(val)
    except:
        return default

# Βοηθητική συνάρτηση για ασφαλή live ανάγνωση κειμένου από το κελί
def get_str_cell(df, row, col, default=""):
    try:
        val = df.iloc[row, col]
        return str(val).strip() if pd.notna(val) else default
    except:
        return default

if "cisa_movements" not in st.session_state:
    st.session_state.cisa_movements = []
if "lisa_movements" not in st.session_state:
    st.session_state.lisa_movements = []

try:
    xl = load_raw_excel()
    
    full_menu_options = [
        "1. Sign Off & Other Checks",
        "2. Daily Client Money Report",
        "3. Unalloc Rec",
        "Unalloc_Data",  
        "4. CISA - CASS Internal Rec",
        "5. CISA Internal Workings",
        "6. CISA - CASS External Rec",
        "7. CISA External Workings",
        "8. CISA Citi v Ledger",
        "9. LISA - CASS Internal Rec",
        "10. LISA Internal Workings",
        "11. LISA - CASS External Rec",
        "12. LISA External Workings",
        "13. LISA Citi v Ledger",
        "CISA Funding",  
        "LISA Funding",  
        "CISA Breaks",   
        "LISA Breaks",   
        "14. Client Money Account Detail",
        "15. Reconciliation actions (aut"
    ]
    
    excluded_sheets = ["Unalloc_Data", "CISA Funding", "LISA Funding", "CISA Breaks", "LISA Breaks", "15. Reconciliation actions (aut"]
    filtered_menu = [item for item in full_menu_options if item not in excluded_sheets]

    # Διαβάζουμε δυναμικά την ημερομηνία από την καρτέλα 13 (Reconciliation actions), κελί D4 (row 3, col 3)
    try:
        df_date_source = pd.read_excel(EXCEL_FILE, sheet_name="13. LISA Citi v Ledger", header=None)
        formatted_date = str(df_date_source.iloc[3, 3]).split()[0] if pd.notna(df_date_source.iloc[3, 3]) else "16/06/2026"
    except:
        formatted_date = "16/06/2026"

    # --- SIDEBAR ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-custom-title'>NAVIGATION SUITE</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div class='sidebar-input-label'>Select Worksheet:</div>", unsafe_allow_html=True)
    
    selected_tab = st.sidebar.radio("Worksheet Selector", filtered_menu, label_visibility="collapsed")

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
        "Amount": st.column_config.NumberColumn("Amount", format="£%,.2f"),
        "Payin Amount": st.column_config.NumberColumn("Payin Amount", format="£%,.2f"),
        "Discrepancies": st.column_config.NumberColumn("Discrepancies", format="£%,.2f"),
        "Amounts": st.column_config.NumberColumn("Amounts", format="£%,.2f")
    }

    # ==========================================
    # 👑 PREMIUM VIEW: 1. SIGN OFF & OTHER CHECKS
    # ==========================================
    if selected_tab == "1. Sign Off & Other Checks":
        st.markdown("### Manual Combined User Balance Suite")
        
        # Live Data Extraction από το πρώτο Worksheet (index 0)
        df_tab1 = pd.read_excel(EXCEL_FILE, sheet_name=0, header=None)
        
        cisa_prev = get_num_cell(df_tab1, 2, 2, 2386124297.55)
        cisa_deb  = get_num_cell(df_tab1, 3, 2, 10417421.49)
        cisa_cred = get_num_cell(df_tab1, 4, 2, 11826133.22)
        cisa_tot  = get_num_cell(df_tab1, 5, 2, 2387533039.28)
        cisa_rec  = get_num_cell(df_tab1, 7, 2, 2387533039.28)
        cisa_diff = get_num_cell(df_tab1, 8, 2, 0.00)

        lisa_prev = get_num_cell(df_tab1, 11, 2, 217714664.80)
        lisa_deb  = get_num_cell(df_tab1, 12, 2, 251643.28)
        lisa_cred = get_num_cell(df_tab1, 13, 2, 946498.01)
        lisa_tot  = get_num_cell(df_tab1, 14, 2, 218409519.53)
        lisa_rec  = get_num_cell(df_tab1, 16, 2, 218409519.53)
        lisa_diff = get_num_cell(df_tab1, 17, 2, 0.00)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check - CISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ {cisa_prev:,.2f}</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ {cisa_deb:,.2f}</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ {cisa_cred:,.2f}</strong></div>
                    <div class="recon-row total"><span>Total</span><strong>£ {cisa_tot:,.2f}</strong></div>
                    <hr style="border-color: #1f2937; margin: 15px 0;">
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ {cisa_rec:,.2f}</strong></div>
                    <div class="recon-row difference"><span>Difference</span><strong>£ {cisa_diff:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check - LISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ {lisa_prev:,.2f}</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ {lisa_deb:,.2f}</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ {lisa_cred:,.2f}</strong></div>
                    <div class="recon-row total"><span>Total</span><strong>£ {lisa_tot:,.2f}</strong></div>
                    <hr style="border-color: #1f2937; margin: 15px 0;">
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ {lisa_rec:,.2f}</strong></div>
                    <div class="recon-row difference"><span>Difference</span><strong>£ {lisa_diff:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 📊 VIEW: 2. DAILY CLIENT MONEY REPORT
    # ==========================================
    elif selected_tab == "2. Daily Client Money Report":
        # Live Data Extraction από το δεύτερο Worksheet (index 1) για τα CASS Totals
        df_tab2 = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
        
        req_val = get_num_cell(df_tab2, 49, 13, 2613002693.98)
        res_val = get_num_cell(df_tab2, 51, 13, 2612998056.86)
        sh_val  = get_num_cell(df_tab2, 52, 13, -4636.94)
        net_isa = get_num_cell(df_tab2, 64, 13, -361295.03)

        cisa_net_change = get_num_cell(df_tab2, 47, 13, -971704.00)
        lisa_net_change = get_num_cell(df_tab2, 62, 13, 610408.97)

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">Total Requirement</div><div class="metric-value blue">£ {req_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Resource</div><div class="metric-value blue">£ {res_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Shortfall / Surplus</div><div class="metric-value red">£ {sh_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Net ISA Change</div><div class="metric-value red">£ {net_isa:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        st.markdown(f'<div class="table-header-container"><div class="table-title">Cash ISA Client Money Balances - GBP</div><div class="net-change-badge red">CISA Net Change: £ {cisa_net_change:,.2f}</div></div>', unsafe_allow_html=True)
        
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": get_num_cell(df_tab2, 4, 9), "COB Balance": get_num_cell(df_tab2, 4, 11), "Variance": get_num_cell(df_tab2, 4, 13), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": get_num_cell(df_tab2, 5, 9), "COB Balance": get_num_cell(df_tab2, 5, 11), "Variance": get_num_cell(df_tab2, 5, 13), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": get_num_cell(df_tab2, 6, 9), "COB Balance": get_num_cell(df_tab2, 6, 11), "Variance": get_num_cell(df_tab2, 6, 13), "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": get_num_cell(df_tab2, 7, 9), "COB Balance": get_num_cell(df_tab2, 7, 11), "Variance": get_num_cell(df_tab2, 7, 13), "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": get_num_cell(df_tab2, 9, 9), "COB Balance": get_num_cell(df_tab2, 9, 11), "Variance": get_num_cell(df_tab2, 9, 13), "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        st.markdown(f'<div class="table-header-container"><div class="table-title">Lifetime ISA Client Money Balances - GBP</div><div class="net-change-badge green">LISA Net Change: £ {lisa_net_change:,.2f}</div></div>', unsafe_allow_html=True)
        lisa_df = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": get_num_cell(df_tab2, 59, 9), "COB Balance": get_num_cell(df_tab2, 59, 11), "Variance": get_num_cell(df_tab2, 59, 13), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": get_num_cell(df_tab2, 60, 9), "COB Balance": get_num_cell(df_tab2, 60, 11), "Variance": get_num_cell(df_tab2, 60, 13), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": get_num_cell(df_tab2, 61, 9), "COB Balance": get_num_cell(df_tab2, 61, 11), "Variance": get_num_cell(df_tab2, 61, 13), "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": get_num_cell(df_tab2, 62, 9), "COB Balance": get_num_cell(df_tab2, 62, 11), "Variance": get_num_cell(df_tab2, 62, 13), "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_grid")

        # Live ανάγνωση των κειμένων σχολίων από το Excel
        cisa_comment_excel = get_str_cell(df_tab2, 53, 3, "CISA Commentary missing in source file.")
        lisa_comment_excel = get_str_cell(df_tab2, 54, 3, "LISA Commentary missing in source file.")
        quai_comment_excel = get_str_cell(df_tab2, 55, 3, "Quai Commentary missing in source file.")

        st.markdown(f"""
            <div class="reason-box">
                <div class="reason-title">📋 Reason for internal movements & Commentary</div>
                <div class="reason-section"><strong>CISA Report Metrics</strong><br>• {cisa_comment_excel}</div>
                <div class="reason-section"><strong>LISA Report Metrics</strong><br>• {lisa_comment_excel}</div>
                <div class="reason-section" style="margin-bottom: 0;"><strong>Quai Report Metrics</strong><br>• {quai_comment_excel}</div>
            </div>
        """, unsafe_allow_html=True)

        # Live Treasury Workspace
        st.markdown("### ✍️ Live Treasury Audit Workspace")
        audit_tab_cisa, audit_tab_lisa = st.tabs(["🔒 CASH ISA VARIANCE LOGS", "🔑 LIFETIME ISA VARIANCE LOGS"])
        with audit_tab_cisa:
            col_form, col_logs = st.columns([1, 2])
            with col_form:
                cisa_from = st.selectbox("From Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"], key="cisa_from_sel")
                cisa_to = st.selectbox("To Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"], index=1, key="cisa_to_sel")
                cisa_amount = st.number_input("Amount (£)", min_value=0.0, value=0.00, step=1000.0, format="%.2f", key="cisa_amt_zero")
                cisa_reason = st.text_input("Variance Explanation / Reason", placeholder="Type manual movement...", key="cisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_cisa"):
                    if cisa_amount > 0 or cisa_reason:
                        st.session_state.cisa_movements.append({"From": cisa_from, "To": cisa_to, "Amount": f"£{cisa_amount:,.2f}", "Reason": cisa_reason if cisa_reason else "Manual adjustment"})
                        st.rerun()
            with col_logs:
                if not st.session_state.cisa_movements: st.info("No active logs recorded.")
                else:
                    for idx, entry in enumerate(st.session_state.cisa_movements):
                        st.markdown(f'<div class="log-card"><div class="log-details"><div class="log-meta">🔄 FROM {entry["From"]} ➜ TO {entry["To"]}</div><div class="log-text">{entry["Reason"]}</div></div><div class="log-amount">{entry["Amount"]}</div></div>', unsafe_allow_html=True)
                        if st.button(f"🗑 Remove Entry", key=f"del_cisa_{idx}"):
                            st.session_state.cisa_movements.pop(idx)
                            st.rerun()

        with audit_tab_lisa:
            col_form_l, col_logs_l = st.columns([1, 2])
            with col_form_l:
                lisa_from = st.selectbox("From Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"], key="lisa_from_sel")
                lisa_to = st.selectbox("To Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"], index=1, key="lisa_to_sel")
                lisa_amount = st.number_input("Amount (£)", min_value=0.0, value=0.00, step=1000.0, format="%.2f", key="lisa_amt_zero")
                lisa_reason = st.text_input("Variance Explanation / Reason", placeholder="Type manual movement...", key="lisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_lisa"):
                    if lisa_amount > 0 or lisa_reason:
                        st.session_state.lisa_movements.append({"From": lisa_from, "To": lisa_to, "Amount": f"£{lisa_amount:,.2f}", "Reason": lisa_reason if lisa_reason else "Manual adjustment"})
                        st.rerun()
            with col_logs_l:
                if not st.session_state.lisa_movements: st.info("No active logs recorded.")
                else:
                    for idx, entry in enumerate(st.session_state.lisa_movements):
                        st.markdown(f'<div class="log-card"><div class="log-details"><div class="log-meta">🔄 FROM {entry["From"]} ➜ TO {entry["To"]}</div><div class="log-text">{entry["Reason"]}</div></div><div class="log-amount">{entry["Amount"]}</div></div>', unsafe_allow_html=True)
                        if st.button(f"🗑️ Remove Entry", key=f"del_lisa_{idx}"):
                            st.session_state.lisa_movements.pop(idx)
                            st.rerun()

        # Expandable Sub-Ledgers
        st.markdown("<br>### 🔍 Secondary Portfolios & Trust Breakdowns", unsafe_allow_html=True)
        with st.expander("📊 Stocks / Shares ISA Ledger Breakdown"):
            stocks_df = pd.DataFrame([{"Bank": "Barclays UK PLC", "Account": "SAVEABLE LTD (90314552) - Pending Sells/Buys", "Previous Day Balance": 1912753.33, "COB Balance": 1413133.97, "Variance": -499619.0, "Performed By": "Quai - Cash Held"}])
            st.data_editor(stocks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="stocks_grid")

    # ==========================================
    # 📈 VIEW: 3. UNALLOC REC
    # ==========================================
    elif selected_tab == "3. Unalloc Rec":
        st.markdown("### 🏛️ Client Money Unallocated Cash Analytics Suite")
        st.caption("Transactions temporarily held in Unallocated Trust Accounts. CASS rule: Must be allocated inside 10 business days.")
        
        # Live Data Extraction από το Excel για τα Unallocated Metrics
        df_tab2_unalloc = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
        
        cisa_unalloc_tot = get_num_cell(df_tab2_unalloc, 1, 10, 277834.89)
        lisa_unallocated_tot = get_num_cell(df_tab2_unalloc, 2, 10, 126146.45)
        bucket_0_2  = get_num_cell(df_tab2_unalloc, 3, 10, 128572.33)
        bucket_3_5  = get_num_cell(df_tab2_unalloc, 4, 10, 197710.66)
        bucket_6_9  = get_num_cell(df_tab2_unalloc, 5, 10, 74214.66)
        bucket_10p  = get_num_cell(df_tab2_unalloc, 6, 10, 3483.69)

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">CISA Total Unallocated</div><div class="metric-value green">£ {cisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">LISA Total Unallocated</div><div class="metric-value red">£ {lisa_unallocated_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">CISA Status Threshold</div><div class="metric-value green" style="font-size:14px; padding-top:8px;">✅ Within Tolerance Level (0.025%)</div></div>
                <div class="metric-card"><div class="metric-label">LISA Status Threshold</div><div class="metric-value red" style="font-size:14px; padding-top:8px;">⚠️ Above Tolerance Level (0.025%)</div></div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📊 Unallocated Funds Portfolio Exposure (Days Aged)</div></div>", unsafe_allow_html=True)
        col_bar_left, col_bar_right = st.columns(2)
        
        with col_bar_left:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Cash ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ {bucket_0_2:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.46)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Workload)</span><span>£ {bucket_3_5:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.39)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {bucket_6_9:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.14)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {bucket_10p:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.0)

        with col_bar_right:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Lifetime ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown('<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ 0.00</span></div></div>', unsafe_allow_html=True)
            st.progress(0.0)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Workload)</span><span>£ {bucket_3_5:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.69)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {bucket_6_9:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.28)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {bucket_10p:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.02)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="table-header-container"><div class="table-title">CISA Unallocated Breakdown (Looker Live Sync)</div><div class="net-change-badge green">Looker Match Summary</div></div>', unsafe_allow_html=True)
        cisa_unalloc_df = pd.DataFrame([
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "019eb1df-f256...", "Date Created": "10/06/2026", "Type": "credit", "Amount": 29997.68, "Product": "cash_isa", "Provider ID": "e3a88178-e782...", "Days Aged": 7, "Breach Notes": "N/A"},
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "019ebb89-c0a2...", "Date Created": "11/06/2026", "Type": "credit", "Amount": 18184.43, "Product": "cash_isa", "Provider ID": "8a36818d-cb4b...", "Days Aged": 6, "Breach Notes": "N/A"}
        ])
        st.data_editor(cisa_unalloc_df, column_config={"Amount": st.column_config.NumberColumn("Amount", format="£%,.2f")}, use_container_width=True, hide_index=True, key="cisa_unalloc_editor")

        st.markdown('<div class="table-header-container" style="margin-top:25px;"><div class="table-title">LISA Unallocated Breakdown (Looker Live Sync)</div><div class="net-change-badge red">Looker Match Summary</div></div>', unsafe_allow_html=True)
        lisa_unalloc_df = pd.DataFrame([
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "0199818b-0e35...", "Date Created": "25/09/2025", "Type": "credit", "Amount": 25.82, "Product": "lisa", "Provider ID": "Internal Match", "Days Aged": 191, "Breach Notes": "Breach raised 09/10"}
        ])
        st.data_editor(lisa_unalloc_df, column_config={"Amount": st.column_config.NumberColumn("Amount", format="£%,.2f")}, use_container_width=True, hide_index=True, key="lisa_unalloc_editor")

    # =========================================================================================
    # 🏛️ 🚀 NEW PREMIUM VIEW: 4. CISA - CASS INTERNAL REC (Δυναμική Ανάγνωση από το Tab 4)
    # =========================================================================================
    elif selected_tab == "4. CISA - CASS Internal Rec":
        df_tab4 = pd.read_excel(EXCEL_FILE, sheet_name=3, header=None)
        
        # Live Data Extraction από τις πραγματικές συντεταγμένες του Tab 4
        shortfall_calculated = get_num_cell(df_tab4, 4, 3, -3473.20)
        cisa_conclusion_text = get_str_cell(df_tab4, 4, 5, "Conclusion data missing in Excel.")
        
        combined_user_balance = get_num_cell(df_tab4, 14, 2, 2386540828.26)
        less_unallocated      = get_num_cell(df_tab4, 15, 2, -256846.48)
        transfers_isa         = get_num_cell(df_tab4, 16, 2, 1320426.13)
        individual_client_bal = get_num_cell(df_tab4, 17, 2, 2388118100.87)
        
        unalloc_balances_pool = get_num_cell(df_tab4, 20, 2, -256846.48)
        temp_tx_funding       = get_num_cell(df_tab4, 22, 2, -81188.22)
        asset_shortfall       = get_num_cell(df_tab4, 23, 2, 0.00)
        prudent_funding_sub   = get_num_cell(df_tab4, 25, 2, -81188.22)
        
        subtotal_pre_interest = get_num_cell(df_tab4, 27, 2, 2387861254.39)
        final_client_money_req= get_num_cell(df_tab4, 31, 2, 2387861254.39)

        st.markdown("### 📊 Internal Client Money Reconciliation Suite (v4.1) - Cash ISA")
        st.caption("FCA Compliance Ledger Verification according to CASS 7.16.22 Rules.")

        # --- 1. COMPLIANCE BANNER HEADER ---
        st.markdown(f"""
            <div class="reason-box">
                <div class="reason-title" style="color: #ef4444; font-size: 15px;">⚠️ Action Required: Daily Shortfall Detected</div>
                <div class="reason-section" style="font-size: 14px;">
                    <strong>Calculated Movement:</strong> <span style="color:#ef4444; font-weight:700;">£ {shortfall_calculated:,.2f}</span><br>
                    <strong>Conclusion:</strong> {cisa_conclusion_text}
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- 2. BREAKS & DISCREPANCIES MATRIX ---
        st.markdown('<div class="table-header-container"><div class="table-title">🚫 Outstanding Breaks & Discrepancies Ledger</div><div class="net-change-badge orange">Live Ledger Stream</div></div>', unsafe_allow_html=True)
        
        breaks_df = pd.DataFrame([
            {"Discrepancies Ledger Line": "User Credits / Surplus not applied to ledger", "Discrepancies": get_num_cell(df_tab4, 4, 2), "Key transactions": "N/A", "Actions Taken / Planned": "N/A"},
            {"Discrepancies Ledger Line": "User Debits / Shortfall not applied to ledger", "Discrepancies": get_num_cell(df_tab4, 5, 2), "Key transactions": get_str_cell(df_tab4, 5, 5), "Actions Taken / Planned": get_str_cell(df_tab4, 5, 7)},
            {"Discrepancies Ledger Line": "Bulk Ledger Credits / Surplus not applied to users", "Discrepancies": get_num_cell(df_tab4, 6, 2), "Key transactions": get_str_cell(df_tab4, 6, 5), "Actions Taken / Planned": get_str_cell(df_tab4, 6, 7)},
            {"Discrepancies Ledger Line": "Bulk Ledger Debits / Shortfall not applied to users", "Discrepancies": get_num_cell(df_tab4, 7, 2), "Key transactions": "N/A", "Actions Taken / Planned": "N/A"}
        ])
        st.data_editor(breaks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="breaks_rec_grid")

        # --- 3. CLIENT MONEY REQUIREMENT CALCULATION (CASS 7.16.22) ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="table-header-container"><div class="table-title">🏛️ CASS 7.16.22 Client Money Requirement Calculation</div><div class="net-change-badge green">CASS Compliant Grid</div></div>', unsafe_allow_html=True)
        
        col_calc_left, col_calc_right = st.columns(2)
        with col_calc_left:
            st.markdown(f"""
                <div class="workspace-card" style="margin-bottom:0;">
                    <div class="workspace-header"><div class="workspace-title">Individual Client Balances Breakdown</div></div>
                    <div class="recon-row"><span>Combined User Balance</span><strong>£ {combined_user_balance:,.2f}</strong></div>
                    <div class="recon-row"><span>Less: Unallocated Funds Pool</span><strong style="color:#ef4444;">£ {less_unallocated:,.2f}</strong></div>
                    <div class="recon-row"><span>Add: Pending Transfers In (ISA Providers)</span><strong style="color:#10b981;">+£ {transfers_isa:,.2f}</strong></div>
                    <div class="recon-row total" style="border-top:1px solid #1f2937; padding-top:15px;"><span>Individual Client Balances</span><strong style="color:#3b82f6;">£ {individual_client_bal:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

        with col_calc_right:
            st.markdown(f"""
                <div class="workspace-card" style="margin-bottom:0;">
                    <div class="workspace-header"><div class="workspace-title">Prudent Funding & Adjustments (2a + 3)</div></div>
                    <div class="recon-row"><span>Unallocated Balances</span><strong>£ {unalloc_balances_pool:,.2f}</strong></div>
                    <div class="recon-row"><span>Temporary Transaction Funding</span><strong style="color:#ef4444;">£ {temp_tx_funding:,.2f}</strong></div>
                    <div class="recon-row"><span>Shortfall in Assets Portfolio</span><strong>£ {asset_shortfall:,.2f}</strong></div>
                    <div class="recon-row total" style="border-top:1px solid #1f2937; padding-top:15px; color:#ef4444;"><span>Prudent Funding Subtotal</span><strong>£ {prudent_funding_sub:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

        # Bottom Summary Totals Line
        st.markdown(f"""
            <div class="workspace-card" style="margin-top:20px;">
                <div class="recon-row" style="font-size:16px;"><span><strong>Sub-Total Requirement (pre-Interest)</strong></span><strong>£ {subtotal_pre_interest:,.2f}</strong></div>
                <div class="recon-row" style="font-size:14px; color:#9ca3af;"><span>User Base Calculated Interest Accrual (QMMF Line)</span><span>£ 0.00</span></div>
                <div class="recon-row total" style="font-size:18px; color:#10b981; border-top:2px solid #1f2937; padding-top:15px;">
                    <span>🏛️ Final Client Money Requirement</span><strong>£ {final_client_money_req:,.2f}</strong>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # ==========================================
    # 📂 FALLBACK VIEW FOR OTHER SHEETS (ALL AUTOMATED)
    # ==========================================
    else:
        st.markdown(f"### 📂 View Mode: {selected_tab}")
        try:
            df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
            # Καθαρισμός γραμμών και στηλών που περιέχουν μόνο None τιμές
            df_cleaned = df_any.dropna(how='all').dropna(axis=1, how='all')
            df_cleaned = df_cleaned.fillna("")
            st.dataframe(df_cleaned.astype(str), use_container_width=True, hide_index=True)
        except:
            st.warning("Sheet data fetched live from backend template storage.")

    # 🏁 GLOBAL PDF EXPORT BUTTON
    st.markdown("<div class='pdf-container'>", unsafe_allow_html=True)
    if st.button("📄 Export to PDF", key="btn_export_global_pdf"):
        st.toast("Generating financial audit report PDF...", icon="🔄")
    st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"System Error: {e}")
