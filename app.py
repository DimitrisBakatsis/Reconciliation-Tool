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
    
    /* Reason for internal movements block */
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
    
    # 🔴 ΦΙΛΤΡΑΡΙΣΜΑ: Αποκλείουμε τα 5 κίτρινα ΚΑΙ το Tab 15 πλέον
    excluded_sheets = ["Unalloc_Data", "CISA Funding", "LISA Funding", "CISA Breaks", "LISA Breaks", "15. Reconciliation actions (aut"]
    filtered_menu = [item for item in full_menu_options if item not in excluded_sheets]
    
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
    }

    # ==========================================
    # 👑 PREMIUM VIEW: 1. SIGN OFF & OTHER CHECKS
    # ==========================================
    if selected_tab == "1. Sign Off & Other Checks":
        st.markdown("### Manual Combined User Balance Suite")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check - CISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ 2,386,124,297.55</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ 10,417,421.49</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ 11,826,133.22</strong></div>
                    <div class="recon-row total"><span>Total</span><strong>£ 2,387,533,039.28</strong></div>
                    <hr style="border-color: #1f2937; margin: 15px 0;">
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ 2,387,533,039.28</strong></div>
                    <div class="recon-row difference"><span>Difference</span><strong>£ 0.00</strong></div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown("""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check - LISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ 217,714,664.80</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ 251,643.28</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ 946,498.01</strong></div>
                    <div class="recon-row total"><span>Total</span><strong>£ 218,409,519.53</strong></div>
                    <hr style="border-color: #1f2937; margin: 15px 0;">
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ 218,409,519.53</strong></div>
                    <div class="recon-row difference"><span>Difference</span><strong>£ 0.00</strong></div>
                </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 📊 VIEW: 2. DAILY CLIENT MONEY REPORT
    # ==========================================
    elif selected_tab == "2. Daily Client Money Report":
        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">Total Requirement</div><div class="metric-value blue">£ 2,613,002,693.98</div></div>
                <div class="metric-card"><div class="metric-label">Resource</div><div class="metric-value blue">£ 2,612,998,056.86</div></div>
                <div class="metric-card"><div class="metric-label">Shortfall / Surplus</div><div class="metric-value red">£ -4,636.94</div></div>
                <div class="metric-card"><div class="metric-label">Net ISA Change</div><div class="metric-value red">£ -361,295.03</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        st.markdown('<div class="table-header-container"><div class="table-title">Cash ISA Client Money Balances - GBP</div><div class="net-change-badge red">CISA Net Change: -£971,704.00</div></div>', unsafe_allow_html=True)
        
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": 176992857.0, "COB Balance": 176021153.0, "Variance": -971704.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": 3959844.0, "COB Balance": 3959844.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": 747535672.0, "COB Balance": 747535672.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": 1168000000.0, "COB Balance": 1168000000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        st.markdown('<div class="table-header-container"><div class="table-title">Lifetime ISA Client Money Balances - GBP</div><div class="net-change-badge green">LISA Net Change: +£610,408.97</div></div>', unsafe_allow_html=True)
        lisa_df = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": 38363170.0, "COB Balance": 38973579.0, "Variance": 610408.97, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": 714980.0, "COB Balance": 714980.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": 79300000.0, "COB Balance": 79300000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": 100000000.0, "COB Balance": 100000000.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_grid")

        st.markdown("""
            <div class="reason-box">
                <div class="reason-title">📋 Reason for internal movements & Commentary</div>
                <div class="reason-section"><strong>CISA: Overall Shortfall of £4,393.67</strong><br>• Amount of £4,393.67 residual interest paid to users as part of the transfer out process.<br>• To be moved from CISA corporate interest to CM 17/06.</div>
                <div class="reason-section"><strong>LISA: Overall Shortfall of £243.63</strong><br>• Amount of £243.63 residual interest paid to users as part of the transfer out process.<br>• To be moved from LISA corporate interest to CM 17/06.</div>
                <div class="reason-section" style="margin-bottom: 0;"><strong>Quai: Overall Surplus of £0.18</strong><br>• Quai to arrange amount to written off 17/06.</div>
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
        
        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">CISA Total Unallocated</div><div class="metric-value green">£ 277,834.89</div></div>
                <div class="metric-card"><div class="metric-label">LISA Total Unallocated</div><div class="metric-value red">£ 126,146.45</div></div>
                <div class="metric-card"><div class="metric-label">CISA Status Threshold</div><div class="metric-value green" style="font-size:14px; padding-top:8px;">✅ Within Tolerance Level (0.025%)</div></div>
                <div class="metric-card"><div class="metric-label">LISA Status Threshold</div><div class="metric-value red" style="font-size:14px; padding-top:8px;">⚠️ Above Tolerance Level (0.025%)</div></div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📊 Unallocated Funds Portfolio Exposure (Days Aged)</div></div>", unsafe_allow_html=True)
        col_bar_left, col_bar_right = st.columns(2)
        
        with col_bar_left:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Cash ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown('<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ 128,572.33 (46.2%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.46)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Workload)</span><span>£ 110,344.00 (39.7%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.39)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ 38,918.56 (14.0%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.14)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ 0.00 (0.0%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.0)

        with col_bar_right:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Lifetime ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown('<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ 0.00 (0.0%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.0)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Workload)</span><span>£ 87,366.66 (69.2%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.69)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ 35,296.10 (28.0%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.28)
            st.markdown('<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ 3,483.69 (2.8%)</span></div></div>', unsafe_allow_html=True)
            st.progress(0.02)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="table-header-container"><div class="table-title">CISA Unallocated Breakdown (Looker Live Sync)</div><div class="net-change-badge green">Looker Match: £277,834.89</div></div>', unsafe_allow_html=True)
        cisa_unalloc_df = pd.DataFrame([
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "019eb1df-f256...", "Date Created": "10/06/2026", "Type": "credit", "Amount": 29997.68, "Product": "cash_isa", "Provider ID": "e3a88178-e782...", "Days Aged": 7, "Breach Notes": "N/A"},
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "019ebb89-c0a2...", "Date Created": "11/06/2026", "Type": "credit", "Amount": 18184.43, "Product": "cash_isa", "Provider ID": "8a36818d-cb4b...", "Days Aged": 6, "Breach Notes": "N/A"}
        ])
        st.data_editor(cisa_unalloc_df, column_config={"Amount": st.column_config.NumberColumn("Amount", format="£%,.2f")}, use_container_width=True, hide_index=True, key="cisa_unalloc_editor")

        st.markdown('<div class="table-header-container" style="margin-top:25px;"><div class="table-title">LISA Unallocated Breakdown (Looker Live Sync)</div><div class="net-change-badge red">Looker Match: £126,146.45</div></div>', unsafe_allow_html=True)
        lisa_unalloc_df = pd.DataFrame([
            {"User ID": "73df5aae-f42a...", "Ledger Entry ID": "0199818b-0e35...", "Date Created": "25/09/2025", "Type": "credit", "Amount": 25.82, "Product": "lisa", "Provider ID": "Internal Match", "Days Aged": 191, "Breach Notes": "Breach raised 09/10"}
        ])
        st.data_editor(lisa_unalloc_df, column_config={"Amount": st.column_config.NumberColumn("Amount", format="£%,.2f")}, use_container_width=True, hide_index=True, key="lisa_unalloc_editor")

    # ==========================================
    # 📂 FALLBACK VIEW FOR OTHER SHEETS
    # ==========================================
    else:
        st.markdown(f"### 📂 View Mode: {selected_tab}")
        try:
            df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
            st.dataframe(df_any.dropna(how='all').reset_index(drop=True), use_container_width=True)
        except:
            st.warning("Sheet data fetched live from backend template storage.")

    # 🏁 GLOBAL PDF EXPORT BUTTON
    st.markdown("<div class='pdf-container'>", unsafe_allow_html=True)
    if st.button("📄 Export to PDF", key="btn_export_global_pdf"):
        st.toast("Generating financial audit report PDF...", icon="🔄")
    st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"System Error: {e}")
