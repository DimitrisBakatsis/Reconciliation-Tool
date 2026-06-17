import streamlit as st
import pandas as pd
import re

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

    .workspace-card { background-color: #0d0f16; border: 1px solid #1f2937; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
    .workspace-header { border-bottom: 1px solid #1f2937; padding-bottom: 14px; margin-bottom: 20px; }
    .workspace-title { font-size: 13px; font-weight: 700; color: #a78bfa; text-transform: uppercase; letter-spacing: 0.5px; }
    .recon-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #1f2937; font-size: 14px; color: #d1d5db; }
    .recon-row.total { border-bottom: none; font-weight: 700; font-size: 15px; color: #3b82f6; padding-top: 14px; }
    
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
    
    /* Inputs */
    div[data-testid="stSelectbox"] > label, div[data-testid="stNumberInput"] > label, div[data-testid="stTextInput"] > label {
        color: #ffffff !important; font-size: 13px !important; font-weight: 600 !important; margin-bottom: 6px !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"], div[data-testid="stNumberInput"] input, div[data-testid="stTextInput"] input {
        background-color: #1a1c24 !important; border: 1px solid #2d3748 !important; color: #ffffff !important; border-radius: 6px !important;
    }
    .stButton > button {
        background-color: #1a1c24 !important; color: #ffffff !important; border: 1px solid #4a5568 !important; border-radius: 6px !important;
    }
    [data-testid="stSidebar"] { background-color: #0d0f16; border-right: 1px solid #1f2937; }
    .sidebar-section-title { font-size: 11px; font-weight: 700; color: #4b5563; letter-spacing: 1px; margin-top: 20px; margin-bottom: 8px; text-transform: uppercase; }
    
    /* Remove padding under data editor headers */
    .stDataEditor { border-top: none !important; border-radius: 0 0 8px 8px !important; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    formatted_date = "16/06/2026"

    # --- SIDEBAR ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-section-title'>Active Worksheets</div>")
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
                        <div class="workspace-header"><div class="workspace-title" style="color:#ffffff;">{title}</div></div>
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
        
        # 1. KPI Cards
        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Total Requirement</div>
                    <div class="metric-value blue">£ 2,601,370,286.70</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Resource</div>
                    <div class="metric-value blue">£ 2,607,556,676.61</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Shortfall / Surplus</div>
                    <div class="metric-value red">£ -1,244.09</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Net Change (COB)</div>
                    <div class="metric-value blue">£ 3,246,757.00</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        
        if "cisa_movements" not in st.session_state:
            st.session_state.cisa_movements = [{"From": "Citibank", "To": "Lloyds EA", "Amount": "£1,000,000.00", "Reason": "CISA treasury movement of 1,000,000 between Citibank and Lloyds EA"}]
        if "lisa_movements" not in st.session_state:
            st.session_state.lisa_movements = []

        # --- 2. CASH ISA LEDGER (WITH ALL BANKS & BADGE) ---
        st.markdown("""
            <div class="table-header-container">
                <div class="table-title">Cash ISA Client Money Balances - GBP</div>
                <div class="net-change-badge red">CISA Net Change: -£971,704.00</div>
            </div>
        """, unsafe_allow_html=True)
        
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": 176992857.0, "COB Balance": 176021153.0, "Variance": -971704.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": 3959844.0, "COB Balance": 3959844.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": 747535672.0, "COB Balance": 747535672.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": 1168000000.0, "COB Balance": 1168000000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BlackRock QMMF", "Account": "Blackrock QMMF", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Notice account (06758170)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Clydesdale Bank PLC", "Account": "Saveable Cash ISA 95 Day Notice (12204224)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "JP Morgan", "Account": "JP Morgan Client Money account (76919500)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        # --- 3. LIFETIME ISA LEDGER (WITH ALL BANKS & BADGE) ---
        st.markdown("""
            <div class="table-header-container">
                <div class="table-title">Lifetime ISA Client Money Balances - GBP</div>
                <div class="net-change-badge green">LISA Net Change: +£610,408.97</div>
            </div>
        """, unsafe_allow_html=True)
        
        lisa_df = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": 38363170.0, "COB Balance": 38973579.0, "Variance": 610408.97, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": 714980.0, "COB Balance": 714980.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": 79300000.0, "COB Balance": 79300000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": 100000000.0, "COB Balance": 100000000.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_grid")

        # 4. REASON FOR INTERNAL MOVEMENTS COMMENTARY BLOCK
        st.markdown("""
            <div class="reason-box">
                <div class="reason-title">📋 Reason for internal movements & Commentary</div>
                <div class="reason-section">
                    <strong>CISA: Overall Shortfall of £4,393.67</strong><br>
                    • Amount of £4,393.67 residual interest paid to users as part of the transfer out process.<br>
                    • To be moved from CISA corporate interest to CM 17/06.
                </div>
                <div class="reason-section">
                    <strong>LISA: Overall Shortfall of £243.63</strong><br>
                    • Amount of £243.63 residual interest paid to users as part of the transfer out process.<br>
                    • To be moved from LISA corporate interest to CM 17/06.
                </div>
                <div class="reason-section" style="margin-bottom: 0;">
                    <strong>Quai: Overall Surplus of £0.18</strong><br>
                    • Quai to arrange amount to written off 17/06.
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 5. AUDITING & COMMENTARY LOGS
        st.markdown("### ✍️ Live Treasury Audit Workspace")
        cisa_accounts = ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"]
        lisa_accounts = ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"]
        
        audit_tab_cisa, audit_tab_lisa = st.tabs(["🔒 CASH ISA VARIANCE LOGS", "🔑 LIFETIME ISA VARIANCE LOGS"])
        
        with audit_tab_cisa:
            st.markdown("<br>", unsafe_allow_html=True)
            col_form, col_logs = st.columns([1, 2])
            with col_form:
                st.markdown("<p style='font-weight:700; color:#fff; font-size:14px; margin-bottom:15px;'>Log Manual Treasury Movement</p>", unsafe_allow_html=True)
                cisa_from = st.selectbox("From Account", cisa_accounts, key="cisa_from_sel")
                cisa_to = st.selectbox("To Account", cisa_accounts, index=1, key="cisa_to_sel")
                cisa_amount = st.number_input("Amount (£)", min_value=0.0, value=1000000.0, step=100000.0, format="%.2f", key="cisa_amt_input")
                cisa_reason = st.text_input("Variance Explanation / Reason", value="Treasury movement liquidity coverage", key="cisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_cisa"):
                    st.session_state.cisa_movements.append({"From": cisa_from, "To": cisa_to, "Amount": f"£{cisa_amount:,.2f}", "Reason": cisa_reason})
                    st.rerun()
            with col_logs:
                st.markdown("<p style='font-weight:700; color:#a78bfa; font-size:12px; letter-spacing:0.5px; margin-bottom:15px;'>ACTIVE VARIANCE ADJUSTMENTS</p>", unsafe_allow_html=True)
                if not st.session_state.cisa_movements:
                    st.info("No active logs recorded for Cash ISA.")
                else:
                    for idx, entry in enumerate(st.session_state.cisa_movements):
                        st.markdown(f"""
                            <div class="log-card">
                                <div class="log-details">
                                    <div class="log-meta">🔄 FROM {entry['From']} ➜ TO {entry['To']}</div>
                                    <div class="log-text">{entry['Reason']}</div>
                                </div>
                                <div class="log-amount">{entry['Amount']}</div>
                            </div>
                        """, unsafe_allow_html=True)
                        if st.button(f"🗑 Removal Entry", key=f"del_cisa_{idx}"):
                            st.session_state.cisa_movements.pop(idx)
                            st.rerun()

        with audit_tab_lisa:
            st.markdown("<br>", unsafe_allow_html=True)
            col_form_l, col_logs_l = st.columns([1, 2])
            with col_form_l:
                st.markdown("<p style='font-weight:700; color:#fff; font-size:14px; margin-bottom:15px;'>Log Manual Treasury Movement</p>", unsafe_allow_html=True)
                lisa_from = st.selectbox("From Account", lisa_accounts, key="lisa_from_sel")
                lisa_to = st.selectbox("To Account", lisa_accounts, index=1, key="lisa_to_sel")
                lisa_amount = st.number_input("Amount (£)", min_value=0.0, value=1000000.0, step=100000.0, format="%.2f", key="lisa_amt_input")
                lisa_reason = st.text_input("Variance Explanation / Reason", value="Treasury movement liquidity coverage", key="lisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_lisa"):
                    st.session_state.lisa_movements.append({"From": lisa_from, "To": lisa_to, "Amount": f"£{lisa_amount:,.2f}", "Reason": lisa_reason})
                    st.rerun()
            with col_logs_l:
                st.markdown("<p style='font-weight:700; color:#a78bfa; font-size:12px; letter-spacing:0.5px; margin-bottom:15px;'>ACTIVE VARIANCE ADJUSTMENTS</p>", unsafe_allow_html=True)
                if not st.session_state.lisa_movements:
                    st.info("No active logs recorded for Lifetime ISA.")
                else:
                    for idx, entry in enumerate(st.session_state.lisa_movements):
                        st.markdown(f"""
                            <div class="log-card">
                                <div class="log-details">
                                    <div class="log-meta">🔄 FROM {entry['From']} ➜ TO {entry['To']}</div>
                                    <div class="log-text">{entry['Reason']}</div>
                                </div>
                                <div class="log-amount">{entry['Amount']}</div>
                            </div>
                        """, unsafe_allow_html=True)
                        if st.button(f"🗑️ Remove Entry", key=f"del_lisa_{idx}"):
                            st.session_state.lisa_movements.pop(idx)
                            st.rerun()

        # 6. EXPANDABLE SUB-LEDGERS
        st.markdown("<br>### 🔍 Secondary Portfolios & Trust Breakdowns", unsafe_allow_html=True)
        with st.expander("📊 Stocks / Shares ISA Ledger Breakdown"):
            stocks_df = pd.DataFrame([
                {"Bank": "Barclays UK PLC", "Account": "SAVEABLE LTD (90314552) - Pending Sells/Buys", "Previous Day Balance": 1912753.33, "COB Balance": 1413133.97, "Variance": -499619.0, "Performed By": "Quai - Cash Held"}
            ])
            st.data_editor(stocks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="stocks_grid")

        # 7. DAILY RECONCILIATION CALCULATION BLOCK
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>⚙️ CASS Corporate Daily Calculation Suite</div></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1: st.number_input("Total Pure Requirement (£)", value=2601370286.70, format="%.2f", disabled=True)
        with col2: st.number_input("Inclusive of Transfers In to Apply (£)", value=6187634.40, format="%.2f", disabled=True)
        with col3: st.number_input("Total Eligible Resource Pool (£)", value=2607556676.61, format="%.2f", disabled=True)
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.markdown(f"### 📂 Sub-Ledger Archive View: {selected_tab}")
        df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
        st.dataframe(df_any.dropna(how='all').reset_index(drop=True), use_container_width=True)

except Exception as e:
    st.error(f"System Error: {e}")
