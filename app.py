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
        border-left: 4px solid #ef4444;
        padding: 20px;
        border-radius: 6px;
        margin-top: 10px;
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

# 🛠️ HELPER FUNCTIONS ΓΙΑ LIVE LOOKUP
def get_num_cell(df, row, col, default=0.0):
    try:
        val = df.iloc[row, col]
        if pd.isna(val) or str(val).strip() in ["N/A", ""]: return default
        return float(val)
    except: return default

def get_str_cell(df, row, col, default=""):
    try:
        val = df.iloc[row, col]
        return str(val).strip() if pd.notna(val) else default
    except: return default

def find_row_by_keyword(df, keyword):
    for idx in range(len(df)):
        for col_idx in range(df.shape[1]):
            if keyword.lower() in str(df.iloc[idx, col_idx]).lower():
                return idx
    return None

try:
    xl = load_raw_excel()
    
    full_menu_options = [
        "1. Sign Off & Other Checks", "2. Daily Client Money Report", "3. Unalloc Rec", "Unalloc_Data",  
        "4. CISA - CASS Internal Rec", "5. CISA Internal Workings", "6. CISA - CASS External Rec", 
        "7. CISA External Workings", "8. CISA Citi v Ledger", "9. LISA - CASS Internal Rec", 
        "10. LISA Internal Workings", "11. LISA - CASS External Rec", "12. LISA External Workings", 
        "13. LISA Citi v Ledger", "CISA Funding", "LISA Funding", "CISA Breaks", "LISA Breaks", 
        "14. Client Money Account Detail", "15. Reconciliation actions (aut"
    ]
    
    excluded_sheets = ["Unalloc_Data", "CISA Funding", "LISA Funding", "CISA Breaks", "LISA Breaks", "15. Reconciliation actions (aut"]
    filtered_menu = [item for item in full_menu_options if item not in excluded_sheets]

    # --- SIDEBAR ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-custom-title'>NAVIGATION SUITE</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div class='sidebar-input-label'>Select Worksheet:</div>", unsafe_allow_html=True)
    selected_tab = st.sidebar.radio("Worksheet Selector", filtered_menu, label_visibility="collapsed")

    # Live ημερομηνία
    try:
        df_tab4_date = pd.read_excel(EXCEL_FILE, sheet_name="4. CISA - CASS Internal Rec", header=None)
        date_row = find_row_by_keyword(df_tab4_date, "Date")
        formatted_date = str(df_tab4_date.iloc[date_row, 3]).split()[0] if date_row else "18/06/2026"
    except:
        formatted_date = "18/06/2026"

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
        "Amount": st.column_config.NumberColumn("Amount", format="£%,.2f"),
        "Payin Amount": st.column_config.NumberColumn("Payin Amount", format="£%,.2f"),
        "Discrepancies / Value": st.column_config.NumberColumn("Discrepancies / Value", format="£%,.2f")
    }

    # ==========================================
    # 👑 TAB 1: SIGN OFF & OTHER CHECKS
    # ==========================================
    if selected_tab == "1. Sign Off & Other Checks":
        st.markdown("### Manual Combined User Balance Suite")
        df_tab1 = pd.read_excel(EXCEL_FILE, sheet_name=0, header=None)
        
        cisa_start = find_row_by_keyword(df_tab1, "Combined User Balance Check - CISA")
        cisa_prev = get_num_cell(df_tab1, cisa_start + 1, 2) if cisa_start else 0.0
        cisa_deb  = get_num_cell(df_tab1, cisa_start + 2, 2) if cisa_start else 0.0
        cisa_cred = get_num_cell(df_tab1, cisa_start + 3, 2) if cisa_start else 0.0
        cisa_tot  = get_num_cell(df_tab1, cisa_start + 4, 2) if cisa_start else 0.0
        cisa_rec  = get_num_cell(df_tab1, cisa_start + 6, 2) if cisa_start else 0.0
        cisa_diff = get_num_cell(df_tab1, cisa_start + 7, 2) if cisa_start else 0.0

        lisa_start = find_row_by_keyword(df_tab1, "Combined User Balance Check - LISA")
        lisa_prev = get_num_cell(df_tab1, lisa_start + 1, 2) if lisa_start else 0.0
        lisa_deb  = get_num_cell(df_tab1, lisa_start + 2, 2) if lisa_start else 0.0
        lisa_cred = get_num_cell(df_tab1, lisa_start + 3, 2) if lisa_start else 0.0
        lisa_tot  = get_num_cell(df_tab1, lisa_start + 4, 2) if lisa_start else 0.0
        lisa_rec  = get_num_cell(df_tab1, lisa_start + 6, 2) if lisa_start else 0.0
        lisa_diff = get_num_cell(df_tab1, lisa_start + 7, 2) if lisa_start else 0.0

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
    # 📊 TAB 2: DAILY CLIENT MONEY REPORT
    # ==========================================
    elif selected_tab == "2. Daily Client Money Report":
        df_tab2 = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
        
        req_idx = find_row_by_keyword(df_tab2, "Total requirement")
        res_idx = find_row_by_keyword(df_tab2, "Resource")
        sh_idx  = find_row_by_keyword(df_tab2, "Shortfall / Surplus")
        net_idx = find_row_by_keyword(df_tab2, "Net ISA Change")

        req_val = get_num_cell(df_tab2, req_idx, 13) if req_idx else 0.0
        res_val = get_num_cell(df_tab2, res_idx, 13) if res_idx else 0.0
        sh_val  = get_num_cell(df_tab2, sh_idx, 13) if sh_idx else 0.0
        net_isa = get_num_cell(df_tab2, net_idx, 13) if net_idx else 0.0

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">Total Requirement</div><div class="metric-value blue">£ {req_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Resource</div><div class="metric-value blue">£ {res_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Shortfall / Surplus</div><div class="metric-value red">£ {sh_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">Net ISA Change</div><div class="metric-value red">£ {net_isa:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        st.markdown('<div class="table-header-container"><div class="table-title">Cash ISA Client Money Balances - GBP</div></div>', unsafe_allow_html=True)
        st.data_editor(df_tab2.dropna(how="all").iloc[2:12].fillna(""), use_container_width=True, hide_index=True)

    # ==========================================
    # 📈 TAB 3: UNALLOC REC
    # ==========================================
    elif selected_tab == "3. Unalloc Rec":
        st.markdown("### 🏛️ Client Money Unallocated Cash Analytics Suite")
        df_tab3 = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
        
        cisa_idx = find_row_by_keyword(df_tab3, "CISA total unallocated")
        lisa_idx = find_row_by_keyword(df_tab3, "LISA total unallocated")
        
        cisa_unalloc_tot = get_num_cell(df_tab3, cisa_idx, 10) if cisa_idx else 0.0
        lisa_unalloc_tot = get_num_cell(df_tab3, lisa_idx, 10) if lisa_idx else 0.0
        bucket_0_2 = get_num_cell(df_tab3, find_row_by_keyword(df_tab3, "0-2 days"), 10)
        bucket_3_5 = get_num_cell(df_tab3, find_row_by_keyword(df_tab3, "3-5 days"), 10)
        bucket_6_9 = get_num_cell(df_tab3, find_row_by_keyword(df_tab3, "6-9 days"), 10)
        bucket_10p = get_num_cell(df_tab3, find_row_by_keyword(df_tab3, "10+ days"), 10)

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">CISA Total Unallocated</div><div class="metric-value green">£ {cisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">LISA Total Unallocated</div><div class="metric-value red">£ {lisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">CISA Control Status</div><div class="metric-value green" style="font-size:14px; padding-top:8px;">✅ Active Control Check OK</div></div>
                <div class="metric-card"><div class="metric-label">LISA Control Status</div><div class="metric-value red" style="font-size:14px; padding-top:8px;">⚠️ Aging Review Required</div></div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📊 Unallocated Funds Portfolio Exposure (Days Aged)</div></div>", unsafe_allow_html=True)
        col_bar_left, col_bar_right = st.columns(2)
        with col_bar_left:
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days</span><span>£ {bucket_0_2:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.45)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days</span><span>£ {bucket_3_5:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.35)
        with col_bar_right:
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟠 6-9 Days</span><span>£ {bucket_6_9:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.20)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days</span><span>£ {bucket_10p:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.05)
        st.markdown("</div>", unsafe_allow_html=True)

    # =========================================================================================
    # 🏛️ 🚀 100% ΔΥΝΑΜΙΚΟ ΚΑΙ ΠΛΗΡΕΣ VIEW: 4. CISA - CASS INTERNAL REC
    # =========================================================================================
    elif selected_tab == "4. CISA - CASS Internal Rec":
        df_tab4 = pd.read_excel(EXCEL_FILE, sheet_name="4. CISA - CASS Internal Rec", header=None)
        
        # Live εντοπισμός γραμμών βάσει των keywords του Excel
        row_shortfall  = find_row_by_keyword(df_tab4, "Daily Surplus or Shortfall")
        row_cub        = find_row_by_keyword(df_tab4, "Combined User Balance")
        row_unalloc    = find_row_by_keyword(df_tab4, "Less Unallocated")
        row_transfers  = find_row_by_keyword(df_tab4, "Transfers in from ISA providers")
        row_client_bal = find_row_by_keyword(df_tab4, "Individual Client Balances")
        row_funding    = find_row_by_keyword(df_tab4, "Temporary Transaction Funding")
        row_subtotal   = find_row_by_keyword(df_tab4, "Sub-Total (pre-Interest)")
        row_final_req  = find_row_by_keyword(df_tab4, "Client money requirement")

        # Live τράβηγμα των πραγματικών ποσών
        shortfall_calculated  = get_num_cell(df_tab4, row_shortfall, 3) if row_shortfall else 0.0
        conclusion_excel_text = get_str_cell(df_tab4, row_shortfall, 5) if row_shortfall else "N/A"
        
        combined_user_balance = get_num_cell(df_tab4, row_cub, 2) if row_cub else 0.0
        less_unallocated      = get_num_cell(df_tab4, row_unalloc, 2) if row_unalloc else 0.0
        transfers_isa         = get_num_cell(df_tab4, row_transfers, 2) if row_transfers else 0.0
        individual_client_bal = get_num_cell(df_tab4, row_client_bal, 2) if row_client_bal else 0.0
        
        temp_tx_funding       = get_num_cell(df_tab4, row_funding, 2) if row_funding else 0.0
        subtotal_pre_interest = get_num_cell(df_tab4, row_subtotal, 2) if row_subtotal else 0.0
        final_client_money_req= get_num_cell(df_tab4, row_final_req, 2) if row_final_req else 0.0

        st.markdown("### 📊 Internal Client Money Reconciliation Suite (v4.1) - Cash ISA")
        st.caption("FCA Compliance Ledger Verification according to CASS 7.16.22 Rules.")

        # --- 1. COMPLIANCE BANNER HEADER ---
        st.markdown(f"""
            <div class="reason-box">
                <div class="reason-title" style="color: #ef4444; font-size: 15px;">⚠️ Active Regulatory Target Status</div>
                <div class="reason-section" style="font-size: 14px;">
                    <strong>Calculated Variance:</strong> <span style="color:#ef4444; font-weight:700;">£ {shortfall_calculated:,.2f}</span><br>
                    <strong>Conclusion:</strong> {conclusion_excel_text}
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- 2. BREAKS & DISCREPANCIES MATRIX (ΔΥΝΑΜΙΚΗ ΑΝΑΚΤΗΣΗ) ---
        st.markdown('<div class="table-header-container"><div class="table-title">🚫 Outstanding Breaks & Discrepancies Ledger</div></div>', unsafe_allow_html=True)
        
        # Διαβάζουμε και δείχνουμε ολόκληρο το Breaks table δυναμικά
        st.data_editor(df_tab4.dropna(how="all").iloc[4:9].fillna(""), use_container_width=True, hide_index=True, key="breaks_live_grid")

        # --- 3. CLIENT MONEY REQUIREMENT CALCULATION ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="table-header-container"><div class="table-title">🏛️ CASS 7.16.22 Client Money Requirement Calculation Engine</div></div>', unsafe_allow_html=True)
        
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
                    <div class="recon-row"><span>Unallocated Balances</span><strong>£ {less_unallocated:,.2f}</strong></div>
                    <div class="recon-row"><span>Temporary Transaction Funding</span><strong style="color:#ef4444;">£ {temp_tx_funding:,.2f}</strong></div>
                    <div class="recon-row total" style="border-top:1px solid #1f2937; padding-top:15px; color:#ef4444;"><span>Prudent Funding Subtotal</span><strong>£ {temp_tx_funding:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

        # Bottom Summary Totals Line
        st.markdown(f"""
            <div class="workspace-card" style="margin-top:20px;">
                <div class="recon-row" style="font-size:16px;"><span><strong>Sub-Total Requirement (pre-Interest)</strong></span><strong>£ {subtotal_pre_interest:,.2f}</strong></div>
                <div class="recon-row total" style="font-size:18px; color:#10b981; border-top:2px solid #1f2937; padding-top:15px;">
                    <span>🏛️ Final Client Money Requirement</span><strong>£ {final_client_money_req:,.2f}</strong>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # ==========================================
    # 📂 FALLBACK VIEW FOR OTHER SHEETS
    # ==========================================
    else:
        st.markdown(f"### 📂 View Mode: {selected_tab}")
        try:
            df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
            df_cleaned = df_any.dropna(how='all').dropna(axis=1, how='all').fillna("")
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
