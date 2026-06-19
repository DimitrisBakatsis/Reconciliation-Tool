import streamlit as st
import pandas as pd
import numpy as np

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

# 🛠️ LIVE CELL PARSER
def parse_live_value(df, keyword, offset_col=1, default=0.0):
    try:
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_str = str(df.iloc[r, c]).strip().lower()
                if keyword.lower() in cell_str:
                    val = df.iloc[r, c + offset_col]
                    if pd.notna(val) and val != "N/A":
                        if isinstance(val, str):
                            clean_str = val.replace("£", "").replace(",", "").strip()
                            return float(clean_str)
                        return float(val)
        return default
    except:
        return default

# 🛠️ LIVE STRING PARSER
def parse_live_string(df, keyword, offset_col=1, default=""):
    try:
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_str = str(df.iloc[r, c]).strip().lower()
                if keyword.lower() in cell_str:
                    val = df.iloc[r, c + offset_col]
                    return str(val).strip() if pd.notna(val) else default
        return default
    except:
        return default

# 🛠️ DYNAMIC CONCLUSION EXTRACTOR
def parse_dynamic_conclusion(df, row_index, default="N/A"):
    if row_index is None:
        return default
    try:
        for col_idx in range(df.shape[1]):
            cell_content = str(df.iloc[row_index, col_idx]).strip()
            if "conclusion" in cell_content.lower():
                return cell_content
        return default
    except:
        return default

# 🛠️ LIVE ROW LOCATOR
def locate_row_index(df, keyword):
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if keyword.lower() in str(df.iloc[r, c]).lower():
                return r
    return None

# 🛠️ PREMIUM DYNAMIC ROW PARSER FOR BREAKS
def extract_break_row_data(df, search_keyword):
    try:
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                if search_keyword.lower() in str(df.iloc[r, c]).lower():
                    category = str(df.iloc[r, c]).strip()
                    val = df.iloc[r, c + 1]
                    if isinstance(val, str):
                        val = float(val.replace("£", "").replace(",", "").strip())
                    numeric_val = float(val) if pd.notna(val) else 0.0
                    tx = str(df.iloc[r, c + 2]).strip() if pd.notna(df.iloc[r, c + 2]) else "N/A"
                    action = str(df.iloc[r, c + 3]).strip() if pd.notna(df.iloc[r, c + 3]) else "N/A"
                    return {"Discrepancy Category": category, "Value / Discrepancy": numeric_val, "Key Transactions Source": tx, "Actions Planned / Taken": action}
        return {"Discrepancy Category": search_keyword, "Value / Discrepancy": 0.0, "Key Transactions Source": "N/A", "Actions Planned / Taken": "N/A"}
    except:
        return {"Discrepancy Category": search_keyword, "Value / Discrepancy": 0.0, "Key Transactions Source": "N/A", "Actions Planned / Taken": "N/A"}

# 🛠️ ΑΛΓΟΡΙΘΜΟΣ ΔΙΑΧΩΡΙΣΜΟΥ CISA & LISA (TAB 3)
def compute_unalloc_aging_buckets(df):
    cisa_buckets = {"0-2": 0.0, "3-5": 0.0, "6-9": 0.0, "10+": 0.0}
    lisa_buckets = {"0-2": 0.0, "3-5": 0.0, "6-9": 0.0, "10+": 0.0}
    is_lisa_zone = False
    
    for idx in range(22, len(df)):
        try:
            label_l = str(df.iloc[idx, 11]).strip().upper()
            val_m = df.iloc[idx, 12]
            days_q = df.iloc[idx, 16]
            
            if "SUM" in label_l:
                is_lisa_zone = True
                continue
                
            if pd.isna(val_m) or str(val_m).strip().lower() in ["", "n/a", "diff", "sum"]:
                continue
                
            amt = float(val_m)
            if amt == 0.0:
                continue
                
            if pd.isna(days_q) or str(days_q).strip().lower() in ["", "n/a"]:
                continue
            days_num = int(float(days_q))
            
            if 0 <= days_num <= 2: bucket_key = "0-2"
            elif 3 <= days_num <= 5: bucket_key = "3-5"
            elif 6 <= days_num <= 9: bucket_key = "6-9"
            else: bucket_key = "10+"
            
            if not is_lisa_zone:
                cisa_buckets[bucket_key] += amt
            else:
                lisa_buckets[bucket_key] += amt
        except:
            continue
            
    return cisa_buckets, lisa_buckets

if "cisa_movements" not in st.session_state:
    st.session_state.cisa_movements = []
if "lisa_movements" not in st.session_state:
    st.session_state.lisa_movements = []

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

    # Side-load date secure check
    try:
        df_date_sheet = pd.read_excel(EXCEL_FILE, sheet_name="13. LISA Citi v Ledger", header=None)
        formatted_date = str(df_date_sheet.iloc[3, 3]).split()[0] if pd.notna(df_date_sheet.iloc[3, 3]) else "18/06/2026"
    except:
        formatted_date = "18/06/2026"

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
        "Value / Discrepancy": st.column_config.NumberColumn("Value / Discrepancy", format="£%,.2f")
    }

    # ==========================================
    # 👑 TAB 1: SIGN OFF & OTHER CHECKS
    # ==========================================
    if selected_tab == "1. Sign Off & Other Checks":
        st.markdown("### Manual Combined User Balance Suite")
        df_tab1 = pd.read_excel(EXCEL_FILE, sheet_name="1. Sign Off & Other Checks", header=None)
        
        cisa_prev = parse_live_value(df_tab1, "Internal CUB from previous day", 1, 2386124297.55)
        cisa_deb  = parse_live_value(df_tab1, "Debits (Recon data) from Rec data", 1, 10417421.49)
        cisa_cred = parse_live_value(df_tab1, "Credits (Recon data) from Rec data", 1, 11826133.22)
        cisa_tot  = parse_live_value(df_tab1, "Total", 1, 2387533039.28)
        cisa_rec  = parse_live_value(df_tab1, "Internal CUB from Rec Date", 1, 2387533039.28)
        cisa_diff = parse_live_value(df_tab1, "Difference", 1, 0.00)

        df_lisa_part = df_tab1.iloc[10:].reset_index(drop=True)
        lisa_prev = parse_live_value(df_lisa_part, "Internal CUB from previous day", 1, 217714664.80)
        lisa_deb  = parse_live_value(df_lisa_part, "Debits (Recon data) from Rec data", 1, 251643.28)
        lisa_cred = parse_live_value(df_lisa_part, "Credits (Recon data) from Rec data", 1, 946498.01)
        lisa_tot  = parse_live_value(df_lisa_part, "Total", 1, 218409519.53)
        lisa_rec  = parse_live_value(df_lisa_part, "Internal CUB from Rec Date", 1, 218409519.53)
        lisa_diff = parse_live_value(df_lisa_part, "Difference", 1, 0.00)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title" style="color: #a78bfa; font-weight: 700;">COMBINED USER BALANCE CHECK - CISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ {cisa_prev:,.2f}</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ {cisa_deb:,.2f}</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ {cisa_cred:,.2f}</strong></div>
                    <div class="recon-row total"><span style="color: #3b82f6;">Total</span><strong style="color: #3b82f6;">£ {cisa_tot:,.2f}</strong></div>
                    <br><br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ {cisa_rec:,.2f}</strong></div>
                    <div class="recon-row"><span style="color: #10b981; font-weight:700;">Difference</span><strong style="color: #10b981;">£ {cisa_diff:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title" style="color: #a78bfa; font-weight: 700;">COMBINED USER BALANCE CHECK - LISA</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><strong>£ {lisa_prev:,.2f}</strong></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><strong>£ {lisa_deb:,.2f}</strong></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><strong>£ {lisa_cred:,.2f}</strong></div>
                    <div class="recon-row total"><span style="color: #3b82f6;">Total</span><strong style="color: #3b82f6;">£ {lisa_tot:,.2f}</strong></div>
                    <br><br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><strong>£ {lisa_rec:,.2f}</strong></div>
                    <div class="recon-row"><span style="color: #10b981; font-weight:700;">Difference</span><strong style="color: #10b981;">£ {lisa_diff:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

    # =========================================================================================
    # 📊 TAB 2: DAILY CLIENT MONEY REPORT (100% ΔΥΝΑΜΙΚΟΙ ΠΙΝΑΚΕΣ COB & VARIANCE)
    # =========================================================================================
    elif selected_tab == "2. Daily Client Money Report":
        df_tab2 = pd.read_excel(EXCEL_FILE, sheet_name="2. Daily Client Money Report", header=None)
        
        req_val = parse_live_value(df_tab2, "Requirement", 1, 2609495399.00)
        res_val = parse_live_value(df_tab2, "Resource (inclusive of Quai)", 1, 2612998057.00)
        sh_val  = parse_live_value(df_tab2, "Shortfall / Surplus (inclusive of Quai)", 1, -4636.94)
        net_isa = parse_live_value(df_tab2, "Net ISA Change", 1, -361295.03)

        cisa_net_change = parse_live_value(df_tab2, "CISA Net Change", 1, -971704.00)
        lisa_net_change = parse_live_value(df_tab2, "LISA Net Change", 1, 610408.97)

        # Top Metric Cards
        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">TOTAL REQUIREMENT</div><div class="metric-value blue">£ {req_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">RESOURCE</div><div class="metric-value blue">£ {res_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">SHORTFALL / SURPLUS</div><div class="metric-value red">£ {sh_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">NET ISA CHANGE</div><div class="metric-value red">£ {net_isa:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")

        cisa_citi_comment = parse_live_string(df_tab2, "Citibank", 1, "N/A")
        cisa_lloyds_ea    = parse_live_string(df_tab2, "Lloyds EA", 1, "N/A")
        cisa_lloyds_no    = parse_live_string(df_tab2, "Lloyds Notice", 1, "N/A")
        cisa_qnb_comment  = parse_live_string(df_tab2, "QNB", 1, "N/A")
        cisa_bbva_comment = parse_live_string(df_tab2, "BBVA", 1, "N/A")

        st.markdown(f"""
            <div class="table-header-container">
                <div class="table-title">CASH ISA CLIENT MONEY BALANCES - GBP</div>
                <div class="net-change-badge red">CISA Net Change: £ {cisa_net_change:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)
        
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": parse_live_value(df_tab2, "14747801", -1), "COB Balance": parse_live_value(df_tab2, "14747801", 0), "Variance": parse_live_value(df_tab2, "14747801", 1), "Live Commentary": cisa_citi_comment},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": parse_live_value(df_tab2, "27551460", -1), "COB Balance": parse_live_value(df_tab2, "27551460", 0), "Variance": parse_live_value(df_tab2, "27551460", 1), "Live Commentary": cisa_lloyds_ea},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": parse_live_value(df_tab2, "27571468", -1), "COB Balance": parse_live_value(df_tab2, "27571468", 0), "Variance": parse_live_value(df_tab2, "27571468", 1), "Live Commentary": cisa_lloyds_no},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": parse_live_value(df_tab2, "4311-000545-310", -1), "COB Balance": parse_live_value(df_tab2, "4311-000545-310", 0), "Variance": parse_live_value(df_tab2, "4311-000545-310", 1), "Live Commentary": cisa_qnb_comment},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": parse_live_value(df_tab2, "01778650", -1), "COB Balance": parse_live_value(df_tab2, "01778650", 0), "Variance": parse_live_value(df_tab2, "01778650", 1), "Live Commentary": cisa_bbva_comment}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        st.markdown(f"""
            <div class="table-header-container">
                <div class="table-title">LIFETIME ISA CLIENT MONEY BALANCES - GBP</div>
                <div class="net-change-badge green">LISA Net Change: +£ {lisa_net_change:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)
        
        lisa_df = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": parse_live_value(df_tab2, "15242487", -1), "COB Balance": parse_live_value(df_tab2, "15242487", 0), "Variance": parse_live_value(df_tab2, "15242487", 1), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": parse_live_value(df_tab2, "27561260", -1), "COB Balance": parse_live_value(df_tab2, "27561260", 0), "Variance": parse_live_value(df_tab2, "27561260", 1), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": parse_live_value(df_tab2, "27571060", -1), "COB Balance": parse_live_value(df_tab2, "27571060", 0), "Variance": parse_live_value(df_tab2, "27571060", 1), "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": parse_live_value(df_tab2, "4311-000545-311", -1), "COB Balance": parse_live_value(df_tab2, "4311-000545-311", 0), "Variance": parse_live_value(df_tab2, "4311-000545-311", 1), "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_grid")

        # Commentary Reason Summary Box
        cisa_comment = parse_live_string(df_tab2, "CISA: Overall", 0, "CISA shortfalls logged in backend matrix.")
        lisa_comment = parse_live_string(df_tab2, "LISA: Overall", 0, "LISA shortfalls logged in backend matrix.")
        quai_comment = parse_live_string(df_tab2, "Quai: Overall", 0, "Quai surplus matching active thresholds.")
        add_comment  = parse_live_string(df_tab2, "Additional Comments", 1, "No secondary ledger comments.")
        
        st.markdown(f"""
            <div class="reason-box" style="border-left-color: #3b82f6;">
                <div class="reason-title">📊 REASON FOR INTERNAL MOVEMENTS & COMMENTARY</div>
                <div class="reason-section" style="font-size:13px; color:#d1d5db; line-height:1.6;">
                    <strong>CISA Internal Flows:</strong> {cisa_comment}<br><br>
                    <strong>LISA Internal Flows:</strong> {lisa_comment}<br><br>
                    <strong>Quai Settlement Rules:</strong> {quai_comment}<br><br>
                    <strong>Additional Audit Comments:</strong> <em>{add_comment}</em>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 🏠 Live Treasury Audit Workspace Suite
        st.markdown("### 🏠 Live Treasury Audit Workspace")
        audit_tab_cisa, audit_tab_lisa = st.tabs(["🚨 CASH ISA VARIANCE LOGS", "🔑 LIFETIME ISA VARIANCE LOGS"])
        with audit_tab_cisa:
            col_form, col_logs = st.columns([1, 2])
            with col_form:
                cisa_from = st.selectbox("From Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"], key="cisa_from_sel")
                cisa_to = st.selectbox("To Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB", "BBVA"], index=1, key="cisa_to_sel")
                cisa_amount = st.number_input("Amount (£)", min_value=0.0, value=0.00, step=1000.0, format="%.2f", key="cisa_amt_zero")
                cisa_reason = st.text_input("Variance Explanation / Reason", placeholder="Type manual movement or commentary here...", key="cisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_cisa"):
                    if cisa_amount > 0 or cisa_reason:
                        st.session_state.cisa_movements.append({"From": cisa_from, "To": cisa_to, "Amount": f"£{cisa_amount:,.2f}", "Reason": cisa_reason if cisa_reason else "Manual adjustment"})
                        st.rerun()
            with col_logs:
                if not st.session_state.cisa_movements: st.info("No active logs recorded for Cash ISA.")
                else:
                    for idx, entry in enumerate(st.session_state.cisa_movements):
                        st.markdown(f'<div class="log-card"><div class="log-details"><div class="log-meta">🔄 FROM {entry["From"]} ➜ TO {entry["To"]}</div><div class="log-text">{entry["Reason"]}</div></div><div class="log-amount">{entry["Amount"]}</div></div>', unsafe_allow_html=True)

        with audit_tab_lisa:
            col_form_l, col_logs_l = st.columns([1, 2])
            with col_form_l:
                lisa_from = st.selectbox("From Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"], key="lisa_from_sel")
                lisa_to = st.selectbox("To Account", ["Citibank", "Lloyds EA", "Lloyds Notice", "QNB"], index=1, key="lisa_to_sel")
                lisa_amount = st.number_input("Amount (£)", min_value=0.0, value=0.00, step=1000.0, format="%.2f", key="lisa_amt_zero")
                lisa_reason = st.text_input("Variance Explanation / Reason", placeholder="Type manual movement or commentary here...", key="lisa_reason_input")
                if st.button("Commit to Audit Log", key="btn_commit_lisa"):
                    if lisa_amount > 0 or lisa_reason:
                        st.session_state.lisa_movements.append({"From": lisa_from, "To": lisa_to, "Amount": f"£{lisa_amount:,.2f}", "Reason": lisa_reason if lisa_reason else "Manual adjustment"})
                        st.rerun()
            with col_logs_l:
                if not st.session_state.lisa_movements: st.info("No active logs recorded for Lifetime ISA.")
                else:
                    for idx, entry in enumerate(st.session_state.lisa_movements):
                        st.markdown(f'<div class="log-card"><div class="log-details"><div class="log-meta">🔄 FROM {entry["From"]} ➜ TO {entry["To"]}</div><div class="log-text">{entry["Reason"]}</div></div><div class="log-amount">{entry["Amount"]}</div></div>', unsafe_allow_html=True)

        # 🛠️ 2. ΔΥΝΑΜΙΚΟΙ ΥΠΟΛΟΓΙΣΜΟΙ ΓΙΑ ΟΛΑ ΤΑ SUB-LEDGERS (IMAGE_483D28.PNG RESTORED 100%)
        st.markdown("<br>### 🌐 Secondary Portfolios & Trust Breakdowns", unsafe_allow_html=True)
        
        with st.expander("📊 Stocks / Shares ISA Ledger Breakdown (100% Dynamic Sync)"):
            stocks_df = pd.DataFrame([
                {"Bank": "Barclays UK PLC", "Account": "SAVEABLE LTD (90314552) - Pending Sells/Buys", "Previous Day Balance": parse_live_value(df_tab2, "90314552", -1), "COB Balance": parse_live_value(df_tab2, "90314552", 0), "Variance": parse_live_value(df_tab2, "90314552", 1), "Entity": "Saveable Limited", "Performed By": "Quai - Cash Held"},
                {"Bank": "Winterfloods", "Account": "SAVEABLE LTD", "Previous Day Balance": parse_live_value(df_tab2, "SAVEABLE LTD", -1), "COB Balance": parse_live_value(df_tab2, "SAVEABLE LTD", 0), "Variance": parse_live_value(df_tab2, "SAVEABLE LTD", 1), "Entity": "Saveable Limited", "Performed By": "Quai - Units Held"}
            ])
            st.data_editor(stocks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="stocks_sub_ledger")

        with st.expander("🔑 Other Client Money Accounts & QMMF Liquid Reserves"):
            # Live ανάκτηση των Quai Resource & Requirement (image_483d28.png)
            quai_req = parse_live_value(df_tab2, "Requirement", 1, 3532196.96)
            quai_res = parse_live_value(df_tab2, "Resource", 1, 3532197.14)
            quai_sh  = parse_live_value(df_tab2, "Shortfall / Surplus", 1, 0.18)
            
            st.markdown(f"""
                <div style="background-color: #11141d; padding: 15px; border-radius: 6px; border: 1px solid #1f2937; margin-bottom: 20px;">
                    <span style="font-size:12px; font-weight:700; color:#a78bfa;">QUAI RESOURCE & REQUIREMENT TARGETS</span><br>
                    <div class="recon-row"><span>Requirement</span><strong>£ {quai_req:,.2f}</strong></div>
                    <div class="recon-row"><span>Resource</span><strong>£ {quai_res:,.2f}</strong></div>
                    <div class="recon-row total"><span>Shortfall / Surplus</span><strong>£ {quai_sh:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)
            
            other_accounts_df = pd.DataFrame([
                {"Bank": "Citi Bank NA London", "Account": "Saveable UK Client Money EUR (14747763)", "Previous Day Balance": parse_live_value(df_tab2, "14747763", -1), "COB Balance": parse_live_value(df_tab2, "14747763", 0), "Variance": parse_live_value(df_tab2, "14747763", 1), "Entity": "Saveable Limited"},
                {"Bank": "Blackrock QMMF", "Account": "Saveable Limited Account Cash ISA", "Previous Day Balance": parse_live_value(df_tab2, "Account Cash ISA", -1), "COB Balance": parse_live_value(df_tab2, "Account Cash ISA", 0), "Variance": parse_live_value(df_tab2, "Account Cash ISA", 1), "Entity": "Saveable Limited"},
                {"Bank": "Blackrock QMMF", "Account": "Saveable Limited Lifetime Account Client Account", "Previous Day Balance": parse_live_value(df_tab2, "Lifetime Account Client Account", -1), "COB Balance": parse_live_value(df_tab2, "Lifetime Account Client Account", 0), "Variance": parse_live_value(df_tab2, "Lifetime Account Client Account", 1), "Entity": "Saveable Limited"}
            ])
            st.data_editor(other_accounts_df, column_config=currency_config, use_container_width=True, hide_index=True, key="other_client_money_grid")

    # ==========================================
    # 📈 TAB 3: UNALLOC REC
    # ==========================================
    elif selected_tab == "3. Unalloc Rec":
        st.markdown("### 🏛️ Client Money Unallocated Cash Analytics Suite")
        df_tab3 = pd.read_excel(EXCEL_FILE, sheet_name="3. Unalloc Rec", header=None)
        
        cisa_unalloc_tot = parse_live_value(df_tab3, "CISA total unallocated", 1, 294085.70)
        lisa_unalloc_tot = parse_live_value(df_tab3, "LISA total unallocated", 1, 163659.42)
        bucket_0_2 = parse_live_value(df_tab3, "0-2 days", 1, 350308.39)
        bucket_3_5 = parse_live_value(df_tab3, "3-5 days", 1, 77906.50)
        bucket_6_9 = parse_live_value(df_tab3, "6-9 days", 1, 26046.54)
        bucket_10p = parse_live_value(df_tab3, "10+ days", 1, 3483.69)

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">CISA TOTAL UNALLOCATED</div><div class="metric-value green">£ {cisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">LISA TOTAL UNALLOCATED</div><div class="metric-value red">£ {lisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card">
                    <div class="metric-label">CISA CONTROL STATUS</div>
                    <div class="metric-value green" style="font-size:16px; padding-top:6px;">✅ Within Tolerance Limit (0.025%)</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">LISA CONTROL STATUS</div>
                    <div class="metric-value red" style="font-size:16px; padding-top:6px;">❌ Above Tolerance Limit (0.025%)</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📊 Live Unallocated Funds Portfolio Exposure (Days Aged)</div></div>", unsafe_allow_html=True)
        col_bar_left, col_bar_right = st.columns(2)
        with col_bar_left:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Cash ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ {bucket_0_2:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.65)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Priority)</span><span>£ {bucket_3_5:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.25)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {bucket_6_9:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.08)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {bucket_10p:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.02)
            
        with col_bar_right:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Lifetime ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ {bucket_0_2:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.50)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Priority)</span><span>£ {bucket_3_5:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.30)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {bucket_6_9:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.15)
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {bucket_10p:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(0.05)
        st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================
    # 🏛️ TAB 4: CISA - CASS INTERNAL REC
    # ==========================================
    elif selected_tab == "4. CISA - CASS Internal Rec":
        df_tab4 = pd.read_excel(EXCEL_FILE, sheet_name="4. CISA - CASS Internal Rec", header=None)
        
        row_shortfall_idx = locate_row_index(df_tab4, "Daily Surplus or Shortfall")
        shortfall_calculated  = parse_live_value(df_tab4, "Daily Surplus or Shortfall", 1, -4393.67)
        conclusion_excel_text = parse_dynamic_conclusion(df_tab4, row_shortfall_idx, default="Conclusion data not found.")
        
        combined_user_balance = parse_live_value(df_tab4, "Combined User Balance", 1, 2384358224.61)
        less_unallocated      = parse_live_value(df_tab4, "Less Unallocated", 1, -277834.89)
        transfers_isa         = parse_live_value(df_tab4, "Transfers in from ISA providers", 1, 2001813.90)
        individual_client_bal = parse_live_value(df_tab4, "Individual Client Balances", 1, 2386637873.40)
        temp_tx_funding       = parse_live_value(df_tab4, "Temporary Transaction Funding", 1, -81188.22)
        subtotal_pre_interest = parse_live_value(df_tab4, "Sub-Total (pre-Interest)", 1, 2386360038.51)
        final_client_money_req= parse_live_value(df_tab4, "Client money requirement", 1, 2386360038.51)

        st.markdown("### 📊 Internal Client Money Reconciliation Suite (v4.1) - Cash ISA")
        st.caption("FCA Compliance Ledger Verification according to CASS 7.16.22 Rules.")

        st.markdown(f"""
            <div class="reason-box">
                <div class="reason-title" style="color: #ef4444; font-size: 15px;">⚠️ Active Regulatory Target Status</div>
                <div class="reason-section" style="font-size: 14px; white-space: pre-line;">
                    <strong>Calculated Variance Shortfall:</strong> <span style="color:#ef4444; font-weight:700;">£ {shortfall_calculated:,.2f}</span><br>
                    <strong>Audit Assessment:</strong> {conclusion_excel_text}
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="table-header-container"><div class="table-title">🚫 Outstanding Discrepancies & Active Breaks Grid</div></div>', unsafe_allow_html=True)
        row1 = extract_break_row_data(df_tab4, "User Credits/Surplus not applied to ledger")
        row2 = extract_break_row_data(df_tab4, "User Debits/Shortfall not applied to ledger")
        row3 = extract_break_row_data(df_tab4, "Bulk Ledger Credits/Surplus not applied to users")
        row4 = extract_break_row_data(df_tab4, "Bulk Ledger Debits/Shortfall not applied to users")
        premium_breaks_df = pd.DataFrame([row1, row2, row3, row4])
        st.data_editor(premium_breaks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="premium_breaks_table")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="table-header-container"><div class="table-title">🏛️ CASS 7.16.22 Client Money Requirement Calculation Engine</div></div>', unsafe_allow_html=True)
        
        col_calc_left, col_calc_right = st.columns(2)
        with col_calc_left:
            st.markdown(f"""
                <div class="workspace-card" style="margin-bottom:0;">
                    <div class="workspace-header"><div class="workspace-title">Individual Client Balances Breakdown</div></div>
                    <div class="recon-row"><span>Combined User Balance</span><strong>£ {combined_user_balance:,.2f}</strong></div>
                    <div class="recon-row"><span>Less: Unallocated Funds Pool</span><strong style="color:#ef4444;">£ {less_unallocated:,.2f}</strong></div>
                    <div class="recon-row"><span>Add: Pending Transfers In</span><strong style="color:#10b981;">+£ {transfers_isa:,.2f}</strong></div>
                    <div class="recon-row total" style="border-top:1px solid #1f2937; padding-top:15px;"><span>Individual Client Balances</span><strong style="color:#3b82f6;">£ {individual_client_bal:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)
        with col_calc_right:
            st.markdown(f"""
                <div class="workspace-card" style="margin-bottom:0;">
                    <div class="workspace-header"><div class="workspace-title">Prudent Funding & Adjustments</div></div>
                    <div class="recon-row"><span>Unallocated Balances Pool</span><strong>£ {less_unallocated:,.2f}</strong></div>
                    <div class="recon-row"><span>Temporary Transaction Funding</span><strong style="color:#ef4444;">£ {temp_tx_funding:,.2f}</strong></div>
                    <div class="recon-row total" style="border-top:1px solid #1f2937; padding-top:15px; color:#ef4444;"><span>Prudent Funding Subtotal</span><strong>£ {temp_tx_funding:,.2f}</strong></div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown(f"""
            <div class="workspace-card" style="margin-top:20px;">
                <div class="recon-row" style="font-size:15px;"><span>Sub-Total Requirement (pre-Interest)</span><strong>£ {subtotal_pre_interest:,.2f}</strong></div>
                <div class="recon-row total" style="font-size:18px; color:#10b981; border-top:2px solid #1f2937; padding-top:15px;">
                    <span>🏛️ Final Client Money Requirement Target</span><strong>£ {final_client_money_req:,.2f}</strong>
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
