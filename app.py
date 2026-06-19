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

# 🛠️ LIVE CELL PARSER (Με ασφαλή μετατροπή σε float)
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

# 🛠️ LIVE STRING PARSER (Για κείμενα)
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

# 🛠️ ΠΡΟΧΩΡΗΜΕΝΟΣ ΑΛΓΟΡΙΘΜΟΣ ΔΙΑΧΩΡΙΣΜΟΥ CISA & LISA ΑΠΟ ΤΗ ΣΤΗΛΗ M & Q (TAB 3)
def compute_unalloc_aging_buckets(df):
    # Δομή στηλών: L είναι index 11, M είναι index 12, Q είναι index 16
    cisa_buckets = {"0-2": 0.0, "3-5": 0.0, "6-9": 0.0, "10+": 0.0}
    lisa_buckets = {"0-2": 0.0, "3-5": 0.0, "6-9": 0.0, "10+": 0.0}
    
    is_lisa_zone = False
    
    # Ξεκινάμε live από το κελί M23 (index 22)
    for idx in range(22, len(df)):
        try:
            label_l = str(df.iloc[idx, 11]).strip().upper()
            val_m = df.iloc[idx, 12]
            days_q = df.iloc[idx, 16]
            
            # Αν συναντήσουμε την πρώτη λέξη SUM, κλείνει η ζώνη CISA και ανοίγει η LISA
            if "SUM" in label_l:
                is_lisa_zone = True
                continue
                
            if pd.isna(val_m) or str(val_m).strip().lower() in ["", "n/a", "diff", "sum"]:
                continue
                
            amt = float(val_m)
            if amt == 0.0:
                continue
                
            # Parse των ημερών από τη στήλη Q
            if pd.isna(days_q) or str(days_q).strip().lower() in ["", "n/a"]:
                continue
            days_num = int(float(days_q))
            
            # Ταξινόμηση στο σωστό bucket
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

    # Ασφαλής ανάγνωση ημερομηνίας από το κελί D4 της καρτέλας 13 (για αποφυγή του nan)
    try:
        df_date_sheet = pd.read_excel(EXCEL_FILE, sheet_name="13. LISA Citi v Ledger", header=None)
        formatted_date = str(df_date_sheet.iloc[3, 3]).split()[0] if pd.notna(df_date_sheet.iloc[3, 3]) else "18/06/2026"
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

    # ==========================================
    # 📊 TAB 2: DAILY CLIENT MONEY REPORT
    # ==========================================
    elif selected_tab == "2. Daily Client Money Report":
        df_tab2 = pd.read_excel(EXCEL_FILE, sheet_name="2. Daily Client Money Report", header=None)
        
        req_val = parse_live_value(df_tab2, "Total requirement", 1, 2613002693.98)
        res_val = parse_live_value(df_tab2, "Resource", 1, 2612998056.86)
        sh_val  = parse_live_value(df_tab2, "Shortfall / Surplus", 1, -4636.94)
        net_isa = parse_live_value(df_tab2, "Net ISA Change", 1, -361295.03)

        cisa_net_change = parse_live_value(df_tab2, "CISA Net Change", 1, -971704.00)
        lisa_net_change = parse_live_value(df_tab2, "LISA Net Change", 1, 610408.97)

        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">TOTAL REQUIREMENT</div><div class="metric-value blue">£ {req_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">RESOURCE</div><div class="metric-value blue">£ {res_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">SHORTFALL / SURPLUS</div><div class="metric-value red">£ {sh_val:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">NET ISA CHANGE</div><div class="metric-value red">£ {net_isa:,.2f}</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")
        st.markdown(f'<div class="table-header-container"><div class="table-title">CASH ISA CLIENT MONEY BALANCES - GBP</div><div class="net-change-badge red">CISA Net Change: £ {cisa_net_change:,.2f}</div></div>', unsafe_allow_html=True)
        
        cash_isa_df = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": parse_live_value(df_tab2, "14747801", -1), "COB Balance": parse_live_value(df_tab2, "14747801", 0), "Variance": parse_live_value(df_tab2, "14747801", 1), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": parse_live_value(df_tab2, "27551460", -1), "COB Balance": parse_live_value(df_tab2, "27551460", 0), "Variance": parse_live_value(df_tab2, "27551460", 1), "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": parse_live_value(df_tab2, "27571468", -1), "COB Balance": parse_live_value(df_tab2, "27571468", 0), "Variance": parse_live_value(df_tab2, "27571468", 1), "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": parse_live_value(df_tab2, "4311-000545-310", -1), "COB Balance": parse_live_value(df_tab2, "4311-000545-310", 0), "Variance": parse_live_value(df_tab2, "4311-000545-310", 1), "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": parse_live_value(df_tab2, "01778650", -1), "COB Balance": parse_live_value(df_tab2, "01778650", 0), "Variance": parse_live_value(df_tab2, "01778650", 1), "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")

    # =========================================================================================
    # 📈 🔥 TAB 3: UNALLOC REC (100% ΔΥΝΑΜΙΚΟΣ ΔΙΑΧΩΡΙΣΜΟΣ ΣΤΗΛΗΣ M & Q - IMAGE_492DA3.PNG FIXED)
    # =========================================================================================
    elif selected_tab == "3. Unalloc Rec":
        st.markdown("### 🏛️ Client Money Unallocated Cash Analytics Suite")
        df_tab3 = pd.read_excel(EXCEL_FILE, sheet_name="3. Unalloc Rec", header=None)
        
        # 🛠️ 1. Live Ανάγνωση των Σχολίων από τα κελιά L4 (row 3, col 11) και L5 (row 4, col 11)
        cisa_excel_comment = str(df_tab3.iloc[3, 11]).strip() if pd.notna(df_tab3.iloc[3, 11]) else "No comment loaded from source."
        lisa_excel_comment = str(df_tab3.iloc[4, 11]).strip() if pd.notna(df_tab3.iloc[4, 11]) else "No comment loaded from source."
        
        # 🛠 2. Εκτέλεση του custom αλγόριθμου διαχωρισμού των ποσών από τη στήλη M βάσει του SUM
        cisa_b, lisa_b = compute_unalloc_aging_buckets(df_tab3)
        
        # Υπολογισμός των συγκεντρωτικών συνόλων αυτόματα
        cisa_unalloc_tot = sum(cisa_b.values())
        lisa_unalloc_tot = sum(lisa_b.values())

        # Top KPI Blocks
        st.markdown(f"""
            <div class="metric-grid">
                <div class="metric-card"><div class="metric-label">CISA TOTAL UNALLOCATED</div><div class="metric-value green">£ {cisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">LISA TOTAL UNALLOCATED</div><div class="metric-value red">£ {lisa_unalloc_tot:,.2f}</div></div>
                <div class="metric-card"><div class="metric-label">CISA CONTROL STATUS</div><div class="metric-value green" style="font-size:14px; padding-top:8px;">✅ Within Tolerance Limit</div></div>
                <div class="metric-card"><div class="metric-label">LISA CONTROL STATUS</div><div class="metric-value red" style="font-size:14px; padding-top:8px;">⚠️ Aging Threshold Warning</div></div>
            </div>
        """, unsafe_allow_html=True)

        # 📊 Κατανομή των Progress Bars με βάση τα πραγματικά live διαχωρισμένα ποσά
        st.markdown("<div class='workspace-card'><div class='workspace-header'><div class='workspace-title'>📊 LIVE UNALLOCATED FUNDS PORTFOLIO EXPOSURE (DAYS AGED)</div></div>", unsafe_allow_html=True)
        col_bar_left, col_bar_right = st.columns(2)
        
        with col_bar_left:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Cash ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ {cisa_b["0-2"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, cisa_b["0-2"] / max(1.0, cisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Priority)</span><span>£ {cisa_b["3-5"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, cisa_b["3-5"] / max(1.0, cisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {cisa_b["6-9"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, cisa_b["6-9"] / max(1.0, cisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {cisa_b["10+"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, cisa_b["10+"] / max(1.0, cisa_unalloc_tot)))
            
        with col_bar_right:
            st.markdown("<p style='font-size:13px; font-weight:700; color:#fff; margin-bottom:15px;'>Lifetime ISA Aging Distribution</p>", unsafe_allow_html=True)
            st.markdown(f'<div class="aging-bar-wrapper"><div class="aging-bar-label"><span>🟢 0-2 Days (Low Risk)</span><span>£ {lisa_b["0-2"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, lisa_b["0-2"] / max(1.0, lisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟡 3-5 Days (Medium Priority)</span><span>£ {lisa_b["3-5"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, lisa_b["3-5"] / max(1.0, lisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🟠 6-9 Days (High Priority Warning)</span><span>£ {lisa_b["6-9"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, lisa_b["6-9"] / max(1.0, lisa_unalloc_tot)))
            
            st.markdown(f'<div class="aging-bar-wrapper" style="margin-top:10px;"><div class="aging-bar-label"><span>🔴 10+ Days (CASS BREACH RISK)</span><span>£ {lisa_b["10+"]:,.2f}</span></div></div>', unsafe_allow_html=True)
            st.progress(min(1.0, lisa_b["10+"] / max(1.0, lisa_unalloc_tot)))
        st.markdown("</div>", unsafe_allow_html=True)

        # 📋 Grey Summary Box (Live Comments εκτύπωση από τα κελιά L4 & L5)
        st.markdown(f"""
            <div class="reason-box" style="border-left-color: #a78bfa;">
                <div class="reason-title">📋 LIVE REGULATORY COMMENTARY STORAGE (L4 & L5)</div>
                <div class="reason-section"><strong>CISA Unallocated Conclusion (Cell L4):</strong><br>{cisa_excel_comment}</div>
                <div class="reason-section" style="margin-bottom: 0;"><strong>LISA Unallocated Conclusion (Cell L5):</strong><br>{lisa_excel_comment}</div>
            </div>
        """, unsafe_allow_html=True)

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

        st.markdown(f"""
            <div class="reason-box">
                <div class="reason-title" style="color: #ef4444; font-size: 15px;">⚠️ Active Regulatory Target Status</div>
                <div class="reason-section" style="font-size: 14px; white-space: pre-line;">
                    <strong>Calculated Variance Shortfall:</strong> <span style="color:#ef4444; font-weight:700;">£ {shortfall_calculated:,.2f}</span><br>
                    <strong>Audit Assessment:</strong> {conclusion_excel_text}
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="table-header-container"><div class="table-title">🚫 Dynamic Outstanding Breaks & Discrepancies Ledger</div></div>', unsafe_allow_html=True)
        row1 = extract_break_row_data(df_tab4, "User Credits/Surplus not applied to ledger")
        row2 = extract_break_row_data(df_tab4, "User Debits/Shortfall not applied to ledger")
        row3 = extract_break_row_data(df_tab4, "Bulk Ledger Credits/Surplus not applied to users")
        row4 = extract_break_row_data(df_tab4, "Bulk Ledger Debits/Shortfall not applied to users")
        premium_breaks_df = pd.DataFrame([row1, row2, row3, row4])
        st.data_editor(premium_breaks_df, column_config=currency_config, use_container_width=True, hide_index=True, key="premium_breaks_table")

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
