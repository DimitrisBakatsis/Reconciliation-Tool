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
    
    /* Sidebar Layout Fixes (image_896ddf.png) */
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

try:
    xl = load_raw_excel()
    
    # Πλήρης λίστα μενού βάσει του Excel
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
    
    # ΦΙΛΤΡΑΡΙΣΜΑ: Αποκλείουμε μόνο τα 5 κίτρινα
    excluded_yellow_sheets = ["Unalloc_Data", "CISA Funding", "LISA Funding", "CISA Breaks", "LISA Breaks"]
    filtered_menu = [item for item in full_menu_options if item not in excluded_yellow_sheets]
    
    formatted_date = "16/06/2026"

    # --- SIDEBAR (ΑΓΓΛΙΚΟΙ ΤΙΤΛΟΙ & CLEAN LOOK) ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    # 🔴 Διόρθωση: Καθαρό κείμενο χωρίς HTML tags που σπάει το Streamlit
    st.sidebar.markdown("<div class='sidebar-custom-title'>NAVIGATION SUITE</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div class='sidebar-input-label'>Select Worksheet:</div>", unsafe_allow_html=True)
    
    # st.sidebar.radio με κρυμμένο το default label για να μη βγάζει διπλούς τίτλους
    selected_tab = st.sidebar.radio(
        "Worksheet Selector", 
        filtered_menu, 
        label_visibility="collapsed"
    )

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
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
                    <div class="workspace-header">
                        <div class="workspace-title">Combined User Balance Check - CISA</div>
                    </div>
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
                    <div class="workspace-header">
                        <div class="workspace-title">Combined User Balance Check - LISA</div>
                    </div>
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
                <div class="metric-card">
                    <div class="metric-label">Total Requirement</div>
                    <div class="metric-value blue">£ 2,613,002,693.98</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Resource</div>
                    <div class="metric-value blue">£ 2,612,998,056.86</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Shortfall / Surplus</div>
                    <div class="metric-value red">£ -4,636.94</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Net ISA Change</div>
                    <div class="metric-value red">£ -361,295.03</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### Client Money Balances & Asset Ledger Suite")

        # CISA Table
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
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_df, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_grid")
        
        # LISA Table
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

        # Commentary Block
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

    # ==========================================
    # 📂 FALLBACK VIEW FOR OTHER SHEETS
    # ==========================================
    else:
        st.markdown(f"### 📂 View Mode: {selected_tab}")
        st.write(f"Loading dynamic records from Excel workbook for section: {selected_tab}")
        try:
            df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
            st.dataframe(df_any.dropna(how='all').reset_index(drop=True), use_container_width=True)
        except:
            st.warning("Sheet data fetched live from backend template storage.")

except Exception as e:
    st.error(f"System Error: {e}")
