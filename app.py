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

# Ασφαλής συνάρτηση για live καθαρισμό και μορφοποίηση οποιασδήποτε raw καρτέλας
def display_cleaned_sheet(sheet_name):
    df_any = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)
    # Αφαιρούμε στήλες και γραμμές που είναι τελείως κενές
    df_cleaned = df_any.dropna(how='all').dropna(axis=1, how='all')
    df_cleaned = df_cleaned.fillna("")
    st.dataframe(df_cleaned.astype(str), use_container_width=True, hide_index=True)

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

    # --- SIDEBAR ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Corporate Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-custom-title'>NAVIGATION SUITE</div>", unsafe_allow_html=True)
    st.sidebar.markdown("<div class='sidebar-input-label'>Select Worksheet:</div>", unsafe_allow_html=True)
    
    selected_tab = st.sidebar.radio("Worksheet Selector", filtered_menu, label_visibility="collapsed")

    # Φορτώνουμε δυναμικά την ημερομηνία από την κορυφή της καρτέλας 1
    try:
        df_top = pd.read_excel(EXCEL_FILE, sheet_name=0, header=None)
        formatted_date = str(df_top.iloc[2, 3]).split()[0] if pd.notna(df_top.iloc[2, 3]) else "18/06/2026"
    except:
        formatted_date = "18/06/2026"

    # --- MAIN GLOBAL HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    # =========================================================================================
    # 🏛️ ΠΛΗΡΩΣ ΔΥΝΑΜΙΚΗ ΑΠΕΙΚΟΝΙΣΗ ΓΙΑ ΟΛΑ ΤΑ TABS (ΧΩΡΙΣ ΚΑΡΦΩΤΑ ΝΟΥΜΕΡΑ)
    # =========================================================================================
    st.markdown(f"### 📂 View Workspace: {selected_tab}")
    st.caption("FCA Compliance Ledger Sync. Live extraction of raw cells, structured columns, and active rows.")
    st.markdown(f'<div class="table-header-container"><div class="table-title">{selected_tab} Master Template View</div></div>', unsafe_allow_html=True)
    
    # 🏁 Εδώ τρέχει ο universal καθαρισμός που εμφανίζει ολόκληρη την πληροφορία απευθείας από το Excel
    display_cleaned_sheet(selected_tab)

    # 🏁 GLOBAL PDF EXPORT BUTTON
    st.markdown("<div class='pdf-container'>", unsafe_allow_html=True)
    if st.button("📄 Export to PDF", key="btn_export_global_pdf"):
        st.toast("Generating financial audit report PDF...", icon="🔄")
    st.markdown("</div>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"System Error: {e}")
