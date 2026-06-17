import streamlit as st
import pandas as pd

# 1. Page Config & Enterprise Reconciliation UI Styling
st.set_page_config(page_title="CASS Reconciliation Hub", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0b0c10; color: #e5e7eb; font-family: 'Inter', sans-serif; }
    
    /* Global Enterprise Header */
    .main-header { font-size: 26px; font-weight: 700; color: #ffffff; letter-spacing: -0.5px; }
    .date-subheader { font-size: 13px; color: #9ca3af; margin-bottom: 25px; display: flex; align-items: center; gap: 6px; }
    
    /* Workspace Containers (Enterprise Look) */
    .workspace-card { background-color: #11131c; border: 1px solid #1f2937; border-radius: 8px; padding: 20px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3); }
    .workspace-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #1f2937; padding-bottom: 12px; margin-bottom: 16px; }
    .workspace-title { font-size: 15px; font-weight: 600; color: #f3f4f6; text-transform: uppercase; letter-spacing: 0.5px; }
    
    /* System Status Badges (Duco/BlackLine style) */
    .badge-matched { background-color: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; text-transform: uppercase; }
    .badge-variance { background-color: rgba(239, 68, 68, 0.12); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; text-transform: uppercase; }
    
    /* Financial Summary Rows */
    .recon-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #1f2937; font-size: 14px; color: #d1d5db; }
    .recon-row.total { border-bottom: none; font-weight: 700; font-size: 15px; color: #3b82f6; padding-top: 14px; }
    
    /* Data Editor Overrides για να δένει με το Enterprise Dark UI */
    div[data-testid="stDataEditor"] { border: 1px solid #1f2937 !important; border-radius: 6px !important; overflow: hidden; }
    
    /* Sidebar Navigation Customization */
    [data-testid="stSidebar"] { background-color: #0d0f16; border-right: 1px solid #1f2937; }
    .sidebar-section-title { font-size: 11px; font-weight: 700; color: #4b5563; letter-spacing: 1px; margin-top: 20px; margin-bottom: 8px; text-transform: uppercase; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    
    # 📁 1. Δυναμικό Διάβασμα Ημερομηνίας (Tab 13 -> Κελί D4)
    try:
        df_tab13 = pd.read_excel(EXCEL_FILE, sheet_name=12, header=None)
        raw_date = df_tab13.iloc[3, 3]
        formatted_date = str(raw_date).split()[0] if pd.notna(raw_date) else "16/06/2026"
    except:
        formatted_date = "16/06/2026"

    # --- SIDEBAR NAVIGATION (ΕΠΑΝΕΦΕΡΑ ΚΑΙ ΤΙΣ 15 ΚΑΡΤΕΛΕΣ) ---
    st.sidebar.markdown("<div style='padding-top: 10px;'><span style='font-size: 16px; font-weight: 700; color: #fff;'>CASS Recon Portal</span></div>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<div class='sidebar-section-title'>Excel Worksheets</div>", unsafe_allow_html=True)
    
    # Εδώ είναι όλες οι σελίδες σου διαθέσιμες ξανά
    selected_tab = st.sidebar.radio("Καρτέλες:", sheet_names)

    # --- MAIN TOP HEADER ---
    st.markdown("<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'><span>📅</span> Close of Business Date: <strong>{formatted_date}</strong></div>", unsafe_allow_html=True)

    # Column configuration για λίρες και κόμματα στους πίνακες
    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
    }

    # ==========================================
    # 🟢 WORKSPACE: TAB 1 (CUB CHECK)
    # ==========================================
    if selected_tab == sheet_names[0]:
        st.markdown("### Ledger vs CUB Verification Workspace")
        
        cisa_data = {"prev_day": 2386124297.55, "debits": 10417421.49, "credits": 11826163.22, "total": 2387533039.28, "rec_date": 2387533039.28, "diff": 0.00}
        lisa_data = {"prev_day": 217714664.80, "debits": 251643.28, "credits": 946498.01, "total": 218409519.53, "rec_date": 218409519.53, "diff": 0.00}

        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check (CISA)</div><div class="badge-matched">MATCHED</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {cisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {cisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {cisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {cisa_data['total']:,.2f}</span></div>
                    <div class="recon-row" style="margin-top: 15px; border-top: 1px solid #1f2937;"><span>Internal CUB from Rec Date</span><span>£ {cisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #10b981;"><span>Variance Difference</span><span>£ {cisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

        with col_right:
            st.markdown(f"""
                <div class="workspace-card">
                    <div class="workspace-header"><div class="workspace-title">Combined User Balance Check (LISA)</div><div class="badge-matched">MATCHED</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {lisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {lisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {lisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {lisa_data['total']:,.2f}</span></div>
                    <div class="recon-row" style="margin-top: 15px; border-top: 1px solid #1f2937;"><span>Internal CUB from Rec Date</span><span>£ {lisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #10b981;"><span>Variance Difference</span><span>£ {lisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 🟢 WORKSPACE: TAB 2 (DAILY CM REPORT)
    # ==========================================
    elif selected_tab == sheet_names[1]:
        st.markdown("### Client Money Balances & Asset Ledger")
        
        # --- CASH ISA SECTION ---
        st.markdown("""
            <div class="workspace-card" style="padding-bottom: 5px;">
                <div class="workspace-header">
                    <div class="workspace-title" style="color: #8b5cf6;">Cash ISA Client Money Balances</div>
                    <div class="badge-variance">Net Change: -£971,704.00</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        cash_isa_data = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": 176992857.0, "COB Balance": 176021153.0, "Variance": -971704.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": 3959844.0, "COB Balance": 3959844.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": 747535672.0, "COB Balance": 747535672.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": 1168000000.0, "COB Balance": 1168000000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(cash_isa_data, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa")
        st.markdown("<br>", unsafe_allow_html=True)

        # --- LIFETIME ISA SECTION ---
        st.markdown("""
            <div class="workspace-card" style="padding-bottom: 5px;">
                <div class="workspace-header">
                    <div class="workspace-title" style="color: #8b5cf6;">Lifetime ISA Client Money Balances</div>
                    <div class="badge-matched">Net Change: +£610,408.97</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        lisa_data = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": 38363170.0, "COB Balance": 38973579.0, "Variance": 610408.97, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": 714980.0, "COB Balance": 714980.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": 79300000.0, "COB Balance": 79300000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": 100000000.0, "COB Balance": 100000000.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        st.data_editor(lisa_data, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa")

    # ==========================================
    # 🟢 ANY OTHER TAB (AUTOMATIC RAW DATA VIEW)
    # ==========================================
    else:
        st.markdown(f"### 📂 Archive Data View: {selected_tab}")
        df_any = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab, header=None)
        st.dataframe(df_any.dropna(how='all').reset_index(drop=True), use_container_width=True)

except Exception as e:
    st.error(f"System Error: {e}")
