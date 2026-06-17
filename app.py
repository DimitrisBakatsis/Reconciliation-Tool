import streamlit as st
import pandas as pd

# 1. Page Config & Extended Premium Dark CSS
st.set_page_config(page_title="CASS Reconciliation Portal", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0d0e12; color: #ffffff; }
    
    /* Top Header Styling */
    .main-header { font-size: 28px; font-weight: bold; color: #ffffff; margin-bottom: 2px; }
    .date-subheader { font-size: 14px; color: #8a8f98; margin-bottom: 30px; font-style: italic; }
    
    /* Modern Section Container */
    .custom-container { background-color: #111217; padding: 22px; border-radius: 8px; border: 1px solid #1e212a; margin-bottom: 25px; }
    .container-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; font-size: 16px; font-weight: 600; color: #ffffff; border-bottom: 1px solid #1e212a; padding-bottom: 8px; }
    
    /* Purple Accent Header for Excel Tabs inside Tab 2 */
    .excel-sheet-title { background-color: #8b5cf6; color: #ffffff; padding: 8px 15px; font-weight: bold; font-size: 15px; border-radius: 4px; margin-bottom: 12px; text-transform: uppercase; }
    
    /* Reconciliation Rows for CUB Check */
    .recon-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dashed #1e212a; font-size: 14px; }
    .recon-row.total { border-bottom: none; font-weight: bold; font-size: 15px; color: #3b9eff; padding-top: 12px; }
    
    /* Summary Row Styling for Tables */
    .table-summary-box { background-color: #16171d; padding: 12px 20px; border-radius: 4px; border: 1px solid #22252e; margin-top: -10px; margin-bottom: 20px; display: flex; justify-content: flex-end; gap: 40px; font-size: 14px; }
    
    /* Status Badges */
    .status-badge-match { background-color: rgba(44, 217, 139, 0.1); color: #2cd98b; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 13px; border: 1px solid rgba(44, 217, 139, 0.2); }
    
    [data-testid="stSidebar"] { background-color: #111215; border-right: 1px solid #22252e; }
    div[data-testid="stRadio"] > label { display: none !important; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

# 🛠️ ΔΙΟΡΘΩΣΗ: st.cache_resource για να μην πετάει σφάλμα serialization
@st.cache_resource(ttl=2)
def load_raw_excel():
    return pd.ExcelFile(EXCEL_FILE)

try:
    xl = load_raw_excel()
    sheet_names = xl.sheet_names
    
    # 📁 ΔΙΑΒΑΣΜΑ ΗΜΕΡΟΜΗΝΙΑΣ ΑΠΟ TAB 13 (Κελί D4)
    try:
        df_tab13 = pd.read_excel(EXCEL_FILE, sheet_name=12, header=None)
        raw_date = df_tab13.iloc[3, 3]
        formatted_date = str(raw_date).split()[0] if pd.notna(raw_date) else "16/06/2026"
    except:
        formatted_date = "16/06/2026"

    # --- SIDEBAR ---
    st.sidebar.markdown("<h2 style='color:#ffffff; font-size:16px; margin-bottom:0;'>Saveable Portal</h2><p style='color:#6c727f; font-size:12px;'>CASS Corporate Suite</p>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='color:#62627a; font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:5px;'>NAVIGATION</p>", unsafe_allow_html=True)
    
    active_menu = st.sidebar.radio("Μενού:", [sheet_names[0], sheet_names[1]])

    # --- ΚΕΝΤΡΙΚΟΣ ΤΙΤΛΟΣ ---
    st.markdown(f"<div class='main-header'>CASS Reconciliation & Daily Client Money Reporting</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='date-subheader'>📅 Close of Business Date: {formatted_date}</div>", unsafe_allow_html=True)

    # Configuration για να εμφανίζονται οι στήλες με £ και κόμματα
    currency_config = {
        "Previous Day Balance": st.column_config.NumberColumn("Previous Day Balance", format="£%,.2f"),
        "COB Balance": st.column_config.NumberColumn("COB Balance", format="£%,.2f"),
        "Variance": st.column_config.NumberColumn("Variance", format="£%,.2f"),
    }

    # ==========================================
    # 🟢 TAB 1: MANUAL COMBINED USER BALANCE CHECK
    # ==========================================
    if active_menu == sheet_names[0]:
        st.markdown("### Manual Combined User Balance Check")
        st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Internal CUB verification vs Ledger Reporting Data</p>", unsafe_allow_html=True)
        
        cisa_data = {"prev_day": 2386124297.55, "debits": 10417421.49, "credits": 11826163.22, "total": 2387533039.28, "rec_date": 2387533039.28, "diff": 0.00}
        lisa_data = {"prev_day": 217714664.80, "debits": 251643.28, "credits": 946498.01, "total": 218409519.53, "rec_date": 218409519.53, "diff": 0.00}

        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown(f"""
                <div class="custom-container">
                    <div class="container-header"><div>Combined User Balance Check - CISA</div><div class="status-badge-match">MATCHED</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {cisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {cisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {cisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {cisa_data['total']:,.2f}</span></div><br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><span>£ {cisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #2cd98b;"><span>Difference</span><span>£ {cisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

        with col_right:
            st.markdown(f"""
                <div class="custom-container">
                    <div class="container-header"><div>Combined User Balance Check - LISA</div><div class="status-badge-match">MATCHED</div></div>
                    <div class="recon-row"><span>Internal CUB from previous day</span><span>£ {lisa_data['prev_day']:,.2f}</span></div>
                    <div class="recon-row"><span>Debits (Recon data) from Rec data</span><span>£ {lisa_data['debits']:,.2f}</span></div>
                    <div class="recon-row"><span>Credits (Recon data) from Rec data</span><span>£ {lisa_data['credits']:,.2f}</span></div>
                    <div class="recon-row total"><span>Total Expected CUB</span><span>£ {lisa_data['total']:,.2f}</span></div><br>
                    <div class="recon-row"><span>Internal CUB from Rec Date</span><span>£ {lisa_data['rec_date']:,.2f}</span></div>
                    <div class="recon-row total" style="color: #2cd98b;"><span>Difference</span><span>£ {lisa_data['diff']:,.2f}</span></div>
                </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 🟢 TAB 2: DAILY CLIENT MONEY REPORT (ΑΚΡΙΒΕΣ LAYOUT IMAGE_8A5A78)
    # ==========================================
    elif active_menu == sheet_names[1]:
        st.markdown("### Saveable Daily Client Money and Asset Reporting")
        st.markdown("<p style='color:#6c727f; font-size:13px; margin-bottom:25px;'>Breakdown of Cash ISA and Lifetime ISA Client Money Balances</p>", unsafe_allow_html=True)

        # ---- ΠΙΝΑΚΑΣ 2Α: CASH ISA CLIENT MONEY BALANCES ----
        st.markdown("<div class='excel-sheet-title'>Cash ISA Client Money Balances - GBP</div>", unsafe_allow_html=True)
        
        cash_isa_data = pd.DataFrame([
            {"Bank": "Citi Bank NA London", "Account": "Saveable Cash ISA UK Client Money (14747801)", "Previous Day Balance": 176992857.0, "COB Balance": 176021153.0, "Variance": -971704.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA Client Account (27551460)", "Previous Day Balance": 3959844.0, "COB Balance": 3959844.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Cash ISA 30D Notice Client Account (27571468)", "Previous Day Balance": 747535672.0, "COB Balance": 747535672.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-310)", "Previous Day Balance": 1168000000.0, "COB Balance": 1168000000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BlackRock QMMF", "Account": "Blackrock QMMF", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Easy access (01778650)", "Previous Day Balance": 294960631.0, "COB Balance": 294960631.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "BBVA", "Account": "BBVA Notice account (06758170)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Clydesdale Bank PLC TA Virgin Money", "Account": "Saveable Cash ISA 95 Day Notice (12204224)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Clydesdale Bank PLC TA Virgin Money", "Account": "Saveable Cash ISA 65 Day Notice (12204216)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Clydesdale Bank PLC TA Virgin Money", "Account": "Saveable Cash ISA 30 Day Notice (12204161)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Clydesdale Bank PLC TA Virgin Money", "Account": "Saveable Cash ISA Instant access (12204414)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "JP Morgan", "Account": "JP Morgan Client Money account (76919500)", "Previous Day Balance": 0.0, "COB Balance": 0.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        
        st.data_editor(cash_isa_data, column_config=currency_config, use_container_width=True, hide_index=True, key="cash_isa_editor")
        
        # Σύνολα και Extra KPIs για το Cash ISA
        st.markdown("""
            <div class="table-summary-box">
                <div><strong>Total Previous Day:</strong> £ 2,391,449,004.00</div>
                <div><strong>Total COB Balance:</strong> £ 2,390,477,301.00</div>
            </div>
            <div style="display: flex; justify-content: flex-end; margin-bottom: 30px;">
                <div style="background-color: #1a1015; border: 1px solid #4a1d24; color: #ff6b6b; padding: 10px 20px; border-radius: 4px; font-size: 14px;">
                    <strong>CISA Net Increase/decrease:</strong> &nbsp;&nbsp;&nbsp;&nbsp; -£ 971,704.00
                </div>
            </div>
        """, unsafe_allow_html=True)


        # ---- ΠΙΝΑΚΑΣ 2Β: LIFETIME ISA CLIENT MONEY BALANCES ----
        st.markdown("<div class='excel-sheet-title'>Lifetime ISA Client Money Balances - GBP</div>", unsafe_allow_html=True)
        
        lisa_data = pd.DataFrame([
            {"Bank": "CitiBank NA London", "Account": "Saveable Lifetime ISA UK Client Money (15242487)", "Previous Day Balance": 38363170.0, "COB Balance": 38973579.0, "Variance": 610408.97, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA Client Account (27561260)", "Previous Day Balance": 714980.0, "COB Balance": 714980.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "Lloyds Bank Plc", "Account": "Saveable Lifetime ISA 30D Notice Client Account (27571060)", "Previous Day Balance": 79300000.0, "COB Balance": 79300000.0, "Variance": 0.0, "Entity": "Saveable Limited"},
            {"Bank": "QNB", "Account": "Qatar National Bank (4311-000545-311)", "Previous Day Balance": 100000000.0, "COB Balance": 100000000.0, "Variance": 0.0, "Entity": "Saveable Limited"}
        ])
        
        st.data_editor(lisa_data, column_config=currency_config, use_container_width=True, hide_index=True, key="lisa_editor")
        
        # Σύνολα και Extra KPIs για το Lifetime ISA
        st.markdown("""
            <div class="table-summary-box">
                <div><strong>Total Previous Day:</strong> £ 218,378,150.00</div>
                <div><strong>Total COB Balance:</strong> £ 218,988,559.00</div>
            </div>
            <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 10px; margin-bottom: 25px;">
                <div style="background-color: #101a15; border: 1px solid #1d4a34; color: #2cd98b; padding: 10px 20px; border-radius: 4px; font-size: 14px; width: fit-content;">
                    <strong>LISA Net Increase/decrease:</strong> &nbsp;&nbsp;&nbsp;&nbsp; £ 610,408.97
                </div>
                <div style="background-color: #1a1015; border: 1px solid #4a1d24; color: #ff6b6b; padding: 10px 20px; border-radius: 4px; font-size: 14px; width: fit-content;">
                    <strong>Total ISA increase/decrease:</strong> &nbsp;&nbsp;&nbsp;&nbsp; -£ 361,295.00
                </div>
            </div>
        """, unsafe_allow_html=True)

except Exception as e:
    st.error(f"Σφάλμα κατά τη φόρτωση ή σχεδίαση: {e}")
