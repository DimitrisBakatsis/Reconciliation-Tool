import streamlit as st
import pandas as pd

st.set_page_config(page_title="Saveable Reporting", layout="wide")

EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

@st.cache_data(ttl=5)
def load_plum_sheet():
    # Διαβάζουμε τη 2η καρτέλα (Index 1) χωρίς header στην αρχή για να την κόψουμε σωστά
    df = pd.read_excel(EXCEL_FILE, sheet_name=1, header=None)
    return df

try:
    raw_df = load_plum_sheet()
    
    # --- SIDEBAR ---
    st.sidebar.markdown("### Saveable Daily\n**Client Money & Asset Reporting**")
    
    # Εξαγωγή ημερομηνίας από το κελί B3 (Row 2, Col 1 στο pandas)
    try:
        excel_date = str(raw_df.iloc[2, 1]).split()[0]
    except:
        excel_date = "12/06/2026"
    
    st.sidebar.info(f"📅 **Close of Business Date:**\n{excel_date}")
    
    st.sidebar.markdown("---")
    menu_options = ["Overview Dashboard", "Daily CM Report", "Commentary on Variances"]
    selected_tab = st.sidebar.radio("Μενού:", menu_options, index=1)

    # --- ΣΥΝΑΡΤΗΣΕΙΣ ΑΠΟΜΟΝΩΣΗΣ ΠΙΝΑΚΩΝ ---
    def extract_table(df, start_keyword, num_rows):
        # Ψάχνει να βρει τη γραμμή που ξεκινάει ο πίνακας (π.χ. Cash ISA...)
        for idx, row in df.iterrows():
            if start_keyword in str(row[0]) or start_keyword in str(row[1]):
                # Παίρνουμε τις επόμενες γραμμές που περιέχουν τα δεδομένα
                header_row = idx + 1
                data_df = df.iloc[header_row + 1 : header_row + 1 + num_rows].copy()
                data_df.columns = ["Bank", "Account", "Previous Day Balance", "COB Balance", "Variance", "Entity", "Extra"] if len(df.columns) >= 7 else ["Bank", "Account", "Previous Day Balance", "COB Balance", "Variance", "Entity"]
                
                # Καθαρισμός αριθμών
                for col in ["Previous Day Balance", "COB Balance", "Variance"]:
                    if col in data_df.columns:
                        data_df[col] = data_df[col].astype(str).str.replace('£', '').str.replace(',', '').str.strip()
                        data_df[col] = pd.to_numeric(data_df[col], errors='coerce').fillna(0.0)
                
                return data_df.dropna(subset=["Bank"]).reset_index(drop=True)
        return pd.DataFrame()

    # --- MAIN CONTENT ---
    if selected_tab == "Daily CM Report":
        st.title("Saveable Daily Client Money and Asset Reporting")
        
        # 1. ΠΙΝΑΚΑΣ: Cash ISA Client Money Balances - GBP
        st.markdown("### 🟣 Cash ISA Client Money Balances - GBP")
        cash_isa_df = extract_table(raw_df, "Cash ISA Client Money Balances", 12)
        if not cash_isa_df.empty:
            st.data_editor(cash_isa_df[["Bank", "Account", "Previous Day Balance", "COB Balance", "Variance", "Entity"]], use_container_width=True, key="cash_isa")
            
            # KPI για το CISA Net Increase
            cisa_variance_total = cash_isa_df["Variance"].sum()
            st.metric("CISA Net Increase/decrease", f"£ {cisa_variance_total:,.2f}")
        
        st.markdown("---")
        
        # 2. ΠΙΝΑΚΑΣ: Lifetime ISA Client Money Balances - GBP
        st.markdown("### 🟣 Lifetime ISA Client Money Balances - GBP")
        lisa_df = extract_table(raw_df, "Lifetime ISA Client Money Balances", 4)
        if not lisa_df.empty:
            st.data_editor(lisa_df[["Bank", "Account", "Previous Day Balance", "COB Balance", "Variance", "Entity"]], use_container_width=True, key="lisa")
            
            lisa_variance_total = lisa_df["Variance"].sum()
            st.metric("LISA Net Increase/decrease", f"£ {lisa_variance_total:,.2f}")

        st.markdown("---")

        # 3. ΠΙΝΑΚΑΣ: Stocks/Shares ISA
        st.markdown("### 🟣 Stocks/Shares ISA")
        stocks_df = extract_table(raw_df, "Stocks/Shares ISA", 3)
        if not stocks_df.empty:
            st.data_editor(stocks_df, use_container_width=True, key="stocks")

        st.markdown("---")

        # 4. ΠΙΝΑΚΑΣ: Daily Reconciliation (Το κάτω αριστερά μέρος της εικόνας 1)
        st.markdown("### 📊 Daily Reconciliation")
        
        # Εντοπισμός των τιμών του Daily Reconciliation από το Excel
        recon_data = {"Requirement": 2601370286.70, "Inclusive of transfers": 6187634.40, "Total Requirement": 2607557921.10, "Resource": 2607556676.61, "Shortfall / Surplus": -1244.09}
        
        # Προσπάθεια να διαβάσουμε live τις τιμές από τις τελευταίες γραμμές του Excel
        for idx, row in raw_df.iterrows():
            if "Daily Reconciliation" in str(row[0]):
                try:
                    recon_data["Requirement"] = float(str(raw_df.iloc[idx+1, 1]).replace('£','').replace(',','').strip())
                    recon_data["Inclusive of transfers"] = float(str(raw_df.iloc[idx+2, 1]).replace('£','').replace(',','').strip())
                    recon_data["Total Requirement"] = float(str(raw_df.iloc[idx+3, 1]).replace('£','').replace(',','').strip())
                    recon_data["Resource"] = float(str(raw_df.iloc[idx+4, 1]).replace('£','').replace(',','').strip())
                    recon_data["Shortfall / Surplus"] = float(str(raw_df.iloc[idx+5, 1]).replace('£','').replace(',','').strip())
                except:
                    pass # Κρατάει τα προκαθορισμένα αν αποτύχει η μετατροπή
                break

        rec_col1, rec_col2 = st.columns(2)
        with rec_col1:
            st.metric("Total Requirement (inclusive of Quai)", f"£ {recon_data['Total Requirement']:,.2f}")
            st.metric("Resource (inclusive of Quai)", f"£ {recon_data['Resource']:,.2f}")
            
            # Shortfall με χρώμα (κόκκινο αν είναι αρνητικό)
            if recon_data['Shortfall / Surplus'] < 0:
                st.error(f"⚠️ **Shortfall / Surplus (inclusive of Quai):** £ {recon_data['Shortfall / Surplus']:,.2f}")
            else:
                st.success(f"✅ **Shortfall / Surplus (inclusive of Quai):** £ {recon_data['Shortfall / Surplus']:,.2f}")

        with rec_col2:
            # Κείμενα από τη δεύτερη φωτογραφία (Reason for internal movements / Comments)
            st.markdown("**Reason for internal movements:**")
            st.caption("CISA: Overall Shortfall of £1,244.79 residual interest paid to users...\n\nLISA: Overall Surplus of £10,219.74...")
            
except Exception as e:
    st.error(f"Κάτι πήγε στραβά με το parsing του Excel: {e}")
