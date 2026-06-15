import streamlit as st
import pandas as pd

st.set_page_config(page_title="CASS Bare Trust", layout="wide")

# Φόρτωση του αρχείου σου (όπως το είχαμε στήσει)
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec).xlsx"

try:
    excel_file = pd.ExcelFile(EXCEL_FILE)
    all_sheet_names = excel_file.sheet_names
except Exception as e:
    st.error(f"Σφάλμα φόρτωσης αρχείου: {e}")
    all_sheet_names = []

# --- SIDEBAR (Αριστερό Μενού) ---
st.sidebar.markdown("### CASS BARE TRUST\n**Client Money Reconciliation**")

# Ημερομηνία
recon_date = st.sidebar.date_input("Reconciliation date")

st.sidebar.markdown("---")
st.sidebar.caption("OVERVIEW")
# Χρησιμοποιούμε radio buttons για να μοιάζει με το μενού της φωτογραφίας
menu_options = ["Sign Off & Checks", "Daily CM Report", "Bare Trust Workings", "Account Details"]
selected_tab = st.sidebar.radio("Μενού:", menu_options, index=1) # Default στο Daily CM Report

st.sidebar.markdown("---")
st.sidebar.caption("EXPORT")
st.sidebar.button("📊 Export to Excel (.xlsx)", use_container_width=True)
st.sidebar.button("🖨️ Print / PDF", use_container_width=True)

# --- MAIN CONTENT (Κύριο Μέρος) ---
if selected_tab == "Daily CM Report":
    
    st.title("Daily Client Money Report")
    st.caption("Saveable – Bare Trust Client Money Balances (GBP)")
    
    # 1. TOP CARDS (Τα 4 γκρι κουτιά στην κορυφή)
    # Θα αρχικοποιήσουμε κάποιες μεταβλητές που θα υπολογίζονται παρακάτω
    total_requirement = 0.00
    total_resource = 0.00
    shortfall_surplus = total_resource - total_requirement
    net_change = 0.00
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.info(f"**Total Requirement**\n\n## £ {total_requirement:,.2f}")
    with col2:
        st.info(f"**Resource**\n\n## £ {total_resource:,.2f}")
    with col3:
        st.info(f"**Shortfall / Surplus**\n\n## £ {shortfall_surplus:,.2f}")
    with col4:
        st.info(f"**Net Change (COB)**\n\n## £ {net_change:,.2f}")
        
    st.markdown("---")
    
    # 2. ACCOUNT BALANCES TABLE (Ο πίνακας με τα editable κουτάκια)
    st.subheader("Account Balances")
    
    # Φτιάχνουμε τα default δεδομένα που είδα στην εικόνα σου
    default_accounts = pd.DataFrame([
        {"BANK": "Modulr", "ACCOUNT": "Net Settlement Account", "PREV DAY (£)": 0.00, "COB BALANCE (£)": 0.00, "VARIANCE (£)": 0.00, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds", "ACCOUNT": "Bare Trust Account (1)", "PREV DAY (£)": 0.00, "COB BALANCE (£)": 0.00, "VARIANCE (£)": 0.00, "ENTITY": "Saveable Limited"},
        {"BANK": "Lloyds", "ACCOUNT": "Bare Trust Account (2)", "PREV DAY (£)": 0.00, "COB BALANCE (£)": 0.00, "VARIANCE (£)": 0.00, "ENTITY": "Saveable Limited"},
    ])
    
    # Το st.data_editor επιτρέπει στον χρήστη να γράφει μέσα στα κουτάκια live!
    edited_df = st.data_editor(
        default_accounts, 
        num_rows="dynamic", # Επιτρέπει το "+ Add account" και το "Remove"
        use_container_width=True,
        key="accounts_table"
    )
    
    # Υπολογισμός των Total για τον πίνακα
    total_prev = edited_df["PREV DAY (£)"].sum()
    total_cob = edited_df["COB BALANCE (£)"].sum()
    total_variance = total_cob - total_prev
    
    # Εμφάνιση της γραμμής TOTAL όπως στην εικόνα
    st.markdown(
        f"| TOTAL | | **£ {total_prev:,.2f}** | **£ {total_cob:,.2f}** | **£ {total_variance:,.2f}** | |", 
        help="Αυτόματο άθροισμα των παραπάνω γραμμών"
    )
    
    st.markdown("---")
    
    # 3. DAILY RECONCILIATION FORM (Τα inputs στο κάτω μέρος)
    st.subheader("Daily Reconciliation")
    
    recon_col1, recon_col2, recon_col3 = st.columns(3)
    with recon_col1:
        requirement_input = st.number_input("Requirement (£)", min_value=0.00, value=0.00, step=100.00)
    with recon_col2:
        transfers_input = st.number_input("Transfers In to Apply (£)", min_value=0.00, value=0.00, step=100.00)
    with recon_col3:
        # Το Resource υπολογίζεται αυτόματα από το Total COB του πάνω πίνακα
        st.number_input("Resource (£)", value=float(total_cob), disabled=True)
        
    # Calculated Shortfall / Surplus μεγάλο κουτί
    calculated_diff = total_cob - requirement_input
    st.text_input("Calculated Shortfall / Surplus", value=f"£ {calculated_diff:,.2f}", disabled=True)
    
    # 4. TEXT AREAS (Λόγοι και Σχόλια)
    st.markdown("---")
    st.text_area("Reason for Internal Movements", placeholder="Describe internal transfers...")
    st.text_area("Additional Comments", placeholder="Additional notes...")

else:
    # Για τις υπόλοιπες καρτέλες, δείχνουμε απλά τα δεδομένα του Excel
    st.subheader(f"Καρτέλα: {selected_tab}")
    try:
        df_excel = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab)
        st.dataframe(df_excel, use_container_width=True)
    except:
        st.info("Δεν βρέθηκαν δεδομένα στο Excel για αυτή την επιλογή.")
