import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Plum Dashboard", layout="wide")
st.title("📊 Τοπικό Excel Dashboard")

# 1. Όνομα του αρχείου που ανέβασες στο GitHub
EXCEL_FILE = "AUTOMATION CASS Reconciliation & Daily Client Money Reporting Template - (Daily Cash Rec)xlsx" 

try:
    # Το pandas διαβάζει αυτόματα όλα τα ονόματα των καρτελών από το Excel!
    excel_reader = pd.ExcelFile(EXCEL_FILE)
    all_sheets = excel_reader.sheet_names

    # Sidebar επιλογή καρτέλας
    st.sidebar.header("⚙️ Επιλογές")
    selected_tab = st.sidebar.selectbox("Διάλεξε Καρτέλα:", all_sheets)

    # Διάβασμα της συγκεκριμένης καρτέλας
    df = pd.read_excel(EXCEL_FILE, sheet_name=selected_tab)
    
    st.subheader(f"📂 Δεδομένα από: {selected_tab}")
    st.dataframe(df, use_container_width=True)
    
    # Ρυθμίσεις Διαγράμματος
    cols = df.columns.tolist()
    if len(cols) >= 2:
        x_axis = st.sidebar.selectbox("Άξονας Χ:", cols)
        y_axis = st.sidebar.selectbox("Άξονας Υ:", cols)
        
        fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_dark")
        st.plotly_chart(fig, use_container_width=True)

except Exception as e:
    st.error(f"Σφάλμα κατά το διάβασμα του αρχείου: {e}")
