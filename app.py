import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Βασικές ρυθμίσεις της σελίδας
st.set_page_config(page_title="Δυναμικό Dashboard", layout="wide")
st.title("📊 Google Sheets Advanced Dashboard")

# 2. Το σταθερό ID του Google Sheet σου (αντικατάστησέ το με το δικό σου)
SHEET_ID = "1G33LdbKgm7L-lA5K3FTZcPWqxVuAeU1Ga1_ZSADIScM" 

# 3. Εδώ δηλώνεις ΚΑΙ ΤΙΣ 15 ΚΑΡΤΕΛΕΣ σου (Όνομα: GID)
# Άνοιξε το GSheet, πάτα κάθε καρτέλα και αντέγραψε το νούμερο μετά το #gid=
ALL_TABS = {
    "1. Sign Off & Other Checks": "1441606621",
    "2. Daily Client Money Report": "1462165387",
    "3. Unalloc Rec": "87135728",
    "5. CISA Internal Workings": "1330436512",
    # ... Πρόσθεσε εδώ και τις υπόλοιπες καρτέλες σου με τον ίδιο τρόπο
}

# 4. Sidebar για τις βασικές παραμέτρους
st.sidebar.header("🎛️ Παράμετροι Εφαρμογής")

# ΠΑΡΑΜΕΤΡΟΣ 1: Επιλογή Καρτέλας
selected_tab_name = st.sidebar.selectbox("1. Διάλεξε Καρτέλα για διάβασμα:", list(ALL_TABS.keys()))
selected_gid = ALL_TABS[selected_tab_name]

# Κατασκευή του δυναμικού URL για τη συγκεκριμένη καρτέλα
csv_url = f"https://docs.google.com/spreadsheets/d/1G33LdbKgm7L-lA5K3FTZcPWqxVuAeU1Ga1_ZSADIScM/export?format=csv&gid=1462165387"

# 5. Συνάρτηση για live "διάβασμα" της επιλεγμένης καρτέλας
@st.cache_data(ttl=30)  # Cache για 30 δευτερόλεπτα ώστε να είναι γρήγορο
def load_sheet_data(url):
    return pd.read_csv(url)

try:
    # Φόρτωση δεδομένων της καρτέλας
    df = load_sheet_data(csv_url)
    
    st.subheader(f"📂 Δεδομένα από: {selected_tab_name}")
    
    # ΠΑΡΑΜΕΤΡΟΣ 2: Δυναμικό Φιλτράρισμα Γραμμών (Προαιρετικό)
    # Αν υπάρχει στήλη που θέλεις να φιλτράρεις, π.χ. "Κατηγορία" ή "Έτος"
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Φίλτρα Δεδομένων")
    
    # Επιλογή στήλης για φιλτράρισμα
    filter_col = st.sidebar.selectbox("Διάλεξε στήλη για φιλτράρισμα (προαιρετικά):", ["Καμία"] + df.columns.tolist())
    
    if filter_col != "Καμία":
        unique_values = df[filter_col].unique().tolist()
        selected_values = st.sidebar.multiselect(f"Διάλεξε τιμές για {filter_col}:", unique_values, default=unique_values)
        # Εφαρμογή του φίλτρου στα δεδομένα
        df = df[df[filter_col].isin(selected_values)]

    # Εμφάνιση του πίνακα με τα φιλτραρισμένα δεδομένα
    st.dataframe(df, use_container_width=True)

    # ΠΑΡΑΜΕΤΡΟΣ 3: Ρυθμίσεις Διαγράμματος
    st.sidebar.markdown("---")
    st.sidebar.subheader("📈 Παράμετροι Διαγράμματος")
    
    columns_list = df.columns.tolist()
    
    if len(columns_list) >= 2:
        x_axis = st.sidebar.selectbox("Άξονας Χ (Κατηγορίες/Ημερομηνίες):", columns_list, index=0)
        y_axis = st.sidebar.selectbox("Άξονας Υ (Νούμερα/Ποσότητες):", columns_list, index=min(1, len(columns_list)-1))
        chart_type = st.sidebar.selectbox("Τύπος Διαγράμματος:", ["Ράβδοι (Bar)", "Γραμμή (Line)", "Διασπορά (Scatter)", "Πίτα (Pie)"])
        
        st.markdown("---")
        st.subheader("📊 Οπτικοποίηση")
        
        # Σχεδιασμός ανάλογα με τον τύπο που επέλεξε ο χρήστης
        if chart_type == "Ράβδοι (Bar)":
            fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_dark")
        elif chart_type == "Γραμμή (Line)":
            fig = px.line(df, x=x_axis, y=y_axis, template="plotly_dark")
        elif chart_type == "Διασπορά (Scatter)":
            fig = px.scatter(df, x=x_axis, y=y_axis, template="plotly_dark")
        elif chart_type == "Πίτα (Pie)":
            fig = px.pie(df, names=x_axis, values=y_axis, template="plotly_dark")
            
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Η καρτέλα πρέπει να έχει τουλάχιστον 2 στήλες για να γίνει διάγραμμα.")

except Exception as e:
    st.error(f"⚠️ Σφάλμα κατά τη φόρτωση της καρτέλας: {e}")
    st.info("Σιγουρέψου ότι το αρχείο είναι προσβάσιμο (Anyone with the link -> Viewer) και ότι έχεις βάλει τα σωστά GID.")
