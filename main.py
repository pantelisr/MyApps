import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime

from datetime import timedelta

import pandas as pd

import os

st.set_page_config(layout="wide")

NUM_TO_GREEK = {
    1: "μίας ημέρας (1)", 2: "δύο ημερών (2)", 3: "τριών ημερών (3)", 4: "τεσσάρων ημερών (4)", 5: "πέντε ημερών (5)",
    6: "έξι ημερών (6)", 7: "επτά ημερών (7)", 8: "οκτώ ημερών (8)", 9: "εννέα ημερών (9)", 10: "δέκα ημερών (10)",
}

def save_to_log(data):
    file_path = "log_adeies.xlsx"
    # Μετατρέπουμε τα δεδομένα σε πίνακα (DataFrame)
    new_entry = pd.DataFrame([data])
    
    if os.path.exists(file_path):
        # Αν το αρχείο υπάρχει, το ανοίγουμε, προσθέτουμε τη γραμμή και το ξανασώζουμε
        df = pd.read_excel(file_path)
        df = pd.concat([df, new_entry], ignore_index=True)
    else:
        # Αν δεν υπάρχει, δημιουργούμε ένα νέο
        df = new_entry
        
    df.to_excel(file_path, index=False)

# Συνάρτηση για τη φόρτωση των δεδομένων
def load_teachers():
    try:
        # Διαβάζει το Excel
        df = pd.read_excel("teachers.xlsx",dtype={'mitrwo': str})

        # Δεν υπάρχουν NaN που γίνονται "nan"
        df['mitrwo'] = df['mitrwo'].fillna('')
        
        # Μετατρέπει το DataFrame σε ένα dictionary που μπορεί να χρησιμοποιήσει το Streamlit
        # Χρησιμοποιούμε το 'full_name' ως κλειδί
        data = df.set_index('full_name').to_dict('index')
        return data
    except Exception as e:
        st.error(f"Σφάλμα κατά τη φόρτωση του αρχείου Excel: {e}")
        return {"Επιλέξτε...": {"eponymo": "", "onoma": "", "klados": "", "mitrwo": "", "email": ""}}

EKPAIDEYTIKOI_DATA = load_teachers()

st.title("📝 Έκδοση Άδειας Εκπαιδευτικού")

# --- 1. Δημιουργία της λίστας για το selectbox (μια κενή επιλογή στην αρχή)
teacher_options = ["-- Επιλέξτε Εκπαιδευτικό --"] + list(EKPAIDEYTIKOI_DATA.keys())

selected_name = st.selectbox("1. Αναζήτηση Εκπαιδευτικού", options=teacher_options)

# --- 2. ΕΠΙΛΟΓΗ ΠΡΟΤΥΠΟΥ ΑΠΟ ΤΟΝ ΦΑΚΕΛΟ /templates ---
TEMPLATE_FOLDER = "templates"

# Δημιουργία του φακέλου αν δεν υπάρχει (για αποφυγή σφαλμάτων)
if not os.path.exists(TEMPLATE_FOLDER):
    os.makedirs(TEMPLATE_FOLDER)

# Λίστα με τα διαθέσιμα πρότυπα μέσα στον φάκελο
available_templates = ["-- Επιλέξτε το Έγγραφο της Άδειας --"] + [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith('.docx')]


if available_templates:
    selected_template_name = st.selectbox("2. Επιλέξτε το έντυπο προς έκδοση:", options=available_templates)
    template_path = os.path.join(TEMPLATE_FOLDER, selected_template_name)
else:
    st.warning("⚠️ Δεν βρέθηκαν αρχεία .docx στον φάκελο /templates. Παρακαλώ προσθέστε τα πρότυπά σας εκεί.")
    template_path = None

st.divider() # Μια οριζόντια γραμμή για να ξεχωρίζει η προετοιμασία από τη φόρμα


# Ανάκτηση στοιχείων
if selected_name != "-- Επιλέξτε Εκπαιδευτικό --":
    teacher_data = EKPAIDEYTIKOI_DATA[selected_name]
else:
    teacher_data = {"eponymo": "", "onoma": "", "klados": "","mitrwo": "", "email": ""}

# Ρύθμιση σελίδας
st.set_page_config(page_title="Αυτοματοποίηση Αδειών Σχολείου", page_icon="📝")

st.subheader("Συμπληρώστε τα στοιχεία για την παραγωγή του εγγράφου")

# Φόρμα Εισαγωγής Στοιχείων
st.subheader("Στοιχεία Άδειας")
col1, col2, col3, col4 = st.columns(4)
    
with col1:
        eponymo = st.text_input("Επώνυμο", value=teacher_data["eponymo"], disabled=True)
        
with col2:     
        onoma = st.text_input("Όνομα ", value=teacher_data["onoma"], disabled=True)
        hmer_protocollou_aithshs = st.date_input("Ημερομηνία Πρωτοκόλλου Αίτησης",format="DD/MM/YYYY")
        days_number = st.number_input("Αριθμός Ημερών", min_value=1, step=1)
        hmer_gnomateyshs = st.date_input("Ημερομηνία Γνωμάτευσης",format="DD/MM/YYYY")
        
with col3:
        mitrwo = st.text_input("ΑΜ ", value=teacher_data["mitrwo"],disabled=True)
        protocollo_aithshs = st.text_input("Αριθμός Πρωτοκόλλου Αίτησης")
        arxh = st.date_input("Ημερομηνία Έναρξης",format="DD/MM/YYYY")
        doctor = st.text_input("Γιατρός")
        
with col4:
        klados = st.text_input("Κλάδος ", value=teacher_data["klados"], disabled=True)
        telos = arxh + timedelta(days=days_number - 1)
        st.markdown("<div style='height: 59px; margin-bottom: 24px;'></div>", unsafe_allow_html=True)        
        days_lektiko = NUM_TO_GREEK.get(days_number, str(days_number))
        st.text_input("Ημερομηνία Λήξης", value=telos.strftime('%d/%m/%Y'), disabled=True)
        
st.divider()
with st.form("leave_form"):     
    st.subheader("Στοιχεία Πρωτοκόλλου Άδειας")
    protocollo_adeias = st.text_input("Πρωτόκολλο Άδειας")
    hmer_protocollou = st.date_input("Ημερομηνία Πρωτοκόλλου Άδειας",format="DD/MM/YYYY")
    
    submitted = st.form_submit_button("Δημιουργία Εγγράφου")

if submitted:
    if not template_path:
        st.error("⚠️ Δεν έχει επιλεγεί πρότυπο αρχείο!")
    else:
        try:
            # Φόρτωση από τη διαδρομή του φακέλου templates
            doc = DocxTemplate(template_path)
            
            # Δεδομένα για το "γέμισμα" του εγγράφου
            context = {
                'eponymo': eponymo,
                'onoma': onoma,
                'hmer_protocollou': hmer_protocollou.strftime("%d/%m/%Y"),
                'protocollo_adeias' : protocollo_adeias,
                'hmer_protocollou_aithshs': hmer_protocollou_aithshs.strftime("%d-%m-%Y"),
                'protocollo_aithshs' : protocollo_aithshs,
                'klados' : klados,
                'mitrwo' : mitrwo,
                'days_number': days_number,
                'days_lektiko' : days_lektiko,
                'doctor' : doctor,
                'hmer_gnomateyshs' : hmer_gnomateyshs.strftime("%d/%m/%Y"),
                'arxh': arxh.strftime("%d/%m/%Y"),
                'telos': telos.strftime("%d/%m/%Y"),
                'hmerominia_ekdosis': datetime.now().strftime("%d/%m/%Y")
            }
            
            # Render το έγγραφο
            doc.render(context)
            
            # Δεδομένα για το ιστορικό (log)
            log_data = {
            'Ημερομηνία Έκδοσης': datetime.now().strftime("%d/%m/%Y %H:%M"),
            'Επώνυμο': eponymo,
            'Όνομα': onoma,
            'Κλάδος': klados,
            'Ημέρες': days_number,
            'Πρωτόκολλο': protocollo_adeias,
            'Ημερομηνία πρωτοκόλλου άδειας' : hmer_protocollou
              }
        
            # Αποθήκευση στο Excel
            try:
                save_to_log(log_data)
                st.info("Η εγγραφή καταγράφηκε επιτυχώς στο ιστορικό (log_adeies.xlsx).")
            except Exception as e:
                st.error(f"Σφάλμα κατά την καταγραφή στο ιστορικό: {e}")


            # Αποθήκευση σε "μνήμη" (BytesIO) για να το κατεβάσει ο χρήστης
            target_stream = io.BytesIO()
            doc.save(target_stream)
            target_stream.seek(0)
            
            st.success(f"Η άδεια για τον/την {eponymo} {onoma} είναι έτοιμη!")
            
            # Κουμπί Download
            st.download_button(
                label="📥 Λήψη Εγγράφου (Word)",
                data=target_stream,
                file_name=f"Adeia_{eponymo}_{protocollo_adeias.replace('/', '-')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Παρουσιάστηκε σφάλμα: {e}")