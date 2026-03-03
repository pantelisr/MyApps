import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime

import pandas as pd

import os

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
        df = pd.read_excel("teachers.xlsx")
        
        # Μετατρέπει το DataFrame σε ένα dictionary που μπορεί να χρησιμοποιήσει το Streamlit
        # Χρησιμοποιούμε το 'full_name' ως κλειδί
        data = df.set_index('full_name').to_dict('index')
        return data
    except Exception as e:
        st.error(f"Σφάλμα κατά τη φόρτωση του αρχείου Excel: {e}")
        return {"Επιλέξτε...": {"eponymo": "", "onoma": "", "klados": "", "mitrwo": "", "email": ""}}

EKPAIDEYTIKOI_DATA = load_teachers()

# Δημιουργία της λίστας για το selectbox (προσθέτουμε μια κενή επιλογή στην αρχή)
teacher_options = ["Επιλέξτε Εκπαιδευτικό"] + list(EKPAIDEYTIKOI_DATA.keys())

selected_name = st.selectbox("Αναζήτηση Εκπαιδευτικού", options=teacher_options)

# Ανάκτηση στοιχείων
if selected_name != "Επιλέξτε Εκπαιδευτικό":
    teacher_data = EKPAIDEYTIKOI_DATA[selected_name]
else:
    teacher_data = {"eponymo": "", "onoma": "", "klados": "","mitrwo": "", "email": ""}

# Ρύθμιση σελίδας
st.set_page_config(page_title="Αυτοματοποίηση Αδειών Σχολείου", page_icon="📝")

st.title("📝 Έκδοση Άδειας Εκπαιδευτικού")
st.subheader("Συμπληρώστε τα στοιχεία για την παραγωγή του εγγράφου")

# Φόρμα Εισαγωγής Στοιχείων
with st.form("leave_form"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        eponymo = st.text_input("Επώνυμο", value=teacher_data["eponymo"])
        onoma = st.text_input("Όνομα ", value=teacher_data["onoma"])
        klados = st.text_input("Κλάδος ", value=teacher_data["klados"])
        mitrwo = st.text_input("ΑΜ ", value=teacher_data["mitrwo"])
        protocollo_aithshs = st.text_input("Αριθμός Πρωτοκόλλου Αίτησης")
        hmer_protocollou_aithshs = st.date_input("Ημερομηνία Πρωτοκόλλου Αίτησης")
        
    with col2:     
        protocollo_adeias = st.text_input("Πρωτόκολλο Άδειας")
        hmer_protocollou = st.date_input("Ημερομηνία Πρωτοκόλλου Άδειας")
        typos_adeias = st.selectbox("Τύπος Άδειας", 
                                  ["Κανονική πολλών ημερών", 
                                   "Κανονική μίας ημέρας", 
                                   "Αναρρωτική με ΥΔ", 
                                   "Αναρρωτική με Ιατρική γνωμάτευση", 
                                   "Ασθενείας τέκνου"])
        
        doctor = st.text_input("Γιατρός")
    
    with col3:
        days_lektiko = st.text_input("Πόσες ημέρες")
        days_number = st.number_input("Αριθμός Ημερών", min_value=1, step=1)
        arxh = st.date_input("Ημερομηνία Έναρξης")
        telos = st.date_input("Ημερομηνία Λήξης")

    # Ανέβασμα του προτύπου Word (προαιρετικά αν δεν το έχεις στον ίδιο φάκελο)
    template_file = st.file_uploader("Ανεβάστε το πρότυπο Word (.docx)", type=["docx"])
    
    submitted = st.form_submit_button("Δημιουργία Εγγράφου")

if submitted:
    if not template_file:
        st.error("Παρακαλώ ανεβάστε το πρότυπο .docx αρχείο πρώτα.")
    else:
        try:
            # Δημιουργία αντικειμένου Word από το template
            doc = DocxTemplate(template_file)
            
            # Δεδομένα για το "γέμισμα" του εγγράφου
            context = {
                'eponymo': eponymo,
                'onoma': onoma,
                'hmer_protocollou': hmer_protocollou,
                'protocollo_adeias' : protocollo_adeias,
                'hmer_protocollou_aithshs' : hmer_protocollou_aithshs,
                'protocollo_aithshs' : protocollo_aithshs,
                'klados' : klados,
                'mitrwo' : mitrwo,
                'typos_adeias': typos_adeias,
                'days_number': days_number,
                'days_lektiko' : days_lektiko,
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
            'Τύπος Άδειας': typos_adeias,
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