import streamlit as st
import pandas as pd
from datetime import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Control Stoc real Arcaprod", page_icon="🏗️", layout="wide")

# Configurații principale
SHEET_NAME = "Gestiune Stoc Arcaprod"
EMAIL_DESTINATAR = "sageata14@gmail.com"
PAROLA_ADMIN = "Arcaprod2026"

ANGAJATI = [
    "Andrei Barbuceanu", "Catalin", "George B", "Ionut R", "Cornel I", 
    "Marian D (Vampirul)", "Dumitru U (Mitel)", "Andrei D (Alifie)", 
    "Eugen T", "Ion T (Nelu)", "Marius V", "Gabriel V", "Iulian G", 
    "Geovani B", "Constantin U (Costi)", "Mircea R", "Aashish", 
    "Andrei B", "Ionut B", "Valentin"
]
TIP_MATERIAL = ["Profil Aluminiu", "Feronerie", "Accesorii", "Consumabile", "Sticlă"]
UNITATI_MASURA = ["cutii", "buc", "bax", "ml", "set", "bare"]

COLOANE_CENTRALIZATOR = ["Timestamp", "Angajat", "Tip Material", "Denumire Material", "Stoc Initial", "Cantitate Ceruta", "Stoc Ramas", "UM", "Status Stoc", "Observatii"]
COLOANE_STOCURI = ["Denumire Material", "Stoc Actual", "UM"]

# --- CONEXIUNE SECURIZATĂ GOOGLE SHEETS ---
@st.cache_resource
def conecteaza_gsheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def citeste_date_gsheet(nume_foaie, coloane_implicite):
    try:
        client = conecteaza_gsheet()
        sheet = client.open(SHEET_NAME).worksheet(nume_foaie)
        date = sheet.get_all_records()
        if not date:
            return pd.DataFrame(columns=coloane_implicite)
        return pd.DataFrame(date)
    except Exception as e:
        st.error(f"Eroare conectare Google Sheets (Foaia {nume_foaie}): {e}")
        return pd.DataFrame(columns=coloane_implicite)

def salveaza_stoc_gsheet(df_stoc):
    try:
        client = conecteaza_gsheet()
        sheet = client.open(SHEET_NAME).worksheet("Stocuri")
        sheet.clear()
        date_lista = [df_stoc.columns.values.tolist()] + df_stoc.values.tolist()
        sheet.update(range_name='A1', values=date_lista)
        return True
    except Exception as e:
        st.error(f"Eroare la salvarea stocului în cloud: {e}")
        return False

def adauga_in_centralizator_gsheet(rand_nou):
    try:
        client = conecteaza_gsheet()
        sheet = client.open(SHEET_NAME).worksheet("Centralizator")
        sheet.append_row(list(rand_nou.values()))
        return True
    except Exception as e:
        st.error(f"Eroare la scrierea istoricului în cloud: {e}")
        return False

# --- GENERARE EXCEL STILIZAT PENTRU EMAIL ---
def genereaza_excel_memorie(df_date, titlu_raport):
    wb = Workbook()
    ws = wb.active
    ws.title = "Date_Arcaprod"
    ws.views.sheetView[0].showGridLines = True
    
    ws["A1"] = titlu_raport.upper()
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F497D")
    ws["A2"] = f"Generat la data: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    
    headers = list(df_date.columns)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.value = header
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    subtire = Side(border_style="thin", color="D9D9D9")
    contur_celula = Border(left=subtire, right=subtire, top=subtire, bottom=subtire)
    
    for row_num, row_data in enumerate(df_date.values, 5):
        for col_num, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.value = str(val) if headers[col_num-1] == "Timestamp" else val
            cell.font = Font(name="Arial", size=10)
            cell.border = contur_celula
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2, 3]: continue
            if cell.value: max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    nume_fisier = f"{titlu_raport.lower().replace(' ', '_')}.xlsx"
    wb.save(nume_fisier)
    return nume_fisier

def trimite_email_cu_atansamente(fisiere):
    EMAIL_EXPEDIATOR = "adresa_ta_de_gmail@gmail.com"
    PAROLA_APLICATIE = "parola_ta_de_aplicatie_gmail" 
    
    if EMAIL_EXPEDIATOR == "adresa_ta_de_gmail@gmail.com":
        st.error("⚠️ Configurează adresa de email în cod pentru funcția de expediere!")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_EXPEDIATOR
    msg['To'] = EMAIL_DESTINATAR
    msg['Subject'] = f"📊 Raport Stocuri & Cereri Arcaprod - {datetime.now().strftime('%d-%m-%Y')}"
    msg.attach(MIMEText("Bună ziua,\n\nAtasat găsiți rapoartele descarcate din baza de date permanentă Google Sheets.\n\nO zi bună!", 'plain'))
    
    for cale_fisier in fisiere:
        if os.path.exists(cale_fisier):
            with open(cale_fisier, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(cale_fisier)}")
                msg.attach(part)
                
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_EXPEDIATOR, PAROLA_APLICATIE)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Eroare trimitere email: {e}")
        return False

# --- PANOU ADMIN SECURIZAT ---
st.sidebar.title("🔐 Panou Administrare Admin")
parola_introdusa = st.sidebar.text_input("Introdu parola Admin:", type="password")

if parola_introdusa == PAROLA_ADMIN:
    st.sidebar.success("🔑 Autentificat ca Admin!")
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Încarcă / Re-actualizează Stocul")
    fisiere_stoc_incarcat = st.sidebar.file_uploader("Încarcă fișierul de stoc existent (XLSX sau CSV):", type=["xlsx", "csv"])
    
    if fisiere_stoc_incarcat is not None:
        try:
            if fisiere_stoc_incarcat.name.endswith('.csv'):
                df_nou = pd.read_csv(fisiere_stoc_incarcat)
            else:
                try:
                    df_nou = pd.read_excel(fisiere_stoc_incarcat, skiprows=3)
                    if 'Denumire Material' not in df_nou.columns:
                        df_nou = pd.read_excel(fisiere_stoc_incarcat)
                except:
                    df_nou = pd.read_excel(fisiere_stoc_incarcat)
            
            df_nou = df_nou.loc[:, ~df_nou.columns.str.contains('^Unnamed')]
            
            if 'Denumire Material' in df_nou.columns and 'Stoc Actual' in df_nou.columns:
                df_nou['Denumire Material'] = df_nou['Denumire Material'].astype(str).str.strip().str.upper()
                df_nou['Stoc Actual'] = pd.to_numeric(df_nou['Stoc Actual']).fillna(0).astype(int)
                if 'UM' not in df_nou.columns: df_nou['UM'] = 'buc'
                
                df_stoc_salvare = df_nou[COLOANE_STOCURI]
                
                if st.sidebar.button("💾 Încarcă noul stoc în Google Sheets", use_container_width=True):
                    if salveaza_stoc_gsheet(df_stoc_salvare):
                        st.sidebar.success("🎯 Noul stoc a fost salvat și sincronizat în Google Sheets!")
                        st.rerun()
            else:
                st.sidebar.error("❌ Coloanele obligatorii în tabel trebuie să fie: 'Denumire Material' și 'Stoc Actual'")
        except Exception as e:
            st.sidebar.error(f"Eroare procesare fișier stoc: {e}")
elif parola_introdusa != "":
    st.sidebar.error("❌ Parolă incorectă!")

# --- INTERFAȚA PRINCIPALĂ ANGAJAȚI ---
st.title("🏗️ Arcaprod - Management & Control Stoc Real")

col_stanga, col_dreapta = st.columns([1, 1.3])

with col_stanga:
    st.subheader("📝 Formular Cerere Materiale")
    with st.form(key="formular_necesar", clear_on_submit=True):
        angajat = st.selectbox("Selectați Angajat:", ANGAJATI)
        tip = st.selectbox("Tip Material:", TIP_MATERIAL)
        denumire = st.text_input("Denumire Material (Ex: PANZA DEBITAT):").strip().upper()
        
        c1, c2 = st.columns(2)
        with c1:
            cantitate = st.number_input("Cantitate Solicitată:", min_value=1, step=1, format="%d")
        with c2:
            um = st.selectbox("UM:", UNITATI_MASURA)
            
        observatii = st.text_area("Observații adiționale:")
        buton_trimite = st.form_submit_button(label="Procesează și Verifică Cererea")

    if buton_trimite and denumire:
        df_stocuri = citeste_date_gsheet("Stocuri", COLOANE_STOCURI)
        if not df_stocuri.empty:
            df_stocuri['Denumire Material'] = df_stocuri['Denumire Material'].astype(str).str.strip().str.upper()
        
        stoc_initial = 0
        stoc_ramas = 0
        status_stoc = "stoc 0"
        
        if not df_stocuri.empty and denumire in df_stocuri["Denumire Material"].values:
            stoc_initial = int(df_stocuri.loc[df_stocuri["Denumire Material"] == denumire, "Stoc Actual"].values[0])
            
            if stoc_initial >= int(cantitate):
                stoc_ramas = stoc_initial - int(cantitate)
                status_stoc = "Existent in STOC"
                df_stocuri.loc[df_stocuri["Denumire Material"] == denumire, "Stoc Actual"] = stoc_ramas
            elif stoc_initial > 0:
                lipsa = int(cantitate) - stoc_initial
                stoc_ramas = 0
                status_stoc = f"stoc 0 (S-au eliberat doar {stoc_initial} {um}, rest {lipsa} de comandat urgent!)"
                df_stocuri.loc[df_stocuri["Denumire Material"] == denumire, "Stoc Actual"] = 0
            else:
                stoc_ramas = 0
                status_stoc = "stoc 0"
                df_stocuri.loc[df_stocuri["Denumire Material"] == denumire, "Stoc Actual"] = 0
        else:
            nou_rand_stoc = pd.DataFrame([{"Denumire Material": denumire, "Stoc Actual": 0, "UM": um}])
            df_stocuri = pd.concat([df_stocuri, nou_rand_stoc], ignore_index=True)
            stoc_initial = 0
            stoc_ramas = 0
            status_stoc = "stoc 0"
            
        salveaza_stoc_gsheet(df_stocuri)
        
        rand_nou_centralizator = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Angajat": angajat,
            "Tip Material": tip,
            "Denumire Material": denumire,
            "Stoc Initial": stoc_initial,
            "Cantitate Ceruta": int(cantitate),
            "Stoc Ramas": stoc_ramas,
            "UM": um,
            "Status Stoc": status_stoc,
            "Observatii": observatii
        }
        
        if adauga_in_centralizator_gsheet(rand_nou_centralizator):
            if status_stoc == "Existent in STOC":
                st.success(f"✅ Material aprobat complet. Status: {status_stoc}. (Stoc rămas în cloud: {stoc_ramas} {um})")
            else:
                st.error(f"🚨 {status_stoc}")

with col_dreapta:
    df_hist = citeste_date_gsheet("Centralizator", COLOANE_CENTRALIZATOR)
    df_stocuri = citeste_date_gsheet("Stocuri", COLOANE_STOCURI)
    
    st.subheader("📧 Expediere Rapoarte")
    if st.button("🚀 Trimite fișierele Excel către sageata14@gmail.com", use_container_width=True):
        with st.spinner("Se descarcă datele din Google Sheets și se trimit..."):
            f1 = genereaza_excel_memorie(df_hist, "Registru Istoric Comenzi Materiale")
            f2 = genereaza_excel_memorie(df_stocuri, "Baza de Date Stocuri Existente")
            if trimite_email_cu_atansamente([f1, f2]):
                st.success(f"📩 Rapoartele Excel au fost livrate cu succes la {EMAIL_DESTINATAR}!")
            if os.path.exists(f1): os.remove(f1)
            if os.path.exists(f2): os.remove(f2)

    st.markdown("---")
    st.subheader("📊 Inventar Depozit în Timp Real (Google Cloud)")
    if not df_stocuri.empty:
        df_stocuri["Stoc Actual"] = pd.to_numeric(df_stocuri["Stoc Actual"]).fillna(0).astype(int)
        df_afisare_stoc = df_stocuri.copy()
        df_afisare_stoc["Status Material"] = df_afisare_stoc["Stoc Actual"].apply(lambda x: "✅ Disponibil în Depozit" if x > 0 else "🚨 STOC 0")
        st.dataframe(df_afisare_stoc, use_container_width=True, hide_index=True)
    else:
        st.info("Baza de date din Google Sheets este goală sau noul stoc nu a fost încărcat din Panoul Admin.")

    st.markdown("---")
    st.subheader("📋 Istoric Cereri și Mișcări de Stoc")
    if not df_hist.empty:
        if datetime.now().hour >= 15:
            st.warning("🔒 După ora 15:00, registrul de modificări zilnice este securizat.")
        st.dataframe(df_hist, use_container_width=True, hide_index=True)
    else:
        st.info("Niciun istoric înregistrat în Google Sheet.")
