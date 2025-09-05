import streamlit as st
import os, re, zipfile, tempfile, io
from datetime import datetime
from collections import defaultdict
import pandas as pd
import fitz  # PyMuPDF

# PDF export
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="RISE Report — Upload libero", layout="wide")

# -----------------------
# Regex & helpers
# -----------------------
WEEKDAYS_ABBR = ("Lu","Ma","Me","Gi","Ve","Sa","Do")
WEEKDAYS_FULL = ("Lun","Mar","Mer","Gio","Ven","Sab","Dom")
WEEKDAYS_MAP = {
    "Lu": "Lunedì", "Ma": "Martedì", "Me": "Mercoledì",
    "Gi": "Giovedì", "Ve": "Venerdì", "Sa": "Sabato", "Do": "Domenica",
    "Lun": "Lunedì", "Mar": "Martedì", "Mer": "Mercoledì",
    "Gio": "Giovedì", "Ven": "Venerdì", "Sab": "Sabato", "Dom": "Domenica"
}

# RIGA DATA: tollerante a abbreviazioni (Lu/Lun ecc.), spazi variabili
RE_ROW_DATE = re.compile(r'^\s*(\d{1,2})\s+(' + "|".join(list(WEEKDAYS_ABBR)+list(WEEKDAYS_FULL)) + r')\b')
# ORARI HH:MM
RE_HHMM = re.compile(r'(\d{2}:\d{2})')
# "RISE" tollerante a spazi/maiuscole (es. R I S E)
RE_RISE = re.compile(r'\bR\s*I\s*S\s*E\b', re.IGNORECASE)

# Mesi italiani per parsing e visualizzazione
MESI = {
    "gennaio":1,"febbraio":2,"marzo":3,"aprile":4,"maggio":5,"giugno":6,
    "luglio":7,"agosto":8,"settembre":9,"ottobre":10,"novembre":11,"dicembre":12
}
MESE_NOME = {v: k.capitalize() for k, v in MESI.items()}

def parse_month_year_from_text(text: str):
    # Cerca MM/YYYY, MM-YYYY, YYYY-MM, Mese YYYY
    m = re.search(r'\b(\d{1,2})\s*[/\-]\s*(\d{4})\b', text)
    if m:
        mm = int(m.group(1)); yy = int(m.group(2))
        if 1 <= mm <= 12:
            return mm, yy
    m2 = re.search(r'\b(\d{4})\s*[-/]\s*(\d{1,2})\b', text)
    if m2:
        yy = int(m2.group(1)); mm = int(m2.group(2))
        if 1 <= mm <= 12:
            return mm, yy
    m3 = re.search(r'(' + "|".join(MESI.keys()) + r')\s+(\d{4})', text, re.IGNORECASE)
    if m3:
        mm = MESI[m3.group(1).lower()]; yy = int(m3.group(2))
        return mm, yy
    return None, None

def walk_pdfs(folder):
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf") and not f.startswith("._"):
                yield os.path.join(root, f)

def analyze_folder(folder):
    records = []
    year_days = defaultdict(set)
    diagnostics = []

    pdfs = sorted(list(walk_pdfs(folder)))
    if not pdfs:
        return pd.DataFrame(records), year_days, diagnostics

    for pdf_path in pdfs:
        fname = os.path.basename(pdf_path)
        month, year = None, None

        try:
            with fitz.open(pdf_path) as doc:
                # prova a recuperare mese/anno dal contenuto
                text0 = doc[0].get_text("text")
                month, year = parse_month_year_from_text(text0)
                if not month or not year:
                    concat = []
                    for p in doc[:min(3, len(doc))]:
                        concat.append(p.get_text("text"))
                    month, year = parse_month_year_from_text("\n".join(concat))

                total_rise_this_pdf = 0

                for page in doc:
                    lines = page.get_text("text").splitlines()
                    block = []
                    current_date = None
                    current_dow = None

                    for line in lines + ["__ENDOFPAGE__"]:
                        m = RE_ROW_DATE.match(line)
                        if m or line == "__ENDOFPAGE__":
                            if block and (month and year):
                                block_text = "\n".join(block)
                                if RE_RISE.search(block_text):
                                    total_rise_this_pdf += 1
                                    times = RE_HHMM.findall(block_text)
                                    if len(times) >= 2 and current_date is not None:
                                        pairs = list(zip(times[0::2], times[1::2]))
                                        year_days[year].add(current_date.date())
                                        for idx, (ing, usc) in enumerate(pairs, 1):
                                            records.append({
                                                "Anno": year, "Mese": month,
                                                "Data": current_date.date().isoformat(),
                                                "Giorno": WEEKDAYS_MAP.get(current_dow, current_dow),
                                                "Entrata": ing, "Uscita": usc, "Coppia": idx,
                                                "File": fname
                                            })
                            block = []
                            if m and (month and year):
                                day_num = int(m.group(1))
                                dow = m.group(2)
                                current_dow = dow
                                try:
                                    current_date = datetime(year, month, day_num)
                                except:
                                    current_date = None
                            else:
                                current_date = None
                                current_dow = None
                        else:
                            block.append(line)

                diagnostics.append({
                    "file": fname,
                    "mese": month,
                    "anno": year,
                    "rise_blocchi_rilevati": total_rise_this_pdf
                })

        except Exception as e:
            st.warning(f"Impossibile leggere {fname}: {e}")
            diagnostics.append({"file": fname, "errore": str(e)})

    return pd.DataFrame(records), year_days, diagnostics

def export_pdf_memory(records_df, full_name, total_days) -> bytes:
    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    title = styles["Heading1"]
    doc = SimpleDocTemplate(buf, pagesize=A4)
    elements = []
    elements.append(Paragraph(f"{full_name} - Totale giorni RISE: {total_days}", title))
    elements.append(Spacer(1, 10))

    if records_df.empty:
        elements.append(Paragraph("Nessuna giornata RISE trovata con orari di ingresso/uscita.", styles["Normal"]))
        doc.build(elements)
        return buf.getvalue()

    df = records_df.copy()
    df["Data_dt"] = pd.to_datetime(df["Data"])
    df = df.sort_values(["Anno", "Data_dt", "Coppia"])

    data_table = [["Anno", "Data", "Giorno", "Entrata", "Uscita"]]
    for _, row in df.iterrows():
        data_table.append([str(row["Anno"]), row["Data"], row["Giorno"], row["Entrata"], row["Uscita"]])

    tab = Table(data_table, repeatRows=1)
    tab.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black)
    ]))
    elements.append(tab)
    doc.build(elements)
    return buf.getvalue()

# -----------------------
# UI
# -----------------------
st.title("RISE Report — Upload libero (senza regole sul nome file)")
st.caption("Carica uno ZIP con i cartellini: l'app leggerà mese/anno dal contenuto dei PDF.")
st.info("🔒 **Privacy**: i file caricati vengono elaborati solo durante la sessione in una cartella temporanea e non vengono salvati in modo permanente sul server.")

# --- Guida: come creare correttamente lo ZIP dei cartellini ---
with st.expander("📦 Istruzioni per creare lo ZIP dei cartellini", expanded=False):
    st.markdown("""
### 📋 Procedura per creare lo ZIP dei cartellini

1. **Accedi al portale WEFS**  
   Vai nella sezione **Cartellino orologio**.

2. **Seleziona mese e anno**  
   - Devi recuperare i cartellini fino a **10 anni indietro** rispetto alla data di spedizione dell’interruttiva.  
   - *Esempio*: se l’interruttiva è datata **settembre 2025**, parti da **settembre 2015**.

3. **Scarica i PDF**  
   - Una volta caricato il mese scelto, in basso a destra troverai l’icona **PDF**.  
   - Cliccala per scaricare il file sul tuo computer.  
   - Ripeti l’operazione per tutti i mesi di tuo interesse.

4. **Crea una cartella con i PDF**  
   - Inserisci dentro la stessa cartella tutti i file PDF scaricati.  
   - Non è importante il nome dei singoli file.

5. **Comprimi la cartella in formato ZIP**  
   - **Windows**: clic destro → *Invia a → Cartella compressa*  
   - **macOS**: clic destro → *Comprimi*  
   - Il file ZIP può avere **qualsiasi nome**.

6. **Carica lo ZIP nell’app**  
   - Inserisci il tuo **nome e cognome** nel campo dedicato.  
   - Carica lo ZIP.  
   - Premi su **Genera Report**.
""")

with st.form("rise-form"):
    full_name = st.text_input("Nome e cognome (per intestazione PDF)", value="")
    uploaded = st.file_uploader("Archivio ZIP cartellini", type=["zip"])
    run = st.form_submit_button("Genera report")

if run:
    if not full_name.strip():
        st.error("Inserisci il nome e cognome.")
        st.stop()
    if not uploaded:
        st.error("Carica un archivio ZIP di cartellini.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "input.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded.getbuffer())

        extracted = os.path.join(tmpdir, "unzipped")
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extracted)

        with st.spinner("Analisi in corso..."):
            df, year_days, diags = analyze_folder(extracted)
            total_days = sum(len(v) for v in year_days.values())

        st.success(f"Analisi completata. Totale giornate con RISE: {total_days}")

        # -----------------------
        # Nuova tabella: mesi/anni con corrispondenze
        # -----------------------
        st.subheader("Scarica le buste paga e i cartellini orologio dei seguenti mesi.")
        if df.empty:
            st.info("Nessuna corrispondenza trovata.")
            months_df = pd.DataFrame(columns=["Anno", "Mese"])
        else:
            months_df = (
                df.groupby(["Anno", "Mese"])["Data"]
                  .nunique()
                  .reset_index(name="Giorni con RISE")
            )
            months_df["Mese"] = months_df["Mese"].map(MESE_NOME).fillna(months_df["Mese"])
            months_df = months_df.sort_values(["Anno", "Mese"], key=lambda s: s.map({n:i for i,n in enumerate(MESE_NOME.values(), start=1)}) if s.name=="Mese" else s)
        st.dataframe(months_df, use_container_width=True)

        # Diagnostica (facoltativa ma utile)
        st.subheader("Diagnostica")
        st.dataframe(pd.DataFrame(diags), use_container_width=True)

        # Dettaglio completo
        st.subheader("Dettaglio (tutte le righe con RISE)")
        if not df.empty:
            df["Data_dt"] = pd.to_datetime(df["Data"])
            df_sorted = df.sort_values(["Anno", "Data_dt", "Coppia"]).drop(columns=["Data_dt"])
        else:
            df_sorted = df
        st.dataframe(df_sorted, use_container_width=True)

        # Output in memoria: PDF
        pdf_bytes = export_pdf_memory(df, full_name, total_days)
        st.download_button(
            "⬇️ Scarica PDF",
            data=pdf_bytes,
            file_name=f"Report_RISE_{full_name.replace(' ','_')}.pdf",
            mime="application/pdf"
        )

st.markdown("---")
st.caption("Nota: se i PDF sono scansioni immagini senza testo, serve l'OCR per estrarre i dati.")

