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

st.set_page_config(page_title="RISE Report (Online)", layout="wide")

WEEKDAYS_ABBR = ("Lu","Ma","Me","Gi","Ve","Sa","Do")
WEEKDAYS_MAP = {
    "Lu": "Lunedì", "Ma": "Martedì", "Me": "Mercoledì",
    "Gi": "Giovedì", "Ve": "Venerdì", "Sa": "Sabato", "Do": "Domenica"
}
RE_ROW_DATE = re.compile(r'^\\s*(\\d{2})\\s+(' + "|".join(WEEKDAYS_ABBR) + r')\\b')
RE_HHMM = re.compile(r'(\\d{2}:\\d{2})')
RE_FILE_NAME = re.compile(r'Cartellino_(\\d{2})_(\\d{4})\\.pdf', re.IGNORECASE)

def parse_month_year_from_name(filename: str):
    m = RE_FILE_NAME.search(filename)
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))

def walk_pdfs(folder):
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf") and not f.startswith("._"):
                yield os.path.join(root, f)

def analyze_folder(folder):
    records = []
    year_days = defaultdict(set)

    for pdf_path in sorted(list(walk_pdfs(folder))):
        fname = os.path.basename(pdf_path)
        mese, anno = parse_month_year_from_name(fname)
        if not mese or not anno:
            continue
        try:
            with fitz.open(pdf_path) as doc:
                for page in doc:
                    lines = page.get_text("text").splitlines()
                    block = []
                    current_date = None
                    current_dow = None
                    for line in lines + ["__ENDOFPAGE__"]:
                        m = RE_ROW_DATE.match(line)
                        if m or line == "__ENDOFPAGE__":
                            if block and current_date is not None:
                                block_text = "\\n".join(block)
                                if "RISE" in block_text:
                                    times = RE_HHMM.findall(block_text)
                                    if len(times) >= 2:
                                        pairs = list(zip(times[0::2], times[1::2]))
                                        year_days[anno].add(current_date.date())
                                        for idx, (ing, usc) in enumerate(pairs, 1):
                                            records.append({
                                                "Anno": anno, "Mese": mese,
                                                "Data": current_date.date().isoformat(),
                                                "Giorno": WEEKDAYS_MAP.get(current_dow, current_dow),
                                                "Entrata": ing, "Uscita": usc, "Coppia": idx,
                                                "File": fname
                                            })
                            block = []
                            if m:
                                day_num = int(m.group(1))
                                dow = m.group(2)
                                current_dow = dow
                                try:
                                    current_date = datetime(anno, mese, day_num)
                                except:
                                    current_date = None
                            else:
                                current_date = None
                                current_dow = None
                        else:
                            block.append(line)
        except Exception as e:
            st.warning(f"Impossibile leggere {fname}: {e}")
    return pd.DataFrame(records), year_days

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

def export_excel_memory(records_df, year_days_map) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df = records_df.copy()
        if not df.empty:
            df["Data_dt"] = pd.to_datetime(df["Data"])
            df = df.sort_values(["Anno", "Data_dt", "Coppia"]).drop(columns=["Data_dt"])
        df.to_excel(writer, index=False, sheet_name="Dettaglio")

        rows = [{"Anno": y, "Giorni con RISE": len(year_days_map[y])} for y in sorted(year_days_map)]
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Riepilogo")
    return buf.getvalue()

def build_months_zip_memory(records_df, extracted_root) -> bytes:
    # ZIP con i soli PDF dei mesi che hanno RISE
    buf = io.BytesIO()
    names = set(records_df["File"].unique()) if not records_df.empty else set()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(extracted_root):
            for f in files:
                if f in names and f.lower().endswith(".pdf") and not f.startswith("._"):
                    full = os.path.join(root, f)
                    z.write(full, arcname=f)
    return buf.getvalue()

st.title("RISE Report — Generatore 100% Online")
st.caption("Carica lo ZIP dei cartellini (Cartellino_MM_YYYY.pdf), inserisci il nome, poi scarica i report.")

with st.form("rise-form"):
    full_name = st.text_input("Nome e cognome", value="")
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
            df, year_days = analyze_folder(extracted)
            total_days = sum(len(v) for v in year_days.values())

        st.success(f"Analisi completata. Totale giornate con RISE: {total_days}")
        # Mostra riepilogo
        rows = [{"Anno": y, "Giorni con RISE": len(year_days[y])} for y in sorted(year_days)]
        st.subheader("Riepilogo per anno")
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

        # Mostra dettaglio
        st.subheader("Dettaglio (tutte le righe con RISE)")
        if not df.empty:
            df["Data_dt"] = pd.to_datetime(df["Data"])
            df_sorted = df.sort_values(["Anno", "Data_dt", "Coppia"]).drop(columns=["Data_dt"])
        else:
            df_sorted = df
        st.dataframe(df_sorted, use_container_width=True)

        # Genera output in memoria
        pdf_bytes = export_pdf_memory(df, full_name, total_days)
        xlsx_bytes = export_excel_memory(df, year_days)
        zip_bytes = build_months_zip_memory(df, extracted)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "⬇️ Scarica PDF",
                data=pdf_bytes,
                file_name=f"Report_RISE_{full_name.replace(' ','_')}.pdf",
                mime="application/pdf"
            )
        with col2:
            st.download_button(
                "⬇️ Scarica Excel",
                data=xlsx_bytes,
                file_name=f"Report_RISE_{full_name.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col3:
            st.download_button(
                "⬇️ Scarica ZIP mesi con RISE",
                data=zip_bytes,
                file_name=f"Cartellini_RISE_mesi_{full_name.replace(' ','_')}.zip",
                mime="application/zip"
            )

st.markdown("---")
st.caption("Suggerimenti: se lo ZIP è molto grande, suddividilo per anno. I file macOS '__MACOSX/._...' vengono ignorati automaticamente.")
