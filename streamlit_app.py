# streamlit_insurance_letter.py
"""
Streamlit app to generate a summary letter for insurance contract movements.

Key features now supported
--------------------------
▶ Multiple input fields (codice fiscale, calculation date) that flow into the
  subject and intro lines.
▶ Translation map extended with **table id** → items that share a table id are
  aggregated into the *same* table.
▶ Tables can opt‑in/out of an automatic **Total** row (`include_total`).
▶ Amounts formatted as Italian‑locale Euro values (e.g. 24.300,45 €).
"""

# ---- Imports --------------------------------------------------------------
import locale
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from babel.numbers import format_currency
from docx import Document

# ---------------------------------------------------------------------------
#  ITALIAN LOCALE FOR CURRENCY FORMATTING
# ---------------------------------------------------------------------------
try:
    locale.setlocale(locale.LC_ALL, "it_IT.utf8")
except locale.Error:
    # fallback so babel still works; streamlit cloud may not have the locale installed
    pass


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  HARD‑CODED CONFIGURATION                                            ║
# ╚══════════════════════════════════════════════════════════════════════╝

# Per‑item configuration: label (ITA) + table id
ITEM_CONFIG = {
    "Acquisition cost deduction from regular premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
    },
    "Contract fee deduction from regular premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
    },
    "Acquisition cost deduction from single premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
    },
    "Contract fee deduction from single premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
    },

    "Administrative deduction": {
        "label": "Costi di caricamento",
        "table": "T2",
    },

    "Investment deduction": {
        "label": "Costi di investimento",
        "table": "T3",
    },
    "Investment deduction from Regular Premium Balance": {
        "label": "Costi di investimento",
        "table": "T3",
    },
    "Investment deduction from Single PremiumBalance": {
        "label": "Costi di investimento",
        "table": "T3",
    },

    "Risk deduction - Death": {
        "label": "Trattenuta copertura rischio morte",
        "table": "T4",
    },
    "Risk deduction - Waiver of premium": {
        "label": "Esonero Pagamento Premi ITP",
        "table": "T4",
    },
    "Risk deduction - Illnesses and operations": {
        "label": "Trattenuta rischio malattia / interventi",
        "table": "T4",
    },
    "Risk deduction - accident insurance deduction": {
        "label": "Trattenuta copertura rischio infortunio",
        "table": "T4",
    },

    "Investment return of Novis Loyalty Bonus": {
        "label": "Rendimento Bonus Fedeltà NOVIS",
        "table": "T5",
    },
    "Investment return from insurance funds": {
        "label": "Capitalizzazione",
        "table": "T5",
    },
    "NOVIS Special Bonus": {
        "label": "NOVIS Special Bonus",
        "table": "T5",
    },

    "Paid Premium": {
        "label": "Pagamenti dei Premi identificati",
        "table": "T6",
    },
}

# Per‑table configuration: title + whether a Total row is appended + its label
TABLE_CONFIG = {
    "T1": {"title": "Costi di emissione e gestione", "include_total": True, "total_label": "Totale costi di emissione e gestione"},
    "T2": {"title": "Costi di caricamento", "include_total": True, "total_label": "Totale costi di caricamento"},
    "T3": {"title": "Costi di investimento", "include_total": True, "total_label": "Totale costi di investimento"},
    "T4": {"title": "Trattenute di rischio", "include_total": True, "total_label": "Totale trattenute di rischio"},
    "T5": {"title": "Rendimenti / Bonus", "include_total": False},
    "T6": {"title": "Premi versati", "include_total": False},
}

# ── Letter text blocks ─────────────────────────────────────────────────
LETTER_SUBJECT_TPL = (
    "Dettaglio costi per il valore della Sua posizione assicurativa polizza n. "
    "{contract_number} al {calc_date} con codice fiscale {cf}."
)

LETTER_BODY_HEADER_TPL = (
    "Egregio/a {client_name},\n\n"
    "siamo con la presente a trasmetterLe di seguito la tabella riportante il "
    "dettaglio dei costi applicati ai fini di calcolo del valore della Sua "
    "posizione assicurativa al {calc_date}."
)

OUTRO_PARAGRAPH = (
    "Qualora necessitasse di ulteriori informazioni in merito, La invitiamo "
    "gentilmente a riferirsi alla Tabella Costi contenuta nelle Condizioni di "
    "Assicurazione.\n\n"
    "Rimaniamo a disposizione per qualsiasi chiarimento e, ringraziando per la "
    "cortese attenzione, Le porgiamo i nostri più cordiali saluti."
)

GOODBYE_LINE = "Cordiali saluti,"

SIGNATURE_BLOCK = (
    "Il team NOVIS\n\n"
    "NOVIS Insurance Company,\n"
    "NOVIS Versicherungsgesellschaft,\n"
    "NOVIS Compagnia di Assicurazioni,\n"
    "NOVIS Poisťovňa a.s."
)

# ---------------------------------------------------------------------------
#  COLUMN HANDLING FOR NOVIS EXPORT
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "EntryDate": "Item date",
    "ValueDate": "Item date",
    "EntryType": "Item name",
    "Amount": "Item value",
}
EXPECTED_COLS = {"Item date", "Item name", "Item value"}


def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    provided = set(df.columns)
    if EXPECTED_COLS.issubset(provided):
        return df
    df = df.rename({k: v for k, v in COLUMN_ALIASES.items() if k in df.columns}, axis=1)
    provided = set(df.columns)
    missing = EXPECTED_COLS - provided
    if missing:
        raise ValueError(f"File privo delle colonne richieste: {', '.join(missing)}")
    return df

# ---------------------------------------------------------------------------
#  DATA AGGREGATION / TABLE SPLIT                                            
# ---------------------------------------------------------------------------

def aggregate_by_table(df: pd.DataFrame):
    """Return dict {table_id: DataFrame} already summed & labelled."""
    df["Item value"] = pd.to_numeric(df["Item value"], errors="coerce")
    df = df.dropna(subset=["Item value"])

    # Map to label & table id
    df["Label"] = df["Item name"].map(lambda x: ITEM_CONFIG.get(x, {}).get("label", x))
    df["Table"] = df["Item name"].map(lambda x: ITEM_CONFIG.get(x, {}).get("table", "ALT"))

    grouped = (
        df.groupby(["Table", "Label"], as_index=False)["Item value"].sum()
        .rename(columns={"Item value": "Amount"})
    )

    # Split per table id
    tables: dict[str, pd.DataFrame] = {}
    for tbl_id, sub in grouped.groupby("Table"):
        # order rows by label for readability
        tables[tbl_id] = sub.drop(columns="Table").sort_values("Label")
    return tables

# ---------------------------------------------------------------------------
#  LETTER BUILDERS                                                           
# ---------------------------------------------------------------------------

def _fmt(amount: float) -> str:
    return format_currency(amount, "EUR", locale="it_IT")


def build_letter_doc(
    client_name: str,
    client_address: str,
    contract_number: str,
    codice_fiscale: str,
    calc_date: str,
    tables: dict[str, pd.DataFrame],
) -> Document:
    doc = Document()

    # heading
    for p in (client_name, client_address, "", date.today().strftime("%d/%m/%Y"), ""):
        doc.add_paragraph(p)

    doc.add_paragraph(
        LETTER_SUBJECT_TPL.format(contract_number=contract_number, calc_date=calc_date, cf=codice_fiscale)
    ).style = "Heading 2"
    doc.add_paragraph("")

    doc.add_paragraph(
        LETTER_BODY_HEADER_TPL.format(client_name=client_name, calc_date=calc_date)
    )

    # iterate over logical tables
    for tbl_id, df in tables.items():
        cfg = TABLE_CONFIG.get(tbl_id, {"title": tbl_id, "include_total": False})
        doc.add_paragraph(cfg["title"]).style = "Heading 3"
        t = doc.add_table(rows=1, cols=2)
        t.rows[0].cells[0].text = "Item"
        t.rows[0].cells[1].text = "Importo"
        for _, row in df.iterrows():
            c1, c2 = t.add_row().cells
            c1.text = row["Label"]
            c2.text = _fmt(row["Amount"])
        if cfg.get("include_total"):
            tot = df["Amount"].sum()
            c1, c2 = t.add_row().cells
            c1.text = cfg.get("total_label", "Totale")
            c2.text = _fmt(tot)
        doc.add_paragraph("")

    # outro
    doc.add_paragraph(OUTRO_PARAGRAPH)
    doc.add_paragraph("")
    doc.add_paragraph(GOODBYE_LINE)
    doc.add_paragraph("")
    doc.add_paragraph(SIGNATURE_BLOCK)
    return doc


def doc_to_bytes(doc: Document) -> BytesIO:
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
#  STREAMLIT APP                                                             
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Insurance Letter Generator", layout="centered")
    st.title("📄 Generatore Lettera Valorizzazione")

    file = st.file_uploader("Carica file XLS/XLSX movimenti", type=["xls", "xlsx"])

    st.subheader("Dati cliente")
    client_name = st.text_input("Nome del cliente")
    contract_number = st.text_input("Numero polizza")
    client_address = st.text_area("Indirizzo cliente")
    codice_fiscale = st.text_input("Codice fiscale")
    calc_date = st.date_input("Data valorizzazione", value=date.today()).strftime("%d/%m/%Y")

    if file is not None:
        try:
            df_raw = pd.read_excel(file)
            df = standardise_columns(df_raw)
        except Exception as e:
            st.error(f"Errore lettura file: {e}")
            st.stop()

        tables = aggregate_by_table(df)

        st.subheader("Anteprima tabelle")
        for tbl_id, df_tbl in tables.items():
            cfg = TABLE_CONFIG.get(tbl_id, {"title": tbl_id})
            st.markdown(f"### {cfg['title']}")
            st.dataframe(
                df_tbl.assign(Importo=df_tbl["Amount"].apply(_fmt)).drop(columns="Amount"),
                use_container_width=True,
            )
            if cfg.get("include_total"):
                st.markdown(f"**{cfg.get('total_label', 'Totale')}: {_fmt(df_tbl['Amount'].sum())}**")

        if all([client_name, contract_number, client_address, codice_fiscale]):
            doc = build_letter_doc(
                client_name,
                client_address,
                contract_number,
                codice_fiscale,
                calc_date,
                tables,
            )
            st.download_button(
                "⬇️ Scarica Word",
                data=doc_to_bytes(doc),
                file_name=f"Valorizzazione_dettagliata_polizza_{contract_number}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.info("Compila tutti i campi del cliente per poter generare la lettera.")


if __name__ == "__main__":
    main()
